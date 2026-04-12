param(
    [Parameter(Mandatory = $true)]
    [string]$DocxPath,

    [string]$PdfPath,

    [int]$TimeoutSeconds = 180,

    [int]$RetryCount = 2,

    [switch]$Worker,

    [string]$LogPath
)

function Remove-StaleWinWord {
    Get-Process WINWORD -ErrorAction SilentlyContinue |
        Where-Object { $_.MainWindowHandle -eq 0 } |
        ForEach-Object {
            try {
                Stop-Process -Id $_.Id -Force -ErrorAction Stop
            }
            catch {
            }
        }
}

function Invoke-ExportWorker {
    param(
        [string]$ResolvedDocx,
        [string]$ResolvedPdf,
        [string]$WorkerLogPath
    )

    $word = $null
    $document = $null
    try {
        $word = New-Object -ComObject Word.Application
        $word.Visible = $false
        $word.DisplayAlerts = 0
        $document = $word.Documents.Open($ResolvedDocx, $false, $true)
        $document.ExportAsFixedFormat($ResolvedPdf, 17)
        Write-Output $ResolvedPdf
        exit 0
    }
    catch {
        $message = @{
            step = "export_word_preview"
            error = $_.Exception.Message
            docx = $ResolvedDocx
            pdf = $ResolvedPdf
            recovery_hint = "Close busy Word documents and retry. If it still fails, open the docx manually and confirm it can be exported to PDF."
        } | ConvertTo-Json -Compress
        if ($WorkerLogPath) {
            Set-Content -LiteralPath $WorkerLogPath -Value $message -Encoding UTF8
        }
        Write-Error $message
        exit 1
    }
    finally {
        if ($document -ne $null) {
            $document.Close([ref]$false)
        }
        if ($word -ne $null) {
            $word.Quit()
        }
        if ($document -ne $null) {
            [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($document)
        }
        if ($word -ne $null) {
            [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($word)
        }
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }
}

$resolvedDocx = (Resolve-Path -LiteralPath $DocxPath).Path

if (-not $PdfPath) {
    $baseName = [System.IO.Path]::GetFileNameWithoutExtension($resolvedDocx)
    $dirName = [System.IO.Path]::GetDirectoryName($resolvedDocx)
    $PdfPath = Join-Path $dirName ($baseName + "_preview.pdf")
}

$pdfFullPath = [System.IO.Path]::GetFullPath($PdfPath)

if ($Worker.IsPresent) {
    Invoke-ExportWorker -ResolvedDocx $resolvedDocx -ResolvedPdf $pdfFullPath -WorkerLogPath $LogPath
    exit 0
}

$attemptErrors = @()
for ($attempt = 1; $attempt -le $RetryCount; $attempt++) {
    Remove-StaleWinWord
    $tempLog = Join-Path ([System.IO.Path]::GetTempPath()) ("word-export-" + [guid]::NewGuid().ToString() + ".log")
    $proc = Start-Process -FilePath "powershell.exe" -ArgumentList @(
        "-NoProfile",
        "-Sta",
        "-ExecutionPolicy", "Bypass",
        "-File", $PSCommandPath,
        "-DocxPath", $resolvedDocx,
        "-PdfPath", $pdfFullPath,
        "-Worker",
        "-LogPath", $tempLog
    ) -WindowStyle Hidden -PassThru

    if (-not $proc.WaitForExit($TimeoutSeconds * 1000)) {
        try {
            Stop-Process -Id $proc.Id -Force -ErrorAction Stop
        }
        catch {
        }
        Remove-StaleWinWord
        $attemptErrors += "attempt $attempt timed out after ${TimeoutSeconds}s"
        continue
    }

    if ($proc.ExitCode -eq 0 -and (Test-Path -LiteralPath $pdfFullPath)) {
        Write-Output $pdfFullPath
        exit 0
    }

    if (Test-Path -LiteralPath $tempLog) {
        $attemptErrors += (Get-Content -LiteralPath $tempLog -Raw)
    }
    else {
        $attemptErrors += "attempt $attempt failed without a worker log"
    }
    Remove-StaleWinWord
}

$errorPayload = @{
    step = "export_word_preview"
    docx = $resolvedDocx
    pdf = $pdfFullPath
    retry_count = $RetryCount
    timeout_seconds = $TimeoutSeconds
    attempts = $attemptErrors
    recovery_hint = "Close all background Word processes and retry. If it still fails, open the document manually to check whether it is damaged, or switch to a non-COM flow."
} | ConvertTo-Json -Depth 4

throw $errorPayload
