param(
    [Parameter(Mandatory = $true)]
    [string]$DocxPath,

    [string]$OutputPath,

    [switch]$AcceptRevisions,

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

function Invoke-FinalizeWorker {
    param(
        [string]$ResolvedDocx,
        [string]$ResolvedOutput,
        [bool]$ShouldAcceptRevisions,
        [string]$WorkerLogPath
    )

    Copy-Item -LiteralPath $ResolvedDocx -Destination $ResolvedOutput -Force

    $word = $null
    $document = $null
    try {
        $word = New-Object -ComObject Word.Application
        $word.Visible = $false
        $word.DisplayAlerts = 0
        $document = $word.Documents.Open($ResolvedOutput)

        foreach ($toc in $document.TablesOfContents) {
            $toc.Update()
        }
        $document.Fields.Update() | Out-Null

        if ($ShouldAcceptRevisions) {
            $document.AcceptAllRevisions()
        }

        for ($i = $document.Comments.Count; $i -ge 1; $i--) {
            $document.Comments.Item($i).Delete()
        }

        $document.Save()
        Write-Output $ResolvedOutput
        exit 0
    }
    catch {
        $message = @{
            step = "finalize_submission_copy"
            error = $_.Exception.Message
            docx = $ResolvedDocx
            output = $ResolvedOutput
            recovery_hint = "Close busy Word documents and retry. If the issue is only about comments, create a clean copy with strip_docx_comments.py first."
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

if (-not $OutputPath) {
    $baseName = [System.IO.Path]::GetFileNameWithoutExtension($resolvedDocx)
    $dirName = [System.IO.Path]::GetDirectoryName($resolvedDocx)
    $OutputPath = Join-Path $dirName ($baseName + "_clean.docx")
}

$outputFullPath = [System.IO.Path]::GetFullPath($OutputPath)

if ($Worker.IsPresent) {
    Invoke-FinalizeWorker -ResolvedDocx $resolvedDocx -ResolvedOutput $outputFullPath -ShouldAcceptRevisions $AcceptRevisions.IsPresent -WorkerLogPath $LogPath
    exit 0
}

$attemptErrors = @()
for ($attempt = 1; $attempt -le $RetryCount; $attempt++) {
    Remove-StaleWinWord
    $tempLog = Join-Path ([System.IO.Path]::GetTempPath()) ("word-finalize-" + [guid]::NewGuid().ToString() + ".log")
    $argumentList = @(
        "-NoProfile",
        "-Sta",
        "-ExecutionPolicy", "Bypass",
        "-File", $PSCommandPath,
        "-DocxPath", $resolvedDocx,
        "-OutputPath", $outputFullPath,
        "-Worker",
        "-LogPath", $tempLog
    )
    if ($AcceptRevisions.IsPresent) {
        $argumentList += "-AcceptRevisions"
    }

    $proc = Start-Process -FilePath "powershell.exe" -ArgumentList $argumentList -WindowStyle Hidden -PassThru

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

    if ($proc.ExitCode -eq 0 -and (Test-Path -LiteralPath $outputFullPath)) {
        Write-Output $outputFullPath
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
    step = "finalize_submission_copy"
    docx = $resolvedDocx
    output = $outputFullPath
    retry_count = $RetryCount
    timeout_seconds = $TimeoutSeconds
    attempts = $attemptErrors
    recovery_hint = "Close all background Word processes and retry. If it still fails, split the job into two steps: strip comments first, then refresh fields separately."
} | ConvertTo-Json -Depth 4

throw $errorPayload
