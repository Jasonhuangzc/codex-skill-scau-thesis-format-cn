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

function Resolve-PythonRunner {
    $python = Get-Command python.exe -ErrorAction SilentlyContinue
    if ($python) {
        return @{
            FilePath = $python.Source
            PrefixArgs = @()
        }
    }

    $py = Get-Command py.exe -ErrorAction SilentlyContinue
    if ($py) {
        return @{
            FilePath = $py.Source
            PrefixArgs = @("-3")
        }
    }

    throw "Python 3 executable not found. Install Python and make sure python.exe or py.exe is on PATH."
}

function Invoke-FinalizeWorker {
    param(
        [string]$ResolvedDocx,
        [string]$ResolvedOutput,
        [bool]$ShouldAcceptRevisions,
        [string]$WorkerLogPath
    )
    try {
        $scriptDir = Split-Path -Parent $PSCommandPath
        $batchScript = Join-Path $scriptDir "batch_word_ops.py"
        if (-not (Test-Path -LiteralPath $batchScript)) {
            throw "batch_word_ops.py not found beside finalize_submission_copy.ps1"
        }

        $pythonRunner = Resolve-PythonRunner
        $plan = @()
        if ($ShouldAcceptRevisions) {
            $plan += @{ action = "accept_all_revisions" }
        }
        $plan += @(
            @{ action = "delete_all_comments" },
            @{
                action = "ensure_page_break_before"
                section = "english_abstract"
            },
            @{
                action = "normalize_body_paragraph_layout"
                first_line_indent_chars = 2.0
                left_indent_chars = 0.0
                line_spacing = 18.0
                line_spacing_rule = 1
                alignment = 3
            },
            @{
                action = "normalize_table_cells"
                target = "all"
                apply_fonts = $false
                first_line_indent_chars = 0.0
                left_indent_chars = 0.0
                line_spacing = 18.0
                line_spacing_rule = 1
                alignment = 1
            },
            @{
                action = "normalize_table_cells"
                target = "abbreviation"
                apply_fonts = $true
                far_east_font = "宋体"
                ascii_font = "Times New Roman"
                size = 12
                first_line_indent_chars = 0.0
                left_indent_chars = 0.0
                line_spacing = 18.0
                line_spacing_rule = 1
                alignment = 1
            },
            @{
                action = "normalize_tail_section_fonts"
                sections = @("references", "acknowledgements")
            },
            @{
                action = "finalize_contents"
                mode = "full"
                update_fields = $true
                ascii_font = "Times New Roman"
            }
        )

        $planPath = Join-Path ([System.IO.Path]::GetTempPath()) ("word-finalize-plan-" + [guid]::NewGuid().ToString() + ".json")
        $tempOutput = Join-Path ([System.IO.Path]::GetTempPath()) ("word-finalize-output-" + [guid]::NewGuid().ToString() + ".docx")
        $plan | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath $planPath -Encoding UTF8

        $argumentList = @()
        $argumentList += $pythonRunner.PrefixArgs
        $argumentList += @(
            $batchScript,
            $ResolvedDocx,
            $planPath,
            "--output",
            $tempOutput
        )

        $pythonResult = & $pythonRunner.FilePath @argumentList 2>&1
        if ($LASTEXITCODE -ne 0) {
            $message = @(
                "batch_word_ops failed",
                ($pythonResult | Out-String)
            ) -join [Environment]::NewLine
            throw $message.Trim()
        }
        if (-not (Test-Path -LiteralPath $tempOutput)) {
            throw "batch_word_ops completed without creating the finalized output file."
        }

        Move-Item -LiteralPath $tempOutput -Destination $ResolvedOutput -Force
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
        if ($planPath -and (Test-Path -LiteralPath $planPath)) {
            Remove-Item -LiteralPath $planPath -Force -ErrorAction SilentlyContinue
        }
        if ($tempOutput -and (Test-Path -LiteralPath $tempOutput)) {
            Remove-Item -LiteralPath $tempOutput -Force -ErrorAction SilentlyContinue
        }
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
