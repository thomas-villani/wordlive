<#
.SYNOPSIS
    Run a wordlive CLI command, retrying automatically while Word is busy.
.DESCRIPTION
    wordlive exits 3 when Word can't service the COM call right now — a modal
    dialog is open (Save As, Find & Replace, …) or it's mid-operation. That is
    the one exit code worth retrying. This wrapper retries exit 3 with
    exponential back-off, prints the parsed JSON on success, and surfaces any
    other non-zero exit immediately.
.PARAMETER WordliveArgs
    The wordlive subcommand and its arguments, e.g.
        outline
        write bookmark Address --text "123 Main St"
.PARAMETER MaxAttempts
    How many times to try before giving up (default 4).
.NOTES
    Prerequisites: Windows, Word running with a document open, `wordlive` on PATH.
.EXAMPLE
    .\Invoke-WordliveWithRetry.ps1 outline
.EXAMPLE
    .\Invoke-WordliveWithRetry.ps1 write bookmark Address --text "123 Main St"
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory, ValueFromRemainingArguments)]
    [string[]] $WordliveArgs,

    [int] $MaxAttempts = 4
)

for ($attempt = 1; $attempt -le $MaxAttempts; $attempt++) {
    $stdout = wordlive @WordliveArgs
    $code = $LASTEXITCODE

    if ($code -ne 3) {
        if ($code -ne 0) {
            Write-Error "wordlive exited $code"
        }
        if ($stdout) { $stdout | ConvertFrom-Json }
        exit $code
    }

    $delay = [math]::Pow(2, $attempt - 1)   # 1, 2, 4, 8 seconds
    Write-Warning "Word busy (exit 3); retrying in $delay s (attempt $attempt/$MaxAttempts)..."
    Start-Sleep -Seconds $delay
}

Write-Error "Word stayed busy after $MaxAttempts attempts."
exit 3
