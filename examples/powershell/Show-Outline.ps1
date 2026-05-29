<#
.SYNOPSIS
    Print the heading outline of the Word document you have open right now.
.DESCRIPTION
    Read-only. Calls `wordlive status` and `wordlive outline`, parses the JSON,
    and prints an indented tree — a tour of the JSON-in / JSON-out CLI and how
    to branch on wordlive's deterministic exit codes.
.NOTES
    Prerequisites: Windows, Microsoft Word running with a document open, and
    `wordlive` on PATH (pip install wordlive  /  uv tool install wordlive).
.EXAMPLE
    .\Show-Outline.ps1
#>
[CmdletBinding()]
param()

$status = wordlive status | ConvertFrom-Json
if ($LASTEXITCODE -eq 4) {
    Write-Error 'Word is not running. Open Word and a document, then retry.'
    exit 4
}

$active = $status | Where-Object { $_.is_active }
Write-Host "Active document: $($active.name)`n"

$outline = wordlive outline | ConvertFrom-Json
if ($LASTEXITCODE -ne 0) {
    Write-Error "wordlive outline failed (exit $LASTEXITCODE)."
    exit $LASTEXITCODE
}

foreach ($entry in $outline) {
    $indent = '  ' * ($entry.level - 1)
    Write-Host "$indent$($entry.text)  [$($entry.anchor_id)]"
}
Write-Host "`n$($outline.Count) heading(s)."
