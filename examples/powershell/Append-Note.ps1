<#
.SYNOPSIS
    Append a note to the end of the open document in a single atomic undo.
.DESCRIPTION
    Builds an `exec` batch and pipes it to `wordlive exec --ops -` over stdin
    (which sidesteps Windows command-line quoting). The whole batch reverts with
    one Ctrl-Z. Safe to run on any document — it only appends one paragraph.

    Note: the JSON is assembled as a here-string with the note text run through
    ConvertTo-Json so it's correctly escaped. (Windows PowerShell 5.1 unwraps a
    single-element array when ConvertTo-Json serialises a hashtable property, so
    a hand-built `ops` literal is the reliable way to keep it a JSON array.)
.NOTES
    Prerequisites: Windows, Word running with a document open, `wordlive` on PATH.
.EXAMPLE
    .\Append-Note.ps1
.EXAMPLE
    .\Append-Note.ps1 -Text 'Reviewed and approved.'
#>
[CmdletBinding()]
param(
    [string] $Text = "Note added $(Get-Date -Format 'yyyy-MM-dd HH:mm') by Append-Note.ps1"
)

# ConvertTo-Json on the bare string yields a properly-escaped, quoted JSON value.
$textJson = $Text | ConvertTo-Json
$batch = @"
{"label": "Append note (example)", "ops": [{"op": "append_paragraph", "text": $textJson}]}
"@

$result = $batch | wordlive exec --ops - | ConvertFrom-Json
switch ($LASTEXITCODE) {
    0 { Write-Host "Appended note ($($result.ops_run) op). Undo with a single Ctrl-Z." }
    3 { Write-Error 'Word is busy (a dialog may be open). Try again.'; exit 3 }
    4 { Write-Error 'Word is not running.'; exit 4 }
    default {
        Write-Error "wordlive exec failed (exit $LASTEXITCODE): $($result | ConvertTo-Json -Compress)"
        exit $LASTEXITCODE
    }
}
