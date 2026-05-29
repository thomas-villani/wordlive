# Examples

Runnable, out-of-the-box scripts for wordlive. Each one **attaches to the Word
instance you already have open**, so open Word with a document before running.

The read-only and append-only examples are safe to try on a real document — an
append is a single paragraph you can remove with one Ctrl-Z.

## Python (`python/`)

| Script | What it does | Mutates? |
| --- | --- | --- |
| [`read_outline.py`](python/read_outline.py) | Print the heading outline + paragraph count of the active document. | No |
| [`append_note.py`](python/append_note.py) | Append a timestamped note at the end, atomically, preserving your cursor. | Appends 1 ¶ |
| [`fuzzy_replace.py`](python/fuzzy_replace.py) | Fuzzy find-and-replace recorded as tracked changes. | Tracked |
| [`snapshot_page.py`](python/snapshot_page.py) | Render a page to PNG for a vision model (needs the `snapshot` extra). | No |

```bash
pip install wordlive                       # or: uv add wordlive
python python/read_outline.py
python python/append_note.py "Reviewed."
python python/fuzzy_replace.py "utilise" "use" --all
pip install "wordlive[snapshot]"
python python/snapshot_page.py 1 page.png
```

## PowerShell (`powershell/`)

These drive the `wordlive` CLI — JSON in, JSON out, deterministic exit codes.

| Script | What it does | Mutates? |
| --- | --- | --- |
| [`Show-Outline.ps1`](powershell/Show-Outline.ps1) | Status + outline as an indented tree. | No |
| [`Append-Note.ps1`](powershell/Append-Note.ps1) | Append a note via an `exec` batch piped over stdin. | Appends 1 ¶ |
| [`Invoke-WordliveWithRetry.ps1`](powershell/Invoke-WordliveWithRetry.ps1) | Run any wordlive command, retrying while Word is busy (exit 3). | Depends |

```powershell
.\powershell\Show-Outline.ps1
.\powershell\Append-Note.ps1 -Text 'Reviewed and approved.'
.\powershell\Invoke-WordliveWithRetry.ps1 write bookmark Address --text "123 Main St"
```

For the patterns these scripts are built from, see the
[Cookbook](https://thomas-villani.github.io/wordlive/cookbook/) and the
[CLI reference](https://thomas-villani.github.io/wordlive/cli/).
