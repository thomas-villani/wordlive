# Examples

The repo ships runnable, out-of-the-box scripts in
[`examples/`](https://github.com/thomas-villani/wordlive/tree/main/examples), in
both Python and PowerShell. Each **attaches to the Word instance you already
have open**, so open Word with a document before running. The read-only and
append-only examples are safe to try on a real document — an append is a single
paragraph you can remove with one Ctrl-Z.

## Python

Use the library directly (`import wordlive as wl`):

| Script | What it does | Mutates? |
| --- | --- | --- |
| [`read_outline.py`](https://github.com/thomas-villani/wordlive/blob/main/examples/python/read_outline.py) | Print the heading outline + paragraph count of the active document. | No |
| [`append_note.py`](https://github.com/thomas-villani/wordlive/blob/main/examples/python/append_note.py) | Append a timestamped note at the end, atomically, preserving your cursor. | Appends 1 ¶ |
| [`fuzzy_replace.py`](https://github.com/thomas-villani/wordlive/blob/main/examples/python/fuzzy_replace.py) | Fuzzy find-and-replace recorded as tracked changes. | Tracked |
| [`snapshot_page.py`](https://github.com/thomas-villani/wordlive/blob/main/examples/python/snapshot_page.py) | Render a page to PNG for a vision model (needs the `snapshot` extra). | No |

```bash
pip install wordlive                       # or: uv add wordlive
python examples/python/read_outline.py
python examples/python/append_note.py "Reviewed."
python examples/python/fuzzy_replace.py "utilise" "use" --all
pip install "wordlive[snapshot]"
python examples/python/snapshot_page.py 1 page.png
```

## PowerShell

Drive the [`wordlive` CLI](cli.md) — JSON in, JSON out, deterministic
[exit codes](errors.md#cli-exit-codes):

| Script | What it does | Mutates? |
| --- | --- | --- |
| [`Show-Outline.ps1`](https://github.com/thomas-villani/wordlive/blob/main/examples/powershell/Show-Outline.ps1) | Status + outline as an indented tree. | No |
| [`Append-Note.ps1`](https://github.com/thomas-villani/wordlive/blob/main/examples/powershell/Append-Note.ps1) | Append a note via an `exec` batch piped over stdin. | Appends 1 ¶ |
| [`Invoke-WordliveWithRetry.ps1`](https://github.com/thomas-villani/wordlive/blob/main/examples/powershell/Invoke-WordliveWithRetry.ps1) | Run any wordlive command, retrying while Word is busy (exit `3`). | Depends |

```powershell
.\examples\powershell\Show-Outline.ps1
.\examples\powershell\Append-Note.ps1 -Text 'Reviewed and approved.'
.\examples\powershell\Invoke-WordliveWithRetry.ps1 write bookmark Address --text "123 Main St"
```

These are the same patterns walked through in the [Cookbook](cookbook.md), packaged
to run end-to-end.
