# wordlive

Drive a running Microsoft Word instance from Python — `xlwings`, but for Word.
Built for both human scripting and LLM agents. Windows-only.

## Install

```
pip install wordlive
```

(Requires Python 3.13+ and `pywin32` on Windows.)

## Python

```python
import wordlive as wl

with wl.attach() as word:
    doc = word.documents.active

    # Reads
    outline = doc.outline()
    bookmarks = doc.bookmarks.list()

    # Polite writes — preserves the user's cursor and view, atomic Ctrl-Z.
    with doc.edit("Update address block"):
        doc.bookmarks["Address"].set_text("123 Main St")
        doc.content_controls["Signatory"].set_text("Jane Doe")
        doc.heading("Introduction").insert_paragraph_after("New context paragraph.")
```

## CLI

JSON in, JSON out — designed to drop straight into an LLM tool-use loop:

```
wordlive status
wordlive outline
wordlive read bookmark Address
wordlive write bookmark Address --text "123 Main St"
wordlive insert --after-heading "Introduction" --text "..."
```

Exit codes: `0` ok, `2` anchor-not-found, `3` Word-busy, `4` Word-not-running, `1` other.

## Design

- **Politeness first** — operations preserve the user's `Selection`, view, and
  scroll. The user keeps editing alongside you.
- **Semantic anchors over `Selection`** — operations target bookmarks, content
  controls, or headings — never the live cursor unless you ask.
- **Atomic undo** — every `doc.edit()` opens a Word `UndoRecord`, so a single
  Ctrl-Z reverts the whole block.
- **Escape hatch** — every wrapper exposes `.com` for the raw COM object;
  you're never blocked by missing coverage.

See [`spec.md`](spec.md) for the full design.
