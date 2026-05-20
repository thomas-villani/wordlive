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

# Address anchors by ID (the IDs `outline` emits — `heading:N`, `bookmark:NAME`, `cc:NAME`):
wordlive replace --anchor-id heading:3 --text "Updated section text"
wordlive go-to --anchor-id bookmark:Address

# Styles + paragraph formatting (atomic-undo):
wordlive style list
wordlive style apply --anchor-id heading:3 --name "Heading 2"
wordlive format-paragraph --anchor-id heading:3 --alignment center --space-before 6

# Tables (cells are anchors: table:N:R:C):
wordlive table list
wordlive table read 1
wordlive replace --anchor-id table:1:2:2 --text "$450"
wordlive table add-row --table 1 --values '["Lodging", "$600"]'

# Collaboration: comments + track changes (the polite, non-destructive surface):
wordlive comment add --anchor-id heading:3 --text "Please expand this." --author Bot
wordlive comment list
wordlive comment resolve --index 1
wordlive track on            # record edits as revisions; `track off` to stop

# Lists & numbering (any anchor's paragraphs):
wordlive list apply --anchor-id heading:6 --type numbered
wordlive list restart --anchor-id heading:6

# Sections, headers & footers (header:S:WHICH / footer:S:WHICH):
wordlive section list
wordlive header write --section 1 --text "ACME Corporation"
wordlive footer read --section 1

# Batch multiple ops in a single Ctrl-Z:
wordlive exec --script ops.json
```

Where `ops.json` looks like:

```json
{
  "label": "Update report",
  "ops": [
    {"op": "write_bookmark", "name": "Address", "text": "123 Main St"},
    {"op": "write_cc", "name": "Signatory", "text": "Jane Doe"},
    {"op": "insert_after_heading", "heading": "Risks", "text": "New risk paragraph."},
    {"op": "replace", "anchor_id": "heading:3", "text": "Updated section text"},
    {"op": "apply_style", "anchor_id": "heading:3", "name": "Heading 2"},
    {"op": "format_paragraph", "anchor_id": "heading:3", "alignment": "center", "space_before": 6}
  ]
}
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
