# wordlive

Drive a running Microsoft Word instance from Python — `xlwings`, but for Word.
Built for both human scripting and LLM agents. Windows-only.

## Install

```
pip install wordlive

# Add to a python project
uv add wordlive

# Or as a `uv` tool
uv tool install wordlive
```

(Requires Python 3.10+ and `pywin32` on Windows.)

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
wordlive outline                  # heading structure (heading:N)
wordlive outline --all            # every paragraph (para:N) — alias of `paragraphs`
wordlive paragraphs               # same: para:N, level, offsets, text
wordlive read bookmark Address
wordlive write bookmark Address --text "123 Main St"

# Insert a new paragraph relative to ANY anchor (heading, paragraph, bookmark, …):
wordlive insert --anchor-id heading:1 --text "..."          # after (default)
wordlive insert --anchor-id para:3 --text "..." --before

# Append / prepend at the very end / start of the document (no anchor needed):
wordlive append  --text "Closing note."                     # new final paragraph
wordlive prepend --text "DRAFT" --inline                    # join the first paragraph

# Address anchors by ID (the IDs `outline`/`paragraphs` emit — `heading:N`, `para:N`, `bookmark:NAME`, `cc:NAME`):
wordlive replace --anchor-id heading:3 --text "Updated section text"
wordlive go-to --anchor-id bookmark:Address

# Explicit cursor surface (the non-preferred mode — deliberately moves the cursor):
wordlive cursor read                              # where is the cursor? which para:N?
wordlive cursor write --text "inserted here"      # type at the cursor

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

# Images — from a file or base64 (--wrap is required: inline | auto | square | …):
wordlive insert-image --anchor-id heading:3 --path diagram.png --wrap auto
base64 logo.png | wordlive insert-image --anchor-id bookmark:Logo --base64 - --wrap inline --width 96

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
    {"op": "insert_paragraph", "anchor_id": "heading:3", "text": "New risk paragraph."},
    {"op": "replace", "anchor_id": "heading:3", "text": "Updated section text"},
    {"op": "apply_style", "anchor_id": "heading:3", "name": "Heading 2"},
    {"op": "format_paragraph", "anchor_id": "heading:3", "alignment": "center", "space_before": 6}
  ]
}
```

Exit codes: `0` ok, `2` anchor-not-found, `3` Word-busy, `4` Word-not-running, `1` other.

## Agent skill

wordlive ships an LLM-facing skill (`SKILL.md`) — a concise CLI reference for
agents. Drop it into a project or your home directory so coding tools discover it:

```
wordlive install-skill            # ./.agents/skills/wordlive/SKILL.md
wordlive install-skill --system   # ~/.agents/skills/wordlive/SKILL.md
```

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
