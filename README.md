# wordlive

[![PyPI](https://img.shields.io/pypi/v/wordlive.svg)](https://pypi.org/project/wordlive/)
[![Python versions](https://img.shields.io/pypi/pyversions/wordlive.svg)](https://pypi.org/project/wordlive/)
[![License: MIT](https://img.shields.io/pypi/l/wordlive.svg)](https://github.com/thomas-villani/wordlive/blob/main/LICENSE)
[![CI](https://github.com/thomas-villani/wordlive/actions/workflows/ci.yml/badge.svg)](https://github.com/thomas-villani/wordlive/actions/workflows/ci.yml)
[![Docs](https://github.com/thomas-villani/wordlive/actions/workflows/docs.yml/badge.svg)](https://thomas-villani.github.io/wordlive/)

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

Rendering pages to PNG (`snapshot`) needs the optional `snapshot` extra, which
pulls in PyMuPDF:

```
pip install "wordlive[snapshot]"
uv add "wordlive[snapshot]"
```

## Python

```python
import wordlive as wl

with wl.attach() as word:
    doc = word.documents.active

    # Reads
    outline = doc.outline()
    bookmarks = doc.bookmarks.list()

    # See it the way a vision model would — render a section to PNG
    # (needs `wordlive[snapshot]`):
    png = doc.heading("Introduction").snapshot()[0].png

    # Polite writes — preserves the user's cursor and view, atomic Ctrl-Z.
    with doc.edit("Update address block"):
        doc.bookmarks["Address"].set_text("123 Main St")
        doc.content_controls["Signatory"].set_text("Jane Doe")
        doc.heading("Introduction").insert_paragraph_after("New context paragraph.")

    # Hand back a deliverable (the Python API is ungated):
    doc.export_pdf("report.pdf")
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

# Navigate & locate by structure (read-only):
wordlive read section "Introduction"                          # body under a heading
wordlive read between --start heading:1 --end heading:3        # block between two headings
wordlive read nearest-heading --anchor-id para:42             # the section a paragraph sits in
wordlive find --text "exact phrase"                            # exact (normalized) → range:S-E
wordlive find-paragraph --text "roughly remembered text"       # FUZZY → ranked para:N + scores

# Insert a new paragraph relative to ANY anchor (heading, paragraph, bookmark, …):
wordlive insert --anchor-id heading:1 --text "..."          # after (default)
wordlive insert --anchor-id para:3 --text "..." --before
wordlive insert --anchor-id end --runs '[{"text":"Bold lead","bold":true},{"text":" — rest"}]'

# Drop a whole styled section in ONE op (item text takes **bold**/*italic* markdown):
wordlive insert-block --anchor-id heading:1 --items \
    '[{"text":"**Politeness** first.","style":"List Bullet"},"Atomic undo."]'
#   → reports range:START-END; then: wordlive list apply --anchor-id range:… --type bulleted

# Or hand it constrained Markdown and get real headings/lists/paragraphs (a subset, not CommonMark):
printf '# Plan\n\nKick-off.\n\n- scope it\n- staff it\n' | wordlive insert-markdown --anchor-id end --markdown -
#   insert-section adds a heading + body in one op; replace-section --anchor-id heading:N rewrites a section's body

# Append / prepend at the very end / start of the document (no anchor needed):
wordlive append  --text "Closing note."                     # new final paragraph
wordlive prepend --text "DRAFT" --inline                    # join the first paragraph

# Address anchors by ID (the IDs `outline`/`paragraphs` emit — `heading:N`, `para:N`, `bookmark:NAME`, `cc:NAME`):
wordlive replace --anchor-id heading:3 --text "Updated section text"
wordlive go-to --anchor-id bookmark:Address

# Durable handles — positional para:N/heading:N renumber under edits; pin for a stable id:
wordlive pin heading:3 --name methods             # -> {"pin": "pin:methods", ...}
wordlive replace --anchor-id pin:methods --text "Survives later inserts"

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
wordlive table records 1                                   # body rows as {header: value} dicts
wordlive replace --anchor-id table:1:2:2 --text "$450"
wordlive table add-row --table 1 --values '["Lodging", "$600"]'
wordlive table append-record --table 1 --record '{"Item":"Lodging","Cost":"$600"}'  # by header name
wordlive table update-row --table 1 --key Travel --values '{"Cost":"$450"}'         # match a row by content
wordlive table create --anchor-id end --data '[["Item","Cost"],["Travel","$400"]]' --header
wordlive table create --anchor-id end \
    --data '[{"Item":"Travel","Cost":"$400"}]'             # records → keys are a header row
#   --rows/--cols are inferred from --data (give them only to pad larger)
wordlive table autofit --table 1 --mode content            # fit columns to cells (or window/fixed)
wordlive table delete 2

# Page / column / section breaks (explicit one-off mark; for a style use --page-break-before):
wordlive insert-break --anchor-id heading:3 --kind page
wordlive format-paragraph --anchor-id heading:3 --page-break-before

# Collaboration: comments + track changes (the polite, non-destructive surface):
wordlive comment add --anchor-id heading:3 --text "Please expand this." --author Bot
wordlive comment list
wordlive comment resolve --index 1
wordlive track on            # record edits as revisions; `track off` to stop
wordlive revisions           # read the tracked changes back (type/author/text/range)
wordlive revision accept --index 1          # …or accept / reject them
wordlive revision accept-all --anchor-id heading:3   # accept every change in one section
wordlive read text --anchor-id para:5 --view original  # text as if changes were rejected
# …and `wordlive snapshot --markup all` renders those changes as visible marks.

# Non-visual layout introspection (reason about pages without a snapshot):
wordlive stats               # pages/words/…/tables/images/comments + saved, one read
wordlive locate --anchor-id heading:8   # {page, end_page, line, column, in_table}
wordlive proofing            # spelling/grammar errors + readability (Flesch, …)
wordlive read format --anchor-id heading:3   # effective style/paragraph/font, with override flags
wordlive lint                # audit formatting/structure: severity-ranked findings (some fixable)
wordlive regularize          # apply the fixable lint findings in one atomic-undo (--dry-run to preview)

# Document metadata, variables, and the read mirrors of link / insert-field:
wordlive properties list                         # built-in + custom document properties
wordlive properties set --name Title --value "Q3 Report"
wordlive variables set --name ClientName --value "Acme"   # invisible DOCVARIABLE storage
wordlive hyperlinks          # every link: text, destination, range:START-END id
wordlive fields              # every field: kind (PAGE/REF/TOC), code, rendered result

# Citations & bibliography (source → cite → build), and a legal table of authorities:
wordlive bibliography-style --style APA
wordlive add-source --type book --author "Smith, Jane" --title "On Risk" --year 2020   # → tag
wordlive insert-citation --anchor-id range:120-140 --tag Smith2020 --pages 15
wordlive insert-bibliography                       # works-cited block (update-fields to fill it)
wordlive mark-citation --anchor-id range:200-240 --long "Brown v. Board, 347 U.S. 483 (1954)" --category cases
wordlive table-of-authorities --category cases     # build TOA from the marks (update-fields for pages)

# Document themes / branding (apply a theme, then set brand colours & fonts):
wordlive apply-theme --theme Facet                 # a built-in theme or a .thmx path (list-themes for names)
wordlive set-theme-colors --accent1 "#1A73E8" --accent2 "#34A853"
wordlive set-theme-fonts --major Arial --minor Calibri
wordlive theme                                      # read back the 12 brand colours + major/minor fonts

# Lists & numbering (any anchor's paragraphs):
wordlive list apply --anchor-id heading:6 --type numbered
wordlive list restart --anchor-id heading:6

# Sections, headers & footers (header:S:WHICH / footer:S:WHICH):
wordlive sections
wordlive header write --section 1 --text "ACME Corporation"
wordlive footer read --section 1

# Images — from a file or base64 (--wrap is required: inline | auto | square | …):
wordlive insert-image --anchor-id heading:3 --path diagram.png --wrap auto
base64 logo.png | wordlive insert-image --anchor-id bookmark:Logo --base64 - --wrap inline --width 96

# …and read embedded pictures back out (image:N ids; for a vision model):
wordlive images                                            # list every embedded picture
wordlive read-image --anchor-id image:1 --out logo.png     # extract bytes + mime

# Floating-shape flourishes — a watermark behind every page, a pull-quote box:
wordlive watermark --text DRAFT                            # …or --remove
wordlive insert-text-box --anchor-id heading:2 --text "Key takeaway" --width 2.5in  # → shape:N

# …then restyle any floating shape in place (text box, floating image, WordArt) by shape:N:
wordlive shapes                                            # list shape:N ids, kind, size, wrap
wordlive set-shape-size --anchor-id shape:1 --width 3in --no-lock-aspect
wordlive format-shape --anchor-id shape:1 --fill navy --border-color white
wordlive replace-shape-image --anchor-id shape:2 --path v2.png   # swap a floating picture in place

# Equations — UnicodeMath (native), LaTeX (needs the `latex` extra), or MathML:
wordlive insert-equation --anchor-id heading:3 --unicodemath "x=(-b±√(b^2-4ac))/(2a)"
wordlive insert-equation --anchor-id heading:3 --latex "\frac{-b}{2a}"   # pip install "wordlive[latex]"
wordlive equations                                         # list equation:N ids, type, preview

# Charts — Excel-backed (needs Excel installed; data becomes static):
wordlive insert-chart --anchor-id end --kind bar --data '{"Q1": 10, "Q2": 25, "Q3": 18}' --title "Quarterly"
echo '[[1.2, 3.4], [2.5, 6.1]]' | wordlive insert-chart --anchor-id end --kind scatter --data -
wordlive charts                                            # list chart:N ids, kind, title
# Format & design an existing chart (no Excel needed; tri-state):
wordlive format-chart --anchor-id chart:1 --chart-style 240 --legend --title "Quarterly"
wordlive format-axis --anchor-id chart:1 --which value --scale log --title "USD (M)"
wordlive add-trendline --anchor-id chart:1 --kind power --display-equation   # law of best fit
wordlive set-series-color --anchor-id chart:1 --color "#2E86C1" --point 2

# Snapshot — render page(s) to PNG so a vision model can SEE the layout
# (needs the `snapshot` extra: pip install "wordlive[snapshot]"):
wordlive snapshot --anchor-id heading:3 --out section.png   # the section's page(s)
wordlive snapshot --page 2 --out p2.png                     # one page
wordlive snapshot --pages 1-3                               # base64 PNGs inline (JSON)
wordlive snapshot --max-dim 1000                            # whole doc, low-res layout check

# Batch multiple ops in a single Ctrl-Z:
wordlive exec --script ops.json

# Save / hand back a deliverable — GATED behind a directory whitelist
# (--save-dir or WORDLIVE_SAVE_DIRS; with none set, saving is off):
wordlive --save-dir C:\out save-as C:\out\report.docx
wordlive --save-dir C:\out export-pdf C:\out\report.pdf    # pixel-faithful PDF
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

Exit codes: `0` ok, `1` other, `2` anchor-not-found, `3` Word-busy, `4` Word-not-running, `5` ambiguous-match (`replace --find` hit several), `6` Excel-not-available (`insert-chart`).

## Agent skills

> Setting up a specific tool (Claude Code, Claude Desktop, Cursor, …)? The
> [Agents & LLM tools](https://thomas-villani.github.io/wordlive/agents/) guide
> has copy-paste setup per client.

wordlive ships **two** LLM-facing skills (`SKILL.md`): `wordlive-cli` (the
command-line workflow) and `wordlive-python` (the `import wordlive as wl` API).
Each covers the anchor model, every verb, and the exit-code / exception contract.

An agent that hits `wordlive --help` is pointed straight at `wordlive llm-help`,
which prints the whole guide to stdout in one shot — no install step, no Word:

```
wordlive llm-help                 # the CLI guide
wordlive llm-help --python        # the Python-API guide
```

Or drop the skill files into a project or your home directory so coding tools
discover them on their own (CLI skill by default; `--python` for just Python,
`--both` for both):

```
wordlive install-skill            # ./.agents/skills/wordlive-cli/SKILL.md
wordlive install-skill --both     # also drops wordlive-python/SKILL.md
wordlive install-skill --system   # into ~/.agents/skills/ instead
```

## MCP server (Claude Desktop & other agents)

Prefer MCP? `wordlive` ships a server so Claude Desktop and other MCP clients can
drive your open document directly. Three ways to set it up, easiest first:

**1. One-click bundle.** Download `wordlive.mcpb` (built from
[`mcpb/`](https://github.com/thomas-villani/wordlive/tree/main/mcpb)) and drop it
onto Claude Desktop → **Settings → Extensions**.

**2. `install-mcp`.** Register the server in your client's config in one command
(it uses `uvx`, so there's no separate install step):

```
wordlive install-mcp                      # → Claude Desktop's config
wordlive install-mcp --client claude-code # → ./.mcp.json
wordlive install-mcp --print              # just print the JSON snippet
```

**3. By hand.** `pip install "wordlive[mcp,snapshot]"` (the `snapshot` extra adds
the vision tool), then add to `claude_desktop_config.json`:

```json
{ "mcpServers": { "wordlive": { "command": "wordlive-mcp" } } }
```

It exposes four dispatch tools — `word_read`, `word_write`, `word_exec`, and
`word_snapshot` (which returns a rendered page as an image). The full op
vocabulary and anchor model are in the one-page guide, fetchable as a tool call
with `word_read(command="guide")` (also the `wordlive://guide` resource). Word
must be running on the same Windows machine. See
[docs/mcp.md](https://thomas-villani.github.io/wordlive/mcp/).

## Examples

Runnable, out-of-the-box scripts live in
[`examples/`](https://github.com/thomas-villani/wordlive/tree/main/examples) — Python
(using the library) and PowerShell (driving the CLI). Each attaches to the
document you already have open; the read-only and append-only ones are safe to
try on a real document.

```bash
python examples/python/read_outline.py            # read-only: print the outline
python examples/python/append_note.py "Reviewed." # append one paragraph (atomic, polite)
```

```powershell
.\examples\powershell\Show-Outline.ps1
.\examples\powershell\Invoke-WordliveWithRetry.ps1 write bookmark Address --text "123 Main St"
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

See [`spec.md`](https://github.com/thomas-villani/wordlive/blob/main/spec.md) for the full design.

## Contributing & security

Contributions are welcome — see
[`CONTRIBUTING.md`](https://github.com/thomas-villani/wordlive/blob/main/CONTRIBUTING.md)
for the dev setup (`uv`), the design invariants, and the test / lint / docs
gates. To report a security issue, please use private disclosure as described in
[`SECURITY.md`](https://github.com/thomas-villani/wordlive/blob/main/SECURITY.md)
rather than a public issue.
