---
name: wordlive-python
description: Read and edit the Microsoft Word document the user has open right now, from Python — `import wordlive as wl`. Attach to the running Word instance, read structure (outline, paragraphs, tables, styles, sections, comments) as dicts/dataclasses, make polite edits inside `doc.edit("label")` blocks (each one atomic — a single Ctrl-Z), address content with stable anchors, render pages/sections to PNG for a vision model, and batch changes. Use when scripting a live Word document from Python on Windows.
---

# wordlive (Python API)

`wordlive` drives a **running** Microsoft Word instance over COM (Windows only).
Unlike `python-docx`, it edits the document the user has **open right now** — and
politely: their cursor, selection, and scroll position are preserved, and every
`doc.edit("…")` block collapses into a single Ctrl-Z.

```python
import wordlive as wl

with wl.attach() as word:        # attach to the already-running Word (raises if none)
    doc = word.documents.active  # the focused document
    print(doc.name)
```

- `wl.attach()` attaches to a running Word and **never launches** one
  (`WordNotRunningError` if none is up). Use
  `wl.connect(launch_if_missing=True, visible=True)` to start Word if needed.
- Both are context managers — do your work inside the `with` block.
- `word.documents.active`, `word.documents["Report.docx"]`, or iterate
  `for doc in word.documents:`.
- Every wrapper exposes the raw pywin32 object via `.com` (`word.com`, `doc.com`,
  `anchor.com`, `doc.selection.com`) — the escape hatch when wordlive doesn't
  cover something.

## Reading (no edit scope needed)

```python
doc.outline()                       # [{"level", "text", "anchor_id": "heading:N"}, …]
doc.paragraphs.list()               # every paragraph: index, anchor_id "para:N", level,
                                    #   style (Word style name), is_heading, start, end, text
doc.find("the risk register")       # fuzzy; [{"anchor_id": "range:S-E", "start","end","text"}]
doc.bookmarks["Address"].text       # bookmark text
doc.content_controls["Signatory"].text
doc.heading("Introduction").section_text()   # body under a heading, to the next ≤-level one
doc.tables[1].grid()                # 2-D list of cell text (1-based index, or by Title)
doc.styles.list()                   # [{"name","type","builtin","in_use"}]
doc.sections[1].page_setup          # margins / orientation / page size (points)
list(doc.comments)                  # review comments
doc.lists[1]                        # RangeAnchor over the 1st list (doc.lists is read-only)
doc.selection.info()                # {"start","end","collapsed","text"} — never moves the user
```

`find()` matching is forgiving: whitespace runs collapse, smart quotes/dashes
fold to ASCII, NBSPs become spaces, and strings are NFKC-normalised — so text an
LLM re-emitted still matches.

## Anchors — how you address things

Operations target stable **anchors**, never the live cursor. Resolve any anchor
from its id with `doc.anchor_by_id(...)`, or use the typed accessors:

| id form | accessor | refers to |
| --- | --- | --- |
| `heading:N` | `doc.heading("Text")` / `anchor_by_id` | the Nth *paragraph*, which must be a heading — same index space as `para:N` (so the first heading is rarely `heading:1`; copy from `outline()`) |
| `para:N` | `doc.paragraphs[N]` | the Nth paragraph (any kind) |
| `bookmark:NAME` | `doc.bookmarks["NAME"]` | a bookmark |
| `cc:NAME` | `doc.content_controls["NAME"]` | a content control (by Title, then Tag) |
| `footnote:N` / `endnote:N` | `doc.footnotes[N]` / `doc.endnotes[N]` | the Nth note's body (1-based) |
| `table:N:R:C` | `doc.tables[N].cell(R, C)` | a table cell |
| `range:START-END` | `doc.anchor_by_id("range:412-429")` | a raw character span (what `find()` emits) |
| `header:S:WHICH` / `footer:S:WHICH` | `doc.sections[S].header(WHICH)` | header/footer (`primary`/`first`/`even`) |
| `start` / `end` | `doc.start` / `doc.end` | before the first / past the last paragraph |

`para:N` and `heading:N` share an index space (a heading at `para:5` is also
`heading:5`); `heading:N` refuses to resolve a non-heading paragraph. The bare
`table:N` and `section:N` are collections, not anchors — use `doc.tables[N]` /
`doc.sections[N]`.

Every anchor shares the same verbs (from `Anchor`):

```python
a = doc.anchor_by_id("heading:3")
a.text                                  # read
a.set_text("Updated section text")      # replace the whole range
a.insert_before("Prefix ")              # inline insert, left
a.insert_after(" (verified)")           # inline insert, right
a.insert_paragraph_before("…", style="Body Text")
a.insert_paragraph_after("New paragraph.", style="Body Text")
a.apply_style("Heading 2")
a.format_paragraph(alignment="center", space_before=6, page_break_before=True)
a.format_run(bold=True, color="#FF0000", highlight="yellow", size="12pt")  # character formatting
a.set_shading(fill="navy")              # range/cell background fill
a.set_borders(sides="all", style="single", weight=0.5, color="black")
a.add_tab_stop("3in", align="right", leader="dots")
a.insert_break(kind="page")             # page | column | section_next | section_continuous
a.insert_field("page")                   # self-updating field: page|numpages|date|… (or "field" + raw code)
a.insert_footnote("See appendix B.")     # → Footnote (footnote:N); insert_endnote(...) mirrors
a.insert_toc(levels=(1, 3))              # table of contents → Toc; doc.add_toc() puts one at the top
a.insert_image("diagram.png", wrap="auto")
a.insert_table(2, 2, data=[["Item", "Cost"], ["Travel", "$400"]], header=True)
a.apply_list("numbered")                # + remove_list/list_info/restart_numbering/indent_list/outdent_list
a.snapshot("section.png")               # render the page(s) it sits on (heading → whole section)
a.delete()
a.com                                   # raw COM Range
```

Define and configure styles (the brand/template primitive — `add` once, then
`apply_style` everywhere):

```python
brand = doc.styles.add("Brand Heading", based_on="Heading 1", next_style="Body Text")
brand.format_run(bold=True, color="#1F3864", size="16pt")   # style's font defaults
brand.format_paragraph(space_before=12, space_after=4)      # style's paragraph defaults
doc.anchor_by_id("heading:3").apply_style("Brand Heading")
```

Page layout and page numbers (per section; `doc.sections[1]` is the whole
document for a single-section file):

```python
doc.sections[1].set_page_setup(orientation="landscape", margins="0.75in", columns=2)
foot = doc.sections[1].footer()         # a HeaderFooter *is* an anchor
foot.insert_field("page"); foot.insert_field("numpages")   # "Page X of Y" building blocks
foot.insert_page_number()               # sugar for insert_field("page")
doc.update_fields()                     # recompute page numbers / refs / dates
```

Footnotes / endnotes and a table of contents:

```python
note = doc.anchor_by_id("range:412-429").insert_footnote("Source: 2025 audit.")
note.set_text("Source: 2025 internal audit.")   # edit the body (footnote:N)
for f in doc.footnotes.list():          # [{index, anchor_id, marker, text, para}]
    print(f["anchor_id"], f["text"])

toc = doc.add_toc(levels=(1, 3))        # TOC at the document start
doc.update_fields()                     # populate its page numbers (or snapshot)
```

## Writing — wrap mutations in `doc.edit("label")`

`doc.edit("label")` opens a Word `UndoRecord` (one Ctrl-Z reverts the whole
block) **and** snapshots/restores the user's selection + scroll on exit.

```python
with doc.edit("Update address block"):
    doc.bookmarks["Address"].set_text("123 Main St")
    doc.content_controls["Signatory"].set_text("Jane Doe")
    doc.heading("Introduction").insert_paragraph_after("New context paragraph.")
```

The scope object exposes one knob — opt in to moving the user:

```python
with doc.edit("Insert and jump") as scope:
    doc.heading("Risks").insert_paragraph_after("Action items: …")
    scope.allow_cursor_move()           # skip the restore
    doc.go_to(doc.heading("Risks"))     # the one call that deliberately moves the cursor
```

### Fuzzy find + replace

```python
with doc.edit("Apply rewrite"):
    applied = doc.find_replace("Q1 2025", "Q2 2025")            # one match
    doc.find_replace("utilise", "use", all=True)                # every match
    doc.find_replace("needs review", "approved",
                     scope=doc.heading("Risks"), all=True)      # scoped to a section
```

Returns the list of replacements applied. Word's range-replace preserves the
matched span's character formatting (bold stays bold). Raises
`AmbiguousMatchError` (carries `.matches`) if there's more than one match and
neither `all` nor `occurrence=N` was given; `AnchorNotFoundError` on zero matches.
To edit text **inside a table**, scope to the cell — `scope=doc.tables[N].cell(R, C)`
(or `doc.anchor_by_id("table:N:R:C")`); a whole-document match that can't be
verified raises `ReplaceVerificationError` rather than risk the wrong cell.

### Append / prepend, build a doc from either end

```python
with doc.edit("Add notes"):
    doc.append_paragraph("Closing note.", style="Body Text")   # new final paragraph
    doc.append(" (verified)")                                  # continue the last paragraph
    doc.prepend_paragraph("DRAFT")                             # new first paragraph
# doc.end / doc.start are anchors too: doc.end.insert_image("logo.png", wrap="inline")
```

### Images

```python
with doc.edit("Add figure"):
    doc.heading("Risks").insert_image("diagram.png", wrap="auto")
    # bytes or base64 work too (an LLM usually holds data, not a path):
    doc.heading("Results").insert_image(png_bytes, wrap="square", width=240,
                                        alt_text="Chart")
```

`wrap` is **required**: `"inline"`, `"auto"` (Square if small, else top-and-bottom),
or a float wrap (`"square"`/`"tight"`/`"through"`/`"top-bottom"`/`"behind"`/`"front"`).
Pass `block=True` to drop the image on its own new line (e.g. with `where="before"`
at a heading, so it lands above the heading instead of joining the heading text).
A bad/unreadable/non-raster source raises `ImageSourceError` before anything is
inserted.

### Tables

```python
with doc.edit("Build + edit tables"):
    t = doc.add_table(3, 3, header=True,                       # append at end of doc
                      data=[["Tier", "Monthly", "SLA"],
                            ["Wobble", "$9", "best effort"],
                            ["Finch", "$99", "99.9%"]])
    t.cell(2, 2).set_text("$19")
    t.add_row(["Owl", "$199", "99.99%"])
    doc.heading("Budget").insert_table(2, 2)                   # at any position anchor
    doc.tables[1].delete_row(3)
    doc.tables[2].delete()
```

`add_table`/`insert_table` default to the `Table Grid` style (visible borders);
a cell *is* an anchor (`table:N:R:C`), so it takes `set_text`/`apply_style`/etc.

### Lists, sections, headers/footers

```python
with doc.edit("Number steps + brand the page"):
    doc.heading("Steps").apply_list("numbered")               # bulleted | numbered | outline
    sec = doc.sections[1]
    sec.header().set_text("ACME Corporation — Internal")
    sec.footer("first").set_text("Confidential")
```

### Suggest, don't overwrite — tracked changes

```python
with doc.tracked_changes(), doc.edit("Suggest plainer wording"):
    doc.find_replace("utilise", "use", all=True)
```

`tracked_changes()` turns Track Changes on for the scope and restores the prior
setting on exit; every edit lands as an accept/reject-able revision.
`doc.track_changes` is the underlying persistent bool.

### The explicit cursor surface

Everything above preserves the cursor. To act *at* it (the deliberately
non-default path):

```python
sel = doc.selection.info()                       # read where the user is
with doc.edit("Type at cursor") as scope:
    scope.allow_cursor_move()
    doc.selection.write("text at the caret", replace=True)
# Prefer: map the caret to an anchor and edit there instead —
para = doc.paragraphs.at(sel["start"])           # the containing Paragraph (para:N) or None
```

## Snapshot — render page(s) to PNG so a vision model can *see* the layout

```python
shots = doc.snapshot(pages=(1, 3), dpi=150)      # all | int | (start, end); read .png bytes
doc.snapshot("report.png", pages=2)              # also write to disk
png = doc.heading("Introduction").snapshot()[0].png   # an anchor's page(s); heading → section
```

Returns a list of `Snapshot` (`.page`, `.png` bytes, `.path`). Word exports a
pixel-faithful PDF and wordlive rasterises it — real fonts, spacing, geometry.
Read-only. Needs the optional `snapshot` extra (PyMuPDF), else `SnapshotError`.

## Errors — a small typed hierarchy

Catch the base `wl.WordliveError` for everything wordlive raises.

| Exception | When | Retryable? |
| --- | --- | --- |
| `WordNotRunningError` | no Word instance | no — until Word is up |
| `DocumentNotFoundError` | no document by that name / none active | no |
| `AnchorNotFoundError` | bookmark/cc/heading/cell/range/… missing, or `find` matched zero | after re-reading state |
| `StyleNotFoundError` | style not defined in the doc (subclass of the above) | after `doc.styles.list()` |
| `AmbiguousMatchError` | `find_replace` hit several; `.matches` lists them | yes — pass `all`/`occurrence` |
| `ReplaceVerificationError` | `find_replace` target unverifiable (e.g. a whole-doc match inside a table) | no — scope to the cell (`table:N:R:C`) |
| `ImageSourceError` | bad/unreadable/non-raster image | no — fix the input |
| `SnapshotError` | PyMuPDF missing or PDF rasterise failed | no — install `wordlive[snapshot]` |
| `WordBusyError` | a modal dialog is open / mid-operation; `.retryable` is True | **yes** — back off and retry |
| `ComError` | other classified COM error (`.hresult`, `.description`) | generally no |

```python
import time

def with_retry(fn, attempts=4, base=0.5):
    for i in range(attempts):
        try:
            return fn()
        except wl.WordBusyError:
            if i == attempts - 1:
                raise
            time.sleep(base * 2**i)      # 0.5, 1, 2, 4 s
```

## Typical workflow

1. `with wl.attach() as word: doc = word.documents.active`.
2. Discover anchors with `doc.outline()` / `doc.paragraphs.list()` / `doc.find(...)`.
3. Mutate inside one `with doc.edit("label"):` block (atomic, polite).
4. Read back to confirm; the user's cursor is untouched.

For the command line instead of Python, the `wordlive` CLI mirrors this model
(`wordlive outline`, `wordlive replace --anchor-id …`, `wordlive exec`); see the
`wordlive-cli` skill (`wordlive llm-help`). Full docs:
https://thomas-villani.github.io/wordlive/
