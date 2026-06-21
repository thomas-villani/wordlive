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
doc.find("the risk register")       # exact substring (normalized); [{"anchor_id":"range:S-E",…}]
doc.find_paragraphs("approx text")  # FUZZY ranking (difflib); [{"anchor_id":"para:N","score",…}]
doc.bookmarks["Address"].text       # bookmark text
doc.content_controls["Signatory"].text
doc.heading("Introduction").section_text()   # body under a heading, to the next ≤-level one
doc.to_markdown()                   # whole doc → clean Markdown (to_html for HTML); within=ID scopes
doc.to_markdown(within="heading:3") # one anchor's range; lossy mirror of insert_markdown
doc.read(budget=4000)               # token-budgeted whole-doc digest; headings verbatim, body sampled, every anchor addressable
doc.between("heading:1", "heading:3")         # RangeAnchor between two anchors (excl. headings)
doc.nearest_heading("para:42")                # heading row nearest a position (direction=before|after)
doc.tables[1].grid()                # 2-D list of cell text (1-based index, or by Title)
doc.styles.list()                   # [{"name","type","builtin","in_use"}]
doc.sections[1].page_setup          # margins / orientation / page size (points)
list(doc.comments)                  # review comments
doc.lists[1]                        # RangeAnchor over the 1st list (doc.lists is read-only)
doc.selection.info()                # {"start","end","collapsed","text"} — never moves the user
doc.tables[1].records()             # body rows as [{header: value}, …] (row 1 = header)
doc.stats()                         # {pages,words,characters,paragraphs,lines,sections,
                                    #   headings,tables,images,comments,revisions,saved}
anchor.location()                   # {page,end_page,line,column,in_table} — "what page is this on"
```

`stats()` and `location()` are the non-visual layout reads — answer "how long is
this" / "what page is this on" without a `snapshot` vision pass. Both repaginate
first (content-neutral: selection/scroll/view untouched), so page numbers are
print-layout truth.

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
| `pin:CODE` | `doc.anchor_by_id("pin:CODE")` | a **durable handle** from `doc.pin(...)` / `doc.pin_outline()` — survives renumbering |
| `cc:NAME` | `doc.content_controls["NAME"]` | a content control (by Title, then Tag) |
| `footnote:N` / `endnote:N` | `doc.footnotes[N]` / `doc.endnotes[N]` | the Nth note's body (1-based) |
| `image:N` | `doc.images[N]` | the Nth embedded picture (1-based, Word's `InlineShapes` order) |
| `equation:N` | `doc.equations[N]` | the Nth equation (1-based, Word's `OMaths` order) |
| `chart:N` | `doc.charts[N]` | the Nth chart (1-based, document order) |
| `shape:N` | `doc.shapes[N]` | the Nth floating shape — text box / floating image / WordArt / group (1-based, document order) |
| `textbox:N` | `doc.text_boxes[N]` | the Nth text box — an alias onto its canonical `shape:N` |
| `table:N:R:C` | `doc.tables[N].cell(R, C)` | a table cell |
| `table:N:row:R` / `table:N:col:C` | `doc.tables[N].row(R)` / `.column(C)` | a whole row / column — a styling handle (`set_shading`/`set_borders`/`apply_style`/`format_run`) |
| `range:START-END` | `doc.anchor_by_id("range:412-429")` | a raw character span (what `find()` emits) |
| `header:S:WHICH` / `footer:S:WHICH` | `doc.sections[S].header(WHICH)` | header/footer (`primary`/`first`/`even`) |
| `start` / `end` | `doc.start` / `doc.end` | before the first / past the last paragraph |

`para:N` and `heading:N` share an index space (a heading at `para:5` is also
`heading:5`); `heading:N` refuses to resolve a non-heading paragraph. The bare
`table:N` and `section:N` are collections, not anchors — use `doc.tables[N]` /
`doc.sections[N]`.

`para:N` / `heading:N` are **positional** — they renumber under inserts/deletes,
and a stale one raises `AnchorNotFoundError` whose `.hint` explains why
(out-of-range vs not-a-heading, the nearest heading) and points to pinning. For a
handle that survives structural edits, `doc.pin(anchor)` mints a hidden bookmark
and returns a `pin:CODE` id (Word keeps it pinned to the same content);
`doc.pin_outline()` does this for every heading at once (idempotent).

```python
pin = doc.pin("para:7")["pin"]          # -> "pin:a3f9c2" (or doc.pin(a, name="intro"))
doc.pin_outline()                       # {"heading:1": "pin:…", …} — durable nav scaffold
doc.anchor_by_id(pin).set_text("…")     # resolves even after later inserts shift para ids
```

Every anchor shares the same verbs (from `Anchor`):

```python
a = doc.anchor_by_id("heading:3")
a.text                                  # read
a.set_text("Updated section text")      # replace the whole range
a.insert_before("Prefix ")              # inline insert, left
a.insert_after(" (verified)")           # inline insert, right
a.insert_paragraph_before("…", style="Body Text")
a.insert_paragraph_after("New paragraph.", style="Body Text")
a.insert_block([                         # contiguous styled paragraphs in one op → RangeAnchor
    {"text": "**Politeness** first.", "style": "List Bullet"},   # item text takes **bold**/*italic*
    {"runs": [{"text": "Atomic undo", "bold": True}, {"text": " — one Ctrl-Z."}], "style": "List Bullet"},
    "Plain third bullet.",
])                                       # → feed rng.anchor_id to apply_list("bulleted")
a.insert_section("Results", ["Body para.", "Another."], level=2)  # heading + body in one op → RangeAnchor
a.insert_markdown("# Title\n\nIntro.\n\n- bullet one\n- bullet two\n\n1. step")  # constrained MD → Word structure
doc.headings["Results"].replace_section_body("New body.", markdown=True)  # rewrite a section, keep its heading
a.apply_style("Heading 2")
a.format_paragraph(alignment="center", space_before=6, line_spacing=1.5,  # leading: 1.5/"double"/"14pt"
                   page_break_before=True, keep_with_next=True, keep_together=True, widow_control=True)
a.format_run(bold=True, color="#FF0000", highlight="yellow", size="12pt")  # character formatting
a.format_info()                         # read mirror of format_paragraph/format_run: {style, paragraph, font}
                                        #   each field → {value, style, override}; font.mixed = fields varying across runs
a.set_shading(fill="navy")              # range/cell background fill
a.set_borders(sides="all", style="single", weight=0.5, color="black")
a.drop_cap(3, position="dropped", font="Georgia")  # oversized initial letter; position="none" removes
a.add_tab_stop("3in", align="right", leader="dots")
a.insert_break(kind="page")             # page | column | section_next | section_continuous
a.insert_field("page")                   # self-updating field: page|numpages|date|… (or "field" + raw code)
a.insert_footnote("See appendix B.")     # → Footnote (footnote:N); insert_endnote(...) mirrors
a.insert_toc(levels=(1, 3))              # table of contents → Toc; doc.add_toc() puts one at the top
a.insert_table_of_figures(label="Figure")  # lists captions of one label → TableOfFigures (field block, like the TOC)
a.mark_index_entry("risk:market")        # mark range as XE index entry ("main:sub" nests); then build the index:
a.insert_index(columns=2)                # back-of-book index from the marks → Index; doc.add_index() puts one at the end
a.insert_citation("Smith2020", pages="15")  # in-text CITATION field → Citation (renders per doc.bibliography_style)
a.insert_bibliography()                  # works-cited block → Bibliography; doc.add_bibliography() puts one at the end
a.mark_citation("Brown v. Board, 347 U.S. 483 (1954)", category="cases")  # mark a TA entry; then build the table:
a.insert_table_of_authorities(category="all")  # → TableOfAuthorities; doc.add_table_of_authorities() at the end
a.link_to(address="https://x")          # hyperlink; or link_to(bookmark="Intro"); text= inserts new linked text
a.insert_cross_reference("bookmark:Intro", kind="page")  # ref a bookmark/heading/footnote/endnote
a.insert_caption("Figure", text="System overview")       # own-paragraph caption (Table→above, else below; position= to override)
a.insert_content_control(kind="dropdown", title="Status", items=["Open", "Done"])  # wrap the range in a CC; cc:Status addresses it
doc.content_controls["Status"].set_properties(title="Stage", lock_contents=True)   # edit a CC in place (no delete+reinsert)
doc.content_controls["Status"].set_items(["Open", "Done", "Blocked"])             # replace a combo_box/dropdown's choices
img = a.insert_image("diagram.png", wrap="auto")  # floating wrap → ShapeAnchor (shape:N); wrap="inline" → None (stays image:N)
a.read_image()                          # → (bytes, mime) — extract the one image in the range
tb = a.insert_text_box("Key takeaway", width="2.5in", fill="#eeeeff")  # → ShapeAnchor (shape:N)
tb.set_wrap("tight").set_size(width="3in", height="1in").format(border="navy")  # chainable restyle-in-place
tb.set_wrap(side="left", distance_top="0.1in")                    # which sides text flows past + standoff gaps (square/tight/through)
tb.set_position(left="center", relative_to="margin"); tb.set_text("Revised")    # left/top = length or "center"
tb.set_rotation(15).set_z_order("front")                          # absolute angle; restack within the float layer
tb.set_text_frame(margin_left="0.1in", word_wrap=False)          # a text box's internal insets / word-wrap
img.replace_image("v2.png")             # swap a floating picture's bits in place (preserves wrap/position/size)
img.set_crop(left="0.2in", bottom=6)    # trim a PICTURE shape in from its edges (a text box would raise)
g = doc.group_shapes("shape:1", "shape:2"); members = g.ungroup()  # group two+ floats into one; ungroup → [ShapeAnchor]
doc.images[1].set_alt_text("Fig. 1").set_size(width="3in").set_crop(right="0.1in")  # restyle/crop an INLINE picture (re-wrap = float = insert_image)
# doc.shapes lists every floating shape; doc.text_boxes is the text-box subset (each keeps its shape:N id)
# shape:N is positional (document order) — it renumbers as shapes are added/removed; re-list, don't cache
# (no autosize knob: Word's "resize-to-fit-text" doesn't expose cleanly over COM)
a.insert_equation(unicodemath="x=(-b±√(b^2-4ac))/(2a)")   # native; or latex= (needs the `latex` extra) / mathml=
a.insert_equation(latex=r"\frac{-b}{2a}", display=False)  # → EquationAnchor; display=True→centred "Equation" style, False→Normal+left
# equation:N is positional (OMaths order) — inserting one before another renumbers it; re-list, don't cache the id
a.insert_chart("bar", {"Q1": 10, "Q2": 25, "Q3": 18}, title="Quarterly")  # → ChartAnchor (chart:N); Excel-backed
a.insert_chart("scatter", [[1.2, 3.4], [1.2, 3.9], [2.5, 6.1]])  # [x,y] pairs → numeric axes, duplicate x kept
# charts need Excel installed (else ExcelNotAvailableError); data is made static (no embedded workbook ships); doc.charts lists them
c = doc.charts[1]                          # ChartAnchor — formatting/design (no Excel needed; chainable; tri-state)
c.format(title="Revenue", legend=True, legend_position="bottom", chart_style=242, background="#F4F6F7", data_labels=True)
c.set_axis("value", title="USD", minimum=0, maximum=30, scale="log")  # which=value|y|category|x; scale linear|log
c.add_trendline(kind="power", display_equation=True)  # linear|exponential|logarithmic|moving_average|polynomial|power
c.set_series_color("#2E86C1")              # whole series; pass point=N to recolour one bar/slice; c.format(chart_type="line") re-types
a.insert_table(data=[["Item", "Cost"], ["Travel", "$400"]], header=True)  # rows/cols inferred from data
a.insert_table(data=[{"Item": "Travel", "Cost": "$400"}])  # records → keys become a bolded header row
a.apply_list("numbered")                # + remove_list/list_info/restart_numbering/indent_list/outdent_list
a.snapshot("section.png")               # render the page(s) it sits on (heading → whole section)
a.delete()
a.com                                   # raw COM Range
```

To **number several paragraphs 1, 2, 3**, apply the list over a *single* range
that spans them — `doc.range(first_start, last_end).apply_list("numbered")` (or a
heading's section). Applying `numbered` to each paragraph one at a time makes N
independent "1." lists; `continue_previous=True` chains a clean in-order apply
but can't *repair* an already-split list (remove the list over the span and
re-apply to fix that).

Delete a whole paragraph (text **and** its mark, so no blank line is left —
unlike `set_text("")`) with `doc.delete_paragraph("para:3")`.

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

Citations, bibliography & a table of authorities — source → cite → build:

```python
doc.bibliography_style = "APA"                      # the citation scheme (MLA/Chicago/IEEE/… — build-dependent)
doc.sources.add("book", author="Smith, Jane", title="On Risk", year=2020)  # tag auto-derives → "Smith2020"
doc.anchor_by_id("range:120-140").insert_citation("Smith2020", pages="15")  # (Smith 2020, 15)
bib = doc.add_bibliography()                         # works-cited block at the document end
doc.update_fields()                                 # fill its entries/page numbers (or snapshot)

# A table of authorities is the same mark-then-build pattern as the index:
doc.anchor_by_id("range:200-240").mark_citation("Brown v. Board, 347 U.S. 483 (1954)", category="cases")
toa = doc.add_table_of_authorities(category="all")  # build from the TA marks (toa.update() refreshes — no update_page_numbers())
doc.update_fields()
```

`doc.sources` is a `SourceCollection` (list/index by tag, `in`, `len`); each
`Source` has `.tag`/`.cited`/`.xml`/`.to_dict()`/`.delete()`. `doc.sources.add_xml(
"<b:Source>…</b:Source>")` is the raw-OOXML escape hatch. A citation to an
unregistered tag still inserts but renders "Invalid source specified.". The
`Bibliography` and `TableOfAuthorities` are field blocks like the TOC, so their
page numbers populate only after `doc.update_fields()` (or a `snapshot`).

Document theme / branding — the document-wide brand primitive (`doc.theme`):

```python
with doc.edit("brand"):
    doc.theme.apply("Facet")                        # a whole theme (built-in name or .thmx path)
    doc.theme.set_colors(scheme="Blue", accent1="#1A73E8")  # named scheme + per-colour overrides
    doc.theme.set_fonts(major="Arial", minor="Calibri")     # heading / body fonts
doc.theme.colors          # {"accent1": "#1A73E8", "text1": "#000000", …} (12 brand colours)
doc.theme.to_dict()       # {"colors": {...}, "major_font": "Arial", "minor_font": "Calibri"}
doc.theme.list_available()  # {"themes": [...], "color_schemes": [...], "font_schemes": [...]}
```

`set_colors` keys are `text1`/`background1`/`text2`/`background2`/`accent1`–
`accent6`/`hyperlink`/`followed_hyperlink` and take a colour name, hex string, or
`(r, g, b)` tuple; `scheme`/`theme` accept a built-in name (see `list_available`)
or a file path.

Anchoring & linking — name a target, then point at it:

```python
doc.bookmarks.add("Intro", "heading:2")          # create a bookmark (name validated)
doc.headings["Methods"].link_to(bookmark="Intro", text="see Intro")   # internal jump
doc.end.insert_cross_reference("bookmark:Intro", kind="page")         # "see page N"
doc.headings["Conclusion"].insert_caption("Figure", text="Overview")  # Figure 1: Overview
doc.update_fields()                              # refresh cross-ref / page numbers after edits
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

**Reading images back out** (for a vision model): `doc.images` lists every
embedded picture, and `read_image()` returns the raw bytes + MIME type.

```python
for img in doc.images.list():   # [{index, anchor_id, mime, width, height, alt_text, para}]
    print(img["anchor_id"], img["mime"], img["alt_text"])
data, mime = doc.images[1].read_image()          # → (b"\x89PNG…", "image/png")
data, mime = doc.anchor_by_id("para:7").read_image()  # the single image in a paragraph
```

`read_image()` needs the range to hold exactly one picture (an `image:N` anchor
always does); zero or several raise `ImageSourceError`. Reading is non-mutating —
no `doc.edit()` needed.

**Reading equations back out**: `doc.equations` lists every equation, and
`EquationAnchor.mathml` round-trips one back to MathML (non-mutating).

```python
for eq in doc.equations.list():   # [{index, anchor_id, type, linear, para}]
    print(eq["anchor_id"], eq["type"], eq["linear"])
doc.equations[1].mathml           # → "<math …>…</math>"  (via Office's OMML→MathML transform)
```

### Tables

```python
with doc.edit("Build + edit tables"):
    t = doc.add_table(3, 3, header=True,                       # append at end of doc
                      data=[["Tier", "Monthly", "SLA"],
                            ["Wobble", "$9", "best effort"],
                            ["Finch", "$99", "99.9%"]])
    t.cell(2, 2).set_text("$19")
    t.add_row(["Owl", "$199", "99.99%"])
    t.append_record({"Tier": "Hawk", "Monthly": "$299"})      # append by header name
    t.update_row("Wobble", {"Monthly": "$12"})                # match first column, set by header
    doc.heading("Budget").insert_table(2, 2)                   # at any position anchor
    doc.tables[1].delete_row(3)
    doc.tables[1].set_heading_row(1)                          # row 1 repeats on every page
    doc.tables[1].autofit("content")                          # fit columns to cells (or "window"/"fixed")
    doc.tables[2].delete()

records = doc.tables[1].records()   # body rows as [{header: value}, …] — read, no edit scope

with doc.edit("Style a table"):
    t = doc.tables[1]
    t.set_style("Grid Table 4 - Accent 1")                   # restyle FIRST (it overwrites cell shading)…
    t.set_alignment("center")                                # …then the whole table across the page
    t.set_borders(sides=["box", "horizontal", "vertical"])   # whole grid in one call
    t.set_banding(first_row=True, banded_rows=True)          # Table Style Options (needs a style)
    t.row(1).set_shading(fill="#2E86C1")                     # whole row: table:1:row:1
    t.row(1).format_run(bold=True, color="white")
    t.column(2).format_paragraph(alignment="right")          # whole column: table:1:col:2
    t.cell(1, 1).set_vertical_alignment("center")            # cell:  table:1:1:1
```

`add_table`/`insert_table` default to the `Table Grid` style (visible borders);
`data` is a row-major 2-D array **or** records (a list of dicts whose keys
become a header row), and `rows`/`cols` are inferred from `data` when omitted.
A cell *is* an anchor (`table:N:R:C`), so it takes `set_text`/`apply_style`/etc.
Treat a table as records keyed by row 1: `records()` reads it back,
`append_record({...})` adds a row by header name, `update_row(key, {...},
column=None)` sets cells on the first row whose key-column equals `key`.
**Styling:** `set_style` restyles an existing table (do it *before* any cell
shading — a style reapply overwrites direct cell colours), `set_alignment`
positions it on the page, `set_borders` rules the whole grid, `set_banding`
toggles the Table Style Options (needs a real style applied). A whole **row**
(`table:N:row:R` / `t.row(R)`) or **column** (`table:N:col:C` / `t.column(C)`) is
an anchor too, so the `set_shading`/`set_borders`/`apply_style`/`format_run`/
`format_paragraph` verbs style the strip in one call. A column op on a table with
merged / mixed-width cells raises `OpError` — style those cells via `table:N:R:C`.
`Cell.set_vertical_alignment("top"|"center"|"bottom")` sets vertical alignment.

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

**See what you recorded.** Read the revisions structurally —
`doc.revisions.list()` → `[{index, type, author, text, anchor_id, …}]` (`type` is
`insert`/`delete`/`format`/…) — or *visually* with `doc.snapshot(markup="all")`.
A plain `anchor.text` read is the **final** view (inserted runs present, deleted
runs gone), so to recover the *original* wording use the revision-aware reads:

```python
para.text_final        # as if every tracked change in it were accepted
para.text_original     # as if rejected (the deleted wording, restored)
para.revision_segments()  # [{text, change}] — change is insert/delete/None
```

(A tracked `find_replace` on the **same** paragraph as a previous one can still
drift — re-read between tracked edits, or use `text_final` to see the result.)

**Resolve them.** Accept or reject changes once you (or the user) are happy:

```python
with doc.edit("Resolve review"):
    doc.revisions[2].accept()                       # or .reject() — one change
    doc.revisions.accept_all()                      # the whole document
    doc.revisions.reject_all(within=doc.heading("Risks"))  # just one section's range
```

`accept`/`reject` renumber the rest (re-list between them); `accept_all`/
`reject_all` return the count resolved. An anchor's range is *literal* — a heading
covers only its line, so scope `within=` to a range/paragraph that spans the
changes.

### Publishing flourishes — watermark & pull quote

```python
with doc.edit("Stamp draft"):
    doc.set_watermark("DRAFT", layout="diagonal")   # WordArt behind every page
    doc.heading("Summary").insert_text_box(          # a floating pull quote
        "Key takeaway.", width="2.5in", fill="#eeeeff", wrap="square")
# doc.remove_watermark() clears it (idempotent; setting again replaces).
```

`set_watermark` draws into each section's header story (replacing a prior one);
`insert_text_box` floats a `Shapes.AddTextbox` anchored to the anchor's paragraph
(`wrap` is the `insert_image` vocabulary minus `inline`).

### Document info — metadata, variables, links, fields, proofing

```python
doc.properties.read()                # {"builtin": {Title, Author, …}, "custom": {…}}
with doc.edit("Set metadata"):
    doc.properties.set("Title", "Q3 Report")        # a built-in property
    doc.properties.set("Project", "Apollo", custom=True)  # a custom one (created if absent)
    doc.variables.set("ClientName", "Acme")         # invisible { DOCVARIABLE } storage
doc.variables.list()                 # {"ClientName": "Acme"}
doc.hyperlinks.list()                # read mirror of link_to: [{text, address, sub_address, anchor_id, …}]
doc.hyperlinks[1].update(address="https://new.example", text="New")  # retarget/relabel a link in place (1-based index)
doc.fields.list()                    # read mirror of insert_field: [{kind, code, result, anchor_id, …}]
doc.proofing()                       # {spelling:{count,errors}, grammar:{…}, readability:{flesch_reading_ease,…}}
doc.lint(within="heading:3")         # audit: [{rule, kind, severity, anchor_id, message, fixable, fix, …}] (pure read)
doc.lint(rules={"exclude": ["mixed-run-format"]})  # rules=None → default set; list to include; {"exclude":[…]} to drop
doc.regularize(within="heading:3", dry_run=True)   # plan the fixable findings (no write)
doc.regularize()                     # apply them in one atomic-undo → {applied, skipped, findings}; idempotent

cp = doc.checkpoint()                 # opaque structural fingerprint NOW (pure read; include=text|text+style|text+format)
# … agent or user edits …
doc.changes_since(cp)                 # content-aligned change list vs now: [{op, anchor_id, text_before/after, …}]
{c["anchor_id"] for c in doc.changes_since(cp) if "anchor_id" in c}  # verify exactly the paras I meant changed
doc.diff(cp_a, cp_b)                  # diff two stored checkpoints; cp.to_json()/Checkpoint.from_json() persist a token
```

`doc.checkpoint()` fingerprints the document's structure (pure read — no event
fires when content changes, so checkpoint+diff is the *only* way to answer "what
changed in session" and to verify your own edits landed). `changes_since(cp)`
diffs a stored checkpoint against the document now, `diff(a, b)` diffs two stored
checkpoints; each change is `replace`/`insert`/`delete`/`restyle`/`reformat`
carrying the **current** `para:N`, aligned by paragraph content (not index — ids
renumber), with an unchanged doc returning `[]`. `include="text+format"` also
catches a direct-formatting edit as a `reformat`.

`doc.properties` (read/write) and `doc.variables` (read/write) manage the file's
metadata and named storage; `doc.hyperlinks` and `doc.fields` are discovery
collections — the read mirrors of `link_to` / `insert_field` — and a
`doc.hyperlinks[N]` is editable in place via `update` / `set_address` / `set_text`
(fields stay read-only).
`doc.proofing()` runs Word's spelling/grammar/readability tools (a pure read, but
heavier than `stats()` — it (re)checks the document). `doc.lint()` audits
formatting/structural/policy issues (a severity-ranked list; each finding's `fix`
is an exec op when `fixable`), and `doc.regularize()` applies the fixable ones in
one `doc.edit("Regularize formatting")` — idempotent (each writes the style's own
value back), Track-Changes-aware, `dry_run=True` to preview. Wrap the writes in
`doc.edit(...)` for atomic undo.

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
doc.snapshot("review.png", markup="all")         # show tracked changes / comments as marks
doc.snapshot(max_dim=1000)                        # whole doc, cheap layout check (see below)
png = doc.heading("Introduction").snapshot()[0].png   # an anchor's page(s); heading → section
```

Returns a list of `Snapshot` (`.page`, `.png` bytes, `.path`). Word exports a
pixel-faithful PDF and wordlive rasterises it — real fonts, spacing, geometry.
Read-only. Needs the optional `snapshot` extra (PyMuPDF), else `SnapshotError`.

To **check styling across a whole multi-page doc cheaply**, render every page
(no `pages=`) with `max_dim=N` — it caps each page's long edge to `N` pixels.
A vision model is billed on pixel area, so the cap is a predictable per-page
token budget (~1000 stays legible); without it, full-resolution pages are
several times the tokens.

## Saving — persist or hand back a deliverable (ungated in Python)

```python
doc.saved                          # False once there are unsaved edits
doc.save()                         # save to the existing file (errors if never saved)
doc.save_as("out/report.docx")     # save a .docx (overwrite=False by default)
doc.export_pdf("out/report.pdf")   # a pixel-faithful PDF; from_page/to_page limit the range
```

The Python API is **trusted and ungated** — it writes wherever you point it.
(The CLI / MCP surfaces gate these behind a `WORDLIVE_SAVE_DIRS` whitelist, since
their input can be prompt-injected.) `export_pdf` is the recommended deliverable
path — same engine as `snapshot`. `save_as` writes `.docx`; PDF is `export_pdf`.

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
| `PathNotAllowedError` | a save/image path was refused by the CLI/MCP whitelist (Python API never raises it) | no — configure a save/image dir |
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
