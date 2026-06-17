# Python API

Every entry on this page is generated from the docstrings in the
[`wordlive`](https://github.com/thomas-villani/wordlive/tree/main/src/wordlive)
package, so it stays in sync with the code. If something looks thin, the fix
is in the source docstring, not here.

The public surface is small on purpose. Three rough layers:

- **Connect** — [`attach`](#wordlive.attach) / [`connect`](#wordlive.connect)
  return a [`Word`](#wordlive.Word) handle.
- **Address** — [`Document`](#wordlive.Document) exposes
  [`Bookmark`](#wordlive.Bookmark), [`ContentControl`](#wordlive.ContentControl),
  and [`Heading`](#wordlive.Heading) anchors, plus
  [`anchor_by_id`](#wordlive.Document) for unified addressing.
- **Mutate** — wrap writes in [`Document.edit()`](#wordlive.Document) →
  [`EditScope`](#wordlive.EditScope) for atomic undo and Selection
  preservation.

The package version is available as `wordlive.__version__` (resolved from the
installed package metadata).

See [Concepts](concepts.md) for the *why* behind these shapes.

---

## Connecting to Word

::: wordlive.attach

::: wordlive.connect

::: wordlive.Word

## Documents

::: wordlive.Document

::: wordlive.DocumentCollection

## Anchors

Every anchor type inherits `apply_style(name)`, `format_paragraph(...)`,
`format_run(...)`, `set_shading(...)`, `set_borders(...)`, `add_tab_stop(...)`,
`insert_paragraph_before/after(...)`, `insert_block(...)`, `insert_image(...)`,
`insert_text_box(...)`, `insert_table(...)`,
`insert_break(...)`, `insert_field(...)`, `insert_footnote(...)`,
`insert_endnote(...)`, `insert_toc(...)`, `insert_table_of_figures(...)`,
`mark_index_entry(...)`, `insert_index(...)`, `insert_citation(...)`,
`insert_bibliography(...)`, `mark_citation(...)`,
`insert_table_of_authorities(...)`, `insert_content_control(...)`,
`link_to(...)`,
`insert_cross_reference(...)`, `insert_caption(...)`, the revision-aware reads
(`text_final`, `text_original`, `revision_segments()` — see
[Track Changes](#track-changes)), and the list verbs (`apply_list`, `remove_list`,
`list_info`, `restart_numbering`, `indent_list`, `outdent_list`) from
[`Anchor`](#wordlive.Anchor), so the same calls work uniformly on bookmarks,
content controls, headings, paragraphs, table cells, header/footer ranges, and
arbitrary range anchors. `insert_image` accepts a file path, raw bytes, or a
base64 string and embeds the picture; `wrap` is required (`"inline"`, `"auto"`,
or a float wrap like `"square"`/`"top-bottom"`), and `block=True` places the
image on its own new line rather than in the anchor's text run. The read mirror
is `read_image()`, which returns `(bytes, mime_type)` for the single picture in
the anchor's range — see [Images](#images).
`insert_block(items, where="after")` inserts a contiguous run of styled
paragraphs in one op (each item a plain string or `{text | runs, style?}`, where
`text` carries `**bold**`/`*italic*` markdown and `runs` is the structured
`[{text, bold?, italic?, underline?, style?}]` form) and returns a
[`RangeAnchor`](#wordlive.RangeAnchor) spanning the block — feed it straight into
`apply_list` to bullet the section. Two opinionated macros build on it:
`insert_section(heading, body, *, level=1, where="after")` places a
`Heading {level}` paragraph plus its body (the same items shape, or a bare
string) in one op, and `insert_markdown(md, *, where="after")` maps a
**constrained-Markdown subset** — `#`/`##`/`###` headings, `-`/`*` bullets, `1.`
numbers, blank-line paragraphs, inline `**bold**`/`*italic*` — to real Word
structure (not CommonMark: no code fences, nested lists, or tables in v1).
Headings additionally have `replace_section_body(body, *, markdown=False)`, which
clears the body under a heading (up to the next same-or-higher heading) and
inserts a replacement, keeping the heading — the "rewrite section X" workflow.
All three return the new content's [`RangeAnchor`](#wordlive.RangeAnchor).

```python
a = doc.headings["Methods"]
a.insert_section("Results", ["We saw a **20%** lift.", "Caveats apply."], level=2)
a.insert_markdown("# Plan\n\nKick-off.\n\n- scope it\n- staff it")
a.replace_section_body("Updated findings.\n\n- point one\n- point two", markdown=True)
```

`insert_table(rows, cols, …)`
creates a new table at the anchor and returns its [`Table`](#wordlive.Table)
(append at the end with [`Document.add_table`](#wordlive.Document)); pass `data`
as a 2-D array or as records (a list of dicts whose keys become a header row),
and `rows`/`cols` are inferred from `data` when omitted.
`insert_break(kind="page"|"column"|"section_next"|"section_continuous")` drops
an explicit break; for a reflow-safe page break tied to a paragraph (e.g. every
`Heading 1`), pass `page_break_before=True` to `format_paragraph` instead.
`format_paragraph` also takes `line_spacing` (the leading within a paragraph: a
multiple like `1.5`, the keywords `"single"`/`"1.5"`/`"double"`, or an exact
length such as `"14pt"`) and the pagination controls `keep_together`,
`keep_with_next`, and `widow_control` (tri-state booleans) for clean multi-page
layout.
`format_run(...)` sets character formatting (bold/italic/underline, `font`,
`size`, `color`, `highlight`, sub/superscript, caps, `spacing`) — the run-level
layer, ideal with a `range:START-END` anchor to style a phrase. `set_shading`,
`set_borders`, and `add_tab_stop` add range/cell fill, borders, and tab stops;
colours accept a name, hex, or `(r, g, b)` and sizes/positions accept points or
a unit string (`"12pt"`, `"1in"`). `drop_cap(lines=3, position="dropped"|"margin"|"none", …)`
turns the first letter of the anchor's paragraph into a real Word drop cap (the
editorial oversized initial; `position="none"` removes one). `insert_field(kind, ...)` drops a
self-updating field (`"page"`, `"numpages"`, `"date"`, …, or `"field"` + a raw
code) — pair it with a footer for page numbers and refresh with
[`Document.update_fields()`](#wordlive.Document). `insert_footnote(text)` /
`insert_endnote(text)` attach a note to the anchor's range and return a
[`Footnote`](#wordlive.Footnote) / [`Endnote`](#wordlive.Endnote) (addressed
`footnote:N` / `endnote:N`); `insert_toc(levels=(1, 3), …)` inserts a table of
contents and returns a [`Toc`](#wordlive.Toc), `insert_table_of_figures(label=
"Figure", …)` lists the captions of one label as a [`TableOfFigures`](#wordlive.TableOfFigures),
and `mark_index_entry(entry, …)` + `insert_index(…)` mark and build a back-of-book
[`Index`](#wordlive.Index). `insert_citation(tag, …)` cites a registered source and
`insert_bibliography(…)` builds the works-cited block, while `mark_citation(
long_citation, …)` + `insert_table_of_authorities(…)` mark and build a
[`TableOfAuthorities`](#wordlive.TableOfAuthorities) — see
[Footnotes, endnotes & TOC](#footnotes-endnotes-toc).
`insert_content_control(kind="rich_text", …)` wraps the anchor's range in a new
content control (see [Anchoring & linking](#anchoring-linking)). `link_to(address=… |
bookmark=…)` makes the anchor a hyperlink, `insert_cross_reference(target, …)`
references another anchor, and `insert_caption(label, …)` adds a numbered
caption — see [Anchoring & linking](#anchoring-linking). Every anchor also has `snapshot(...)`, which
renders the page(s) it sits on to PNG (a heading expands to its whole section) —
see [Snapshots](#snapshots).
`location()` is the non-visual companion: it returns `{page, end_page, line,
column, in_table}` — where the anchor sits in the laid-out document (its page
span, and its first character's line/column) — so an agent can answer "what page
is this on" without a snapshot. It repaginates first (content-neutral; selection
and view untouched), so page numbers are print-layout truth.

::: wordlive.Anchor

::: wordlive.Bookmark

::: wordlive.ContentControl

::: wordlive.Heading

::: wordlive.HeadingCollection

::: wordlive.Paragraph

::: wordlive.ParagraphCollection

::: wordlive.RangeAnchor

::: wordlive.StartAnchor

::: wordlive.EndAnchor

## Images

The read side of the image story (the write side is
[`Anchor.insert_image`](#wordlive.Anchor)). `doc.images` is a read-only
discovery collection over the document's embedded pictures; its `list()` reports
each image's `image:N` id, MIME type, size (points), alt text, and the `para:N`
it sits in. Index it (`doc.images[2]`) for an [`ImageAnchor`](#wordlive.ImageAnchor),
then call [`read_image()`](#wordlive.Anchor) for the raw bytes + MIME — the path
for handing an embedded picture to a vision model. `read_image()` also works on
any anchor whose range contains exactly one picture (e.g. `doc.paragraphs[7]`);
a range with no image, or more than one, raises
[`ImageSourceError`](#wordlive.ImageSourceError). Extraction is non-mutating, so
it needs no `doc.edit(...)`.

::: wordlive.ImageAnchor

::: wordlive.ImageCollection

## Watermarks & text boxes

The floating-shape flourishes, both documented on their host classes above.
`Document.set_watermark(text, …)` stamps a WordArt text watermark
(DRAFT / CONFIDENTIAL) behind every page via each section's header story —
`layout="diagonal"`/`"horizontal"`, `color`, `font`, `semitransparent`; it
replaces any prior text watermark rather than stacking, and
`Document.remove_watermark()` clears it (idempotent). `Anchor.insert_text_box(text, …)`
drops a floating text box / pull quote anchored to any anchor's paragraph, with
`width`/`height` (points or unit strings), `wrap` (the `insert_image` vocabulary
minus `"inline"`), `where`, the text-format kwargs, and `fill`/`border`. Both are
edits — wrap in `doc.edit(...)` for atomic undo. Floating shapes are off the
anchor model (no `textbox:N` id); for precise positioning reach the shape via
`.com` on `doc.com.Shapes`.

## Equations

Mathematical equations as first-class anchors. The write side is
[`Anchor.insert_equation`](#wordlive.Anchor): it takes exactly one of three input
dialects — `unicodemath=` (Word's native linear form, e.g. `"a^2+b^2=c^2"`,
zero-dependency), `latex=` (the optional `latex` extra does the LaTeX→MathML
hop), or `mathml=` (a `<math>` string) — converts it to Office Math, and places
it on its own paragraph with a pinned style so it never inherits a neighbouring
heading's style: `display=True` gives it the dedicated centred `Equation`
paragraph style (created on first use, based on `Normal`); `display=False`
resets the paragraph to `Normal` and left-aligns it (still its own paragraph,
not mid-sentence). It returns an [`EquationAnchor`](#wordlive.EquationAnchor)
addressed `equation:N` — a *positional* id in Word's `OMaths` order, so
inserting another equation before it renumbers it (re-list rather than caching
the id across further inserts). LaTeX and MathML travel LaTeX→MathML→OMML→Word through Office's own
shipped XSLT (`MML2OMML.XSL`), so only the LaTeX→MathML step needs a third-party
library; malformed input or a missing backend raises
[`EquationError`](#wordlive.EquationError).

`doc.equations` is the read side: a discovery collection whose `list()` reports
each equation's `equation:N` id, `type` (`display`/`inline`), a linear preview,
and the `para:N` it sits in. Index it (`doc.equations[2]`) for an
[`EquationAnchor`](#wordlive.EquationAnchor), then read `equation.mathml` (a
non-mutating round-trip back to MathML via Office's `OMML2MML.XSL`) or
`equation.linear`. An equation has no plain text, so `set_text` raises — delete
and re-insert to change it.

::: wordlive.EquationAnchor

::: wordlive.EquationCollection

## Charts

Excel-backed charts as first-class anchors. The write side is
[`Anchor.insert_chart`](#wordlive.Anchor): `kind` is `"bar"` (clustered
columns), `"pie"`, `"line"`, or `"scatter"`, and `data` is either a `{label:
value}` mapping (for bar/pie/line) or an array of `[x, y]` pairs (for `scatter` —
both axes numeric, with duplicate/clustered x preserved as distinct points; line
accepts either). `title=` sets the chart title and series name. It returns a
[`ChartAnchor`](#wordlive.ChartAnchor) addressed `chart:N` — a *positional* id in
document order, so inserting another chart earlier renumbers it.

Charts embed a chart via `InlineShapes.AddChart2`, whose data lives in a hidden
Excel workbook — so **Excel must be installed**. A non-invasive registry probe
gates the insert and raises [`ExcelNotAvailableError`](#wordlive.ExcelNotAvailableError)
(CLI exit 6) *before* touching the document if Excel is absent. After populating
the data wordlive **breaks the data link**, so the chart's data is static: no
embedded workbook ships in the document, and the series data isn't read back
(which keeps the hidden Excel from orphaning). The Python API is ungated; the
CLI/MCP surfaces add the same Excel probe.

`doc.charts` is the read side: a discovery collection whose `list()` reports each
chart's `chart:N` id, `kind`, `title`, and the `para:N` it sits in (metadata
only). Index it (`doc.charts[2]`) for a [`ChartAnchor`](#wordlive.ChartAnchor),
then read `chart.chart_type` / `chart.title`. A chart has no plain text, so
`set_text` raises — delete and re-insert to change it.

::: wordlive.ChartAnchor

::: wordlive.ChartCollection

## Footnotes, endnotes & TOC

Notes and the table of contents are reference structures built from anchors.
`anchor.insert_footnote(text)` / `insert_endnote(text)` drop a reference mark at
the anchor and put the body in the note story; they return a
[`Footnote`](#wordlive.Footnote) / [`Endnote`](#wordlive.Endnote) addressed
`footnote:N` / `endnote:N`, so `note.set_text(...)` edits the body and
`note.delete()` removes the mark and body together. Discover existing notes with
`doc.footnotes` / `doc.endnotes` (read-only collections whose `list()` reports
each note's number, text, and the `para:N` it's anchored at).

`anchor.insert_toc(levels=(1, 3), use_heading_styles=True, hyperlinks=True)`
inserts a table of contents over the document's headings and returns a
[`Toc`](#wordlive.Toc); `doc.add_toc(...)` is the sugar for placing one at the
document start. A TOC's page numbers only populate after repagination — call
`toc.update()`, [`Document.update_fields()`](#wordlive.Document), or take a
`snapshot` (which forces print layout) before reading them.

`anchor.insert_table_of_figures(label="Figure", include_label=True, hyperlinks=True,
right_align_page_numbers=True)` is the same field-block pattern over the captions
wordlive inserts: it lists every caption of one `label` (`"Figure"`/`"Table"`/
`"Equation"`/custom) with its page number and returns a
[`TableOfFigures`](#wordlive.TableOfFigures) with `update()` / `update_page_numbers()`.

A back-of-book index is two steps. `anchor.mark_index_entry(entry,
cross_reference=…, bold=…, italic=…)` marks the anchor's range as an `XE` index
field — `entry` uses `"main:sub"` to nest a subentry — then
`anchor.insert_index(columns=2, run_in=False, right_align_page_numbers=False)`
builds the index from those marks and returns an [`Index`](#wordlive.Index);
`doc.add_index(...)` is the sugar for one at the document end. Like the TOC, the
`TableOfFigures` and `Index` are field blocks: their page numbers populate only
after repagination (`update()`, `update_fields()`, or a `snapshot`).

Citations and a bibliography are a source-then-cite-then-build workflow.
`doc.sources` is a `SourceCollection` over the
document's master source list: `doc.sources.add(source_type="book", author=…,
title=…, year=…, …)` registers a [`Source`](#wordlive.Source) (`source_type` is
`"book"` / `"journal_article"` / `"web_site"` / `"case"` / … — `author` is
`"Last, First"` or a list, and `tag` auto-derives from the first author's surname +
year when omitted), `doc.sources.add_xml("<b:Source>…")` is the raw-OOXML escape
hatch, and the collection is list/index/`in`/`len`-able by tag. `doc.bibliography_style`
is the read/write style property (`"APA"` / `"MLA"` / `"Chicago"` / `"IEEE"` /
`"Turabian"`, build-dependent — an unsupported value raises
[`OpError`](#wordlive.OpError)). `anchor.insert_citation(tag, pages=…,
suppress_author=…, …)` inserts an in-text `CITATION` field rendering per that
style (e.g. `(Smith 2020, 15)`) and returns a [`Citation`](#wordlive.Citation) — a
tag with no registered source still inserts but renders *"Invalid source
specified."*. `anchor.insert_bibliography()` inserts the works-cited block and
returns a [`Bibliography`](#wordlive.Bibliography); `doc.add_bibliography()` is the
sugar for one at the document end.

A table of authorities (the legal-citation index) is the same two-step,
mark-then-build pattern as the back-of-book index.
`anchor.mark_citation(long_citation, short_citation=…, category="cases")` marks the
anchor's range as a `TA` field (`category` is `"cases"` / `"statutes"` / `"other"`
/ … or an int 1–16; `short_citation` defaults to `long_citation`), then
`anchor.insert_table_of_authorities(category="all", passim=True,
keep_entry_formatting=True)` builds the table from those marks and returns a
[`TableOfAuthorities`](#wordlive.TableOfAuthorities); `doc.add_table_of_authorities(...)`
is the sugar for one at the document end. Like the TOC, the `Bibliography` and
`TableOfAuthorities` are field blocks: their entries and page numbers populate only
after repagination (`update()`, `update_fields()`, or a `snapshot`). (Word's
table of authorities has no per-field page-number refresh, so unlike the TOC and
`TableOfFigures` there's no `update_page_numbers()` — use a full `update()`.)

The document **theme** is the document-wide brand primitive — the colour scheme,
font scheme, and effects the Design tab drives. `doc.theme` is a
[`DocumentTheme`](#wordlive.DocumentTheme): `doc.theme.apply("Facet")` applies a
whole theme by built-in name (see `doc.theme.list_available()`) or `.thmx` path;
`doc.theme.set_colors(scheme="Blue", accent1="#1A73E8", text1="navy")` loads a named
colour scheme and/or overrides individual brand colours (keys `text1` /
`background1` / `text2` / `background2` / `accent1`–`accent6` / `hyperlink` /
`followed_hyperlink`, values a colour name / hex / `(r, g, b)`); and
`doc.theme.set_fonts(scheme="Garamond", major="Arial", minor="Calibri")` sets the
heading/body fonts. `doc.theme.colors` / `.major_font` / `.minor_font` /
`.to_dict()` read the current theme back. Wrap theme mutations in `doc.edit(...)`.

::: wordlive.Footnote

::: wordlive.Endnote

::: wordlive.FootnoteCollection

::: wordlive.EndnoteCollection

::: wordlive.Toc

::: wordlive.TableOfFigures

::: wordlive.Index

::: wordlive.Source

::: wordlive.Citation

::: wordlive.Bibliography

::: wordlive.TableOfAuthorities

::: wordlive.DocumentTheme

## Durable handles (pins)

Positional `para:N` / `heading:N` ids renumber when a structural edit shifts the
document. When a positional anchor misses, `AnchorNotFoundError.hint` says why
(out-of-range vs body-text-not-a-heading, the paragraph count, the nearest
heading) and recommends pinning. `doc.pin(anchor, name=None)` (alias
`doc.stamp`) plants a hidden bookmark over an anchor's range and returns a
`pin:<code>` id — random, or a readable slug via `name="budget-intro"` — that
Word keeps attached to the same content across inserts / deletes / edits;
resolve it with `doc.anchor_by_id("pin:…")` like any anchor (a deleted target's
pin vanishes). `doc.pin_outline(levels=…)` pins every heading in one call and
returns the `{anchor_id: pin}` map (idempotent — reuses a heading's existing
handle), and `doc.outline(pin=True)` adds a `pin` to each outline row. Wrap pin
calls in `doc.edit(...)` for atomic undo. In an `exec` batch, `bind: "name"` on
an insert op mints a pin on the new content, and `$ops[N].field` references an
earlier op's output (see [CLI](cli.md#exec) / [MCP](mcp.md#batches)). The methods
are [`Document.pin`](#wordlive.Document) / `stamp`, `pin_outline`, and
`outline(pin=…)`.

## Anchoring & linking

Create a named anchor, then point at it. `doc.bookmarks.add(name, anchor)`
creates a bookmark over an anchor's range (the `name` is validated against
Word's rules first) — the prerequisite for internal navigation. `anchor.link_to(
address=…)` makes the anchor an external hyperlink (URL / `mailto:` / file
path); `anchor.link_to(bookmark=…)` makes it an internal jump to a bookmark.
With `text=None` the anchor's existing range becomes the link; `text=…` instead
inserts new linked text at the end of the range (so a heading keeps its content). `anchor.insert_cross_reference(
target, kind=…)` inserts a reference to another anchor — `target` is a
`bookmark:NAME`, `heading:N`, `footnote:N`, or `endnote:N` id, and `kind` is
`"text"` / `"page"` / `"number"` / `"above_below"`. `anchor.insert_caption(
label="Figure", text=…, position=None)` adds an auto-numbered caption in its own
`Caption`-styled paragraph (never fused into the target); `position` is
`"above"`/`"below"`, defaulting to above for a `Table` and below otherwise, and
on a table cell the caption attaches to the whole table. Pair it with a
cross-reference for "see Figure 2". Cross-references and TOC/page-number fields
go stale when the document shifts — refresh them with
[`Document.update_fields()`](#wordlive.Document).

Content controls are the structured-document fill-in fields (the
read/write side is `doc.content_controls["NAME"]`). `anchor.insert_content_control(
kind="rich_text", title=…, tag=…, items=…, where="wrap", lock_contents=False,
lock_control=False)` creates one and returns the
[`ContentControl`](#wordlive.ContentControl): `where="wrap"` (default) surrounds the
anchor's existing range — e.g. a `range:START-END` from `find` — and `"before"` /
`"after"` insert a fresh empty control. `kind` is `rich_text` (default) / `text` /
`picture` / `combo_box` / `dropdown` / `date` / `checkbox` / `building_block` /
`group` / `repeating_section`; `items` (combo_box/dropdown only) is a list of
strings or `{"text": …, "value": …}` dicts; `lock_contents` stops edits to the
value and `lock_control` stops deletion. A `title` (or, failing that, a `tag`)
names the control so it's addressable later as `cc:TITLE`; the returned wrapper
works even unnamed. `doc.content_controls.add(anchor, kind=…, **kwargs)` takes an
`Anchor` or an anchor-id string.

::: wordlive.BookmarkCollection

## Styles

Styles are document-scoped handles. `Document.styles` is a
[`StyleCollection`](#wordlive.StyleCollection); apply styles to anchors via
[`Anchor.apply_style`](#wordlive.Anchor). Define a new style with
`doc.styles.add(name, type="paragraph", based_on=…, next_style=…)`, which returns
a writable [`Style`](#wordlive.Style): set its defaults with `style.format_run(…)`
/ `style.format_paragraph(…)` (the same kwargs as the anchor methods, minus
`highlight`) and chain styles via `style.base_style` / `style.next_paragraph_style`.
The brand/template workflow: `add` a house style once, then `apply_style` it
everywhere.

::: wordlive.Style

::: wordlive.StyleCollection

## Tables

`Document.tables` is a [`TableCollection`](#wordlive.TableCollection). Index a
table by 1-based position or `Title`, then read or edit it. A
[`Cell`](#wordlive.Cell) *is* an [`Anchor`](#wordlive.Anchor) — its id is
`table:N:R:C`, so `doc.anchor_by_id("table:1:2:3")` returns a cell that works
with `set_text`, `apply_style`, and `format_paragraph` like any other anchor.

Create tables with [`Document.add_table(rows, cols, …)`](#wordlive.Document)
(append at the end) or [`Anchor.insert_table(...)`](#wordlive.Anchor) (at any
position anchor); both return the new [`Table`](#wordlive.Table), populate cells
from a row-major `data` grid, default to the `Table Grid` style, and keep
appended tables from merging into an adjacent one. `Table.delete()` removes a
whole table — the structural mirror of `add_row` / `delete_row`.
`Table.set_heading_row(row=1, heading=True, allow_break=None)` marks a row as a
repeating header that reprints on every page the table spans.

Treat a table as **records** keyed by its header row (row 1) — the read/update
mirror of building one from `data=[{...}]`. `Table.records()` returns the body
rows as a list of `{header: cell_text}` dicts; `Table.append_record({...})`
appends a row from a dict (keys mapped to header columns, missing → empty, extra
→ ignored); `Table.update_row(key, {...}, column=None)` sets cells by header name
on the first row whose key-column (the first column, or the header named by
`column`) equals `key` — addressing a row by content instead of a fragile
1-based index.

::: wordlive.TableCollection

::: wordlive.Table

::: wordlive.Cell

## Comments

`Document.comments` is a [`CommentCollection`](#wordlive.CommentCollection).
`comments.add(anchor, text, author=...)` attaches a review comment to any
anchor's range *without changing the text* — the polite, side-channel way for
an agent to flag something. Existing comments are addressed by 1-based index
(`doc.comments[2]`) to `resolve()` or `delete()`.

::: wordlive.CommentCollection

::: wordlive.Comment

## Track Changes

`Document.tracked_changes()` is a context manager that turns Word's Track
Changes on for the scope and restores the prior setting on exit — pair it with
`edit()` to make a batch of edits *visibly*, as revisions the user can accept or
reject. `Document.track_changes` is the underlying read/write property for the
persistent flag. Both are documented on [`Document`](#wordlive.Document).

`Document.revisions` is a [`RevisionCollection`](#wordlive.RevisionCollection)
that reads those tracked changes back as structured data — the way to *see* what
a tracked batch recorded. `revisions.list()` reports each change as
`{index, type, author, text, anchor_id, start, end, date}`, where `type` is
`"insert"` / `"delete"` / `"format"` / … . The visual counterpart is
`snapshot(markup="all")` (see [Snapshots](#snapshots)).

Resolve them, too: `revisions[N].accept()` / `.reject()` make a single change
permanent / undo it (and renumber the rest), while
`revisions.accept_all(within=anchor)` / `reject_all(within=anchor)` do the whole
document — or just one anchor's range when `within` is given — and return the
count resolved.

For a read that separates a tracked edit's two sides, the [`Anchor`](#wordlive.Anchor)
helpers `text_final` (as if accepted), `text_original` (as if rejected), and
`revision_segments()` (the ordered `{text, change}` breakdown) reconstruct both:
Word's plain `text` read is the *final* view (inserted runs present, deleted runs
gone), so the original wording lives only on the delete revisions.

::: wordlive.RevisionCollection

::: wordlive.Revision

## Hyperlinks & fields

`Document.hyperlinks` and `Document.fields` are read-only discovery collections —
the read mirrors of [`Anchor.link_to`](#wordlive.Anchor) /
[`Anchor.insert_field`](#wordlive.Anchor). `doc.hyperlinks.list()` reports each
link's visible text, external `address` or internal `sub_address` bookmark,
screen tip, and a `range:START-END` / `para:N`; `doc.fields.list()` reports each
field's `kind` (the code's leading keyword — `PAGE` / `REF` / `TOC` / …), raw
`code`, rendered `result`, `locked`, and a `range:START-END` / `para:N`. Index
either (`doc.hyperlinks[2]`, `doc.fields[2]`) for the single-item wrapper.

::: wordlive.HyperlinkCollection

::: wordlive.Hyperlink

::: wordlive.FieldCollection

::: wordlive.Field

## Document metadata, variables & proofing

`Document.properties` is a read/write [`PropertyCollection`](#wordlive.PropertyCollection)
over the document's metadata: `read()` returns `{builtin, custom}` (the Title /
Author / Subject / Keywords / … bag plus any custom name/value pairs), `set(name,
value)` writes a built-in property, `set(name, value, custom=True)` a custom one,
and `delete(name)` removes a custom one. `Document.variables` is a
[`VariableCollection`](#wordlive.VariableCollection) over the invisible named
string storage that backs `{ DOCVARIABLE }` fields — `list()` returns
`{name: value}`; `set` / `get` / `delete` manage them. Wrap writes in
[`doc.edit(...)`](#wordlive.Document) for atomic undo.

`Document.proofing()` runs Word's proofing tools and returns
`{spelling, grammar, readability}`: spelling/grammar each give a `count` plus a
(capped) list of `{text, anchor_id, para}` for the flagged runs, and readability
gives the Flesch Reading Ease, Flesch-Kincaid Grade Level, passive-sentence %, and
averages. It's a pure read but heavier than [`stats()`](#wordlive.Document) — it
asks Word to (re)check the document. Documented on [`Document`](#wordlive.Document).

::: wordlive.PropertyCollection

::: wordlive.VariableCollection

## Lists & numbering

List operations apply to a *range's paragraphs*, so the verbs live on
[`Anchor`](#wordlive.Anchor) — `apply_list("numbered")`, `remove_list()`,
`list_info()`, `restart_numbering()`, and `indent_list()` / `outdent_list()`
work on any anchor. `Document.lists` is a read-only
[`ListCollection`](#wordlive.ListCollection) for discovering the lists already in
the document; index it (`doc.lists[2]`) to get a
[`RangeAnchor`](#wordlive.RangeAnchor) over a list's range.

::: wordlive.ListCollection

## Sections, headers & footers

`Document.sections` is a [`SectionCollection`](#wordlive.SectionCollection). Each
[`Section`](#wordlive.Section) reaches its headers and footers as
[`HeaderFooter`](#wordlive.HeaderFooter) anchors — `doc.sections[1].header()` /
`.footer("first")` — addressed `header:S:WHICH` / `footer:S:WHICH` (WHICH is
`primary` / `first` / `even`). A `HeaderFooter` *is* an `Anchor`, so
`set_text`, `apply_style`, and `format_paragraph` work on it like any other, plus
`insert_page_number()` sugar for a `{ PAGE }` field. `Section.set_page_setup(...)`
is the write mirror of `page_setup()` — margins, orientation, paper size, gutter,
and multi-column layout (`columns=N`), per section.

::: wordlive.SectionCollection

::: wordlive.Section

::: wordlive.HeaderFooter

## Editing

`Selection` is the explicit cursor surface: `doc.selection.info()` reads where
the cursor is, and `doc.selection.write(text, replace=...)` types at it.
`write` deliberately moves the cursor, so wrap it in
[`doc.edit()`](#wordlive.Document) and call
[`scope.allow_cursor_move()`](#wordlive.EditScope) for atomic undo without
snapping the cursor back. Everywhere else, prefer anchors over the cursor.

::: wordlive.EditScope

::: wordlive.Selection

::: wordlive.SelectionSnapshot

## Snapshots

[`Document.snapshot(...)`](#wordlive.Document) and
[`Anchor.snapshot(...)`](#wordlive.Anchor) render page(s) of the live document
to PNG so a vision model can *see* the layout — Word exports a pixel-faithful
PDF and wordlive rasterises the requested pages. `Document.snapshot` selects
pages (all, one, or a span); `Anchor.snapshot` (and
[`Document.snapshot_anchor`](#wordlive.Document)) renders the page(s) an anchor
occupies, expanding a heading to its whole section. Both return a list of
`Snapshot` (one per page) and optionally write the image(s) to `out`. Pass
`markup="all"` to render tracked changes and comments as visible revision marks
and balloons instead of the final document (the structured counterpart is
[`Document.revisions`](#wordlive.RevisionCollection)). `dpi` (default 150) sets
resolution; `max_dim` caps each page's long edge in pixels (only ever lowering
it) — the lever for a cheap *whole-document* layout check, since a vision model's
token cost scales with pixel area, so a long-edge cap is a predictable per-page
budget regardless of paper size (~1000 stays legible). This needs the optional
`snapshot` extra (PyMuPDF); a missing backend raises
[`SnapshotError`](#wordlive.SnapshotError).

```python
import wordlive as wl

with wl.attach() as word:
    doc = word.documents.active
    png = doc.heading("Introduction").snapshot()[0].png   # bytes for a model
    doc.snapshot("report.png", pages=(1, 3))              # write pages 1-3
    doc.snapshot("review.png", markup="all")              # show tracked changes
    shots = doc.snapshot(max_dim=1000)                    # whole doc, cheap layout check
```

::: wordlive.Snapshot

## Constants

`wordlive.constants` re-exports the typed `IntEnum` mirrors of the Word `Wd*`
magic numbers wordlive uses internally (alignment, break types, wrap types,
…). You rarely need these directly — the high-level API takes plain strings
(`"center"`, `"page"`, `"square"`) and maps them — but they're available for
`.com` escape-hatch code that talks to the raw object model.

```python
from wordlive import constants

constants.WdParagraphAlignment.CENTER   # 1
```

## Exceptions

::: wordlive.WordliveError

::: wordlive.WordNotRunningError

::: wordlive.DocumentNotFoundError

::: wordlive.AnchorNotFoundError

::: wordlive.StyleNotFoundError

::: wordlive.AmbiguousMatchError

::: wordlive.ReplaceVerificationError

::: wordlive.ImageSourceError

::: wordlive.OpError

::: wordlive.EquationError

::: wordlive.PathNotAllowedError

::: wordlive.SnapshotError

::: wordlive.WordBusyError

::: wordlive.ComError
