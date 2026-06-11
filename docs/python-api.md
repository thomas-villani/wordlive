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
`insert_paragraph_before/after(...)`, `insert_block(...)`, `insert_image(...)`, `insert_table(...)`,
`insert_break(...)`, `insert_field(...)`, `insert_footnote(...)`,
`insert_endnote(...)`, `insert_toc(...)`, `link_to(...)`,
`insert_cross_reference(...)`, `insert_caption(...)`, and the list verbs (`apply_list`, `remove_list`,
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
a unit string (`"12pt"`, `"1in"`). `insert_field(kind, ...)` drops a
self-updating field (`"page"`, `"numpages"`, `"date"`, …, or `"field"` + a raw
code) — pair it with a footer for page numbers and refresh with
[`Document.update_fields()`](#wordlive.Document). `insert_footnote(text)` /
`insert_endnote(text)` attach a note to the anchor's range and return a
[`Footnote`](#wordlive.Footnote) / [`Endnote`](#wordlive.Endnote) (addressed
`footnote:N` / `endnote:N`); `insert_toc(levels=(1, 3), …)` inserts a table of
contents and returns a [`Toc`](#wordlive.Toc) — see
[Footnotes, endnotes & TOC](#footnotes-endnotes-toc). `link_to(address=… |
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

::: wordlive.Footnote

::: wordlive.Endnote

::: wordlive.FootnoteCollection

::: wordlive.EndnoteCollection

::: wordlive.Toc

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

`Document.revisions` is a read-only [`RevisionCollection`](#wordlive.RevisionCollection)
that reads those tracked changes back as structured data — the way to *see* what
a tracked batch recorded (plain text reads concatenate the inserted and deleted
runs). `revisions.list()` reports each change as
`{index, type, author, text, anchor_id, start, end, date}`, where `type` is
`"insert"` / `"delete"` / `"format"` / … . The visual counterpart is
`snapshot(markup="all")` (see [Snapshots](#snapshots)).

::: wordlive.RevisionCollection

::: wordlive.Revision

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

::: wordlive.PathNotAllowedError

::: wordlive.SnapshotError

::: wordlive.WordBusyError

::: wordlive.ComError
