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
`insert_paragraph_before/after(...)`, `insert_image(...)`, `insert_table(...)`,
`insert_break(...)`, `insert_field(...)`, and the list verbs (`apply_list`, `remove_list`,
`list_info`, `restart_numbering`, `indent_list`, `outdent_list`) from
[`Anchor`](#wordlive.Anchor), so the same calls work uniformly on bookmarks,
content controls, headings, paragraphs, table cells, header/footer ranges, and
arbitrary range anchors. `insert_image` accepts a file path, raw bytes, or a
base64 string and embeds the picture; `wrap` is required (`"inline"`, `"auto"`,
or a float wrap like `"square"`/`"top-bottom"`), and `block=True` places the
image on its own new line rather than in the anchor's text run.
`insert_table(rows, cols, …)`
creates a new table at the anchor and returns its [`Table`](#wordlive.Table)
(append at the end with [`Document.add_table`](#wordlive.Document)).
`insert_break(kind="page"|"column"|"section_next"|"section_continuous")` drops
an explicit break; for a reflow-safe page break tied to a paragraph (e.g. every
`Heading 1`), pass `page_break_before=True` to `format_paragraph` instead.
`format_run(...)` sets character formatting (bold/italic/underline, `font`,
`size`, `color`, `highlight`, sub/superscript, caps, `spacing`) — the run-level
layer, ideal with a `range:START-END` anchor to style a phrase. `set_shading`,
`set_borders`, and `add_tab_stop` add range/cell fill, borders, and tab stops;
colours accept a name, hex, or `(r, g, b)` and sizes/positions accept points or
a unit string (`"12pt"`, `"1in"`). `insert_field(kind, ...)` drops a
self-updating field (`"page"`, `"numpages"`, `"date"`, …, or `"field"` + a raw
code) — pair it with a footer for page numbers and refresh with
[`Document.update_fields()`](#wordlive.Document). Every anchor also has `snapshot(...)`, which
renders the page(s) it sits on to PNG (a heading expands to its whole section) —
see [Snapshots](#snapshots).

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
`Snapshot` (one per page) and optionally write the image(s) to `out`. This needs
the optional `snapshot` extra (PyMuPDF); a missing backend raises
[`SnapshotError`](#wordlive.SnapshotError).

```python
import wordlive as wl

with wl.attach() as word:
    doc = word.documents.active
    png = doc.heading("Introduction").snapshot()[0].png   # bytes for a model
    doc.snapshot("report.png", pages=(1, 3))              # write pages 1-3
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

::: wordlive.SnapshotError

::: wordlive.WordBusyError

::: wordlive.ComError
