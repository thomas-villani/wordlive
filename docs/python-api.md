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
`insert_paragraph_before/after(...)`, and the list verbs (`apply_list`,
`remove_list`, `list_info`, `restart_numbering`, `indent_list`, `outdent_list`)
from [`Anchor`](#wordlive.Anchor), so the same calls work uniformly on
bookmarks, content controls, headings, paragraphs, table cells, header/footer
ranges, and arbitrary range anchors.

::: wordlive.Anchor

::: wordlive.Bookmark

::: wordlive.ContentControl

::: wordlive.Heading

::: wordlive.HeadingCollection

::: wordlive.Paragraph

::: wordlive.ParagraphCollection

::: wordlive.RangeAnchor

## Styles

Styles are document-scoped, read-only handles. `Document.styles` is a
[`StyleCollection`](#wordlive.StyleCollection); apply styles to anchors via
[`Anchor.apply_style`](#wordlive.Anchor).

::: wordlive.Style

::: wordlive.StyleCollection

## Tables

`Document.tables` is a [`TableCollection`](#wordlive.TableCollection). Index a
table by 1-based position or `Title`, then read or edit it. A
[`Cell`](#wordlive.Cell) *is* an [`Anchor`](#wordlive.Anchor) — its id is
`table:N:R:C`, so `doc.anchor_by_id("table:1:2:3")` returns a cell that works
with `set_text`, `apply_style`, and `format_paragraph` like any other anchor.

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
`set_text`, `apply_style`, and `format_paragraph` work on it like any other.

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

## Exceptions

::: wordlive.WordliveError

::: wordlive.WordNotRunningError

::: wordlive.DocumentNotFoundError

::: wordlive.AnchorNotFoundError

::: wordlive.StyleNotFoundError

::: wordlive.AmbiguousMatchError

::: wordlive.WordBusyError

::: wordlive.ComError
