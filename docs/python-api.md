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

Every anchor type inherits `apply_style(name)` and `format_paragraph(...)` from
[`Anchor`](#wordlive.Anchor), so the same calls work uniformly on bookmarks,
content controls, headings, and any future anchor types.

::: wordlive.Anchor

::: wordlive.Bookmark

::: wordlive.ContentControl

::: wordlive.Heading

::: wordlive.HeadingCollection

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

## Editing

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
