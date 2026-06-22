"""Tables — collection, table wrapper, and cell anchors.

Tables are document-scoped collections (`doc.tables`), so they live here rather
than in `_anchors.py`. A `Cell` *is* an `Anchor`, though: it targets the cell's
COM `Range`, so the inherited `apply_style` / `format_paragraph` / `set_text`
machinery works on cells with no special-casing, and `replace --anchor-id
table:N:R:C` resolves through `Document.anchor_by_id` like any other anchor.

The anchor-id schemes are `table:N:R:C` (a single cell, 1-based table/row/
column), `table:N:row:R` (a whole row → [`RowAnchor`][wordlive.RowAnchor]), and
`table:N:col:C` (a whole column → [`ColumnAnchor`][wordlive.ColumnAnchor]). A row
or column anchor *is* an `Anchor`, so it styles the whole strip through the same
`shading` / `borders` / `apply-style` / `format-run` verbs a cell does. The bare
`table:N` form is *not* an anchor — a whole table is a structural collection, not
a single range — so it's addressed via `doc.tables[N]` and the `table` CLI group
instead (where whole-table restyle / borders / alignment / banding live).

Limitation: cell addressing assumes a rectangular grid. Tables with merged or
split cells have a non-uniform COM cell model; `Table.cell(r, c)` follows Word's
own `Table.Cell(row, col)` indexing and may raise inside merged regions. A
**column** anchor needs Word's per-column model, which a merged / mixed-width
table doesn't have — Word raises "mixed cell widths" — so a column op on such a
table raises `OpError` pointing back at per-cell `table:N:R:C` styling. Rows are
always a contiguous range, so `table:N:row:R` is unaffected.
"""

from __future__ import annotations

from collections.abc import Iterator
from typing import TYPE_CHECKING, Any

from . import _com
from ._anchors import Anchor, apply_borders, range_text
from .constants import WdAutoFitBehavior, WdCellVerticalAlignment, WdRowAlignment
from .exceptions import AnchorNotFoundError, ComError, OpError

# `autofit` mode keys -> WdAutoFitBehavior. "fixed" pins current widths,
# "content" sizes columns to their cell contents, "window" stretches to the page.
_AUTOFIT_MODES: dict[str, WdAutoFitBehavior] = {
    "fixed": WdAutoFitBehavior.FIXED,
    "content": WdAutoFitBehavior.CONTENT,
    "window": WdAutoFitBehavior.WINDOW,
}

# Whole-table alignment keys -> WdRowAlignment (the table across the page width).
_ROW_ALIGN: dict[str, WdRowAlignment] = {
    "left": WdRowAlignment.LEFT,
    "center": WdRowAlignment.CENTER,
    "centre": WdRowAlignment.CENTER,
    "right": WdRowAlignment.RIGHT,
}

# Cell vertical-alignment keys -> WdCellVerticalAlignment (0/1/3 — 2 is invalid).
_CELL_VALIGN: dict[str, WdCellVerticalAlignment] = {
    "top": WdCellVerticalAlignment.TOP,
    "center": WdCellVerticalAlignment.CENTER,
    "centre": WdCellVerticalAlignment.CENTER,
    "bottom": WdCellVerticalAlignment.BOTTOM,
}


def _coerce_align(value: Any, table: dict[str, Any], label: str) -> int:
    """Map an alignment keyword to its enum int, raising `OpError` on a miss."""
    key = str(value).strip().lower()
    if key not in table:
        choices = ", ".join(sorted(table))
        raise OpError(f"{label} must be one of {choices}; got {value!r}")
    return int(table[key])


if TYPE_CHECKING:
    from ._document import Document


def _strip_cell_text(raw: Any) -> str:
    """Cell text minus Word's trailing end-of-cell markers.

    A cell's `Range.Text` ends with CR + the cell mark (BEL, `\\x07`); a cell
    that contains multiple paragraphs repeats the pattern. `rstrip` of those
    code points gives the human-visible text, mirroring `paragraph_text`.
    """
    return str(raw or "").rstrip("\r\n\x07")


def index_of(doc_com: Any, table_com: Any) -> int:
    """1-based document position of `table_com` within `doc_com.Tables`.

    Tables can't overlap, so matching on `Range.Start` uniquely identifies the
    table even though Word hands back a fresh wrapper on each collection access
    (so identity comparison is unreliable). Used to report the index of a
    freshly-created table. Falls back to the table count if no start matches —
    which shouldn't happen for a table that was just added.
    """
    with _com.translate_com_errors():
        target = int(table_com.Range.Start)
        for i, t in enumerate(doc_com.Tables, start=1):
            try:
                if int(t.Range.Start) == target:
                    return i
            except Exception:
                continue
        return int(doc_com.Tables.Count)


class Cell(Anchor):
    """A single table cell, addressed by 1-based (row, column).

    Subclasses `Anchor`, so it inherits `insert_before` / `insert_after` /
    `delete` / `apply_style` / `format_paragraph` unchanged. Only the bits that
    differ for cells — the COM range, text read/write, and the anchor id — are
    overridden here.
    """

    kind = "cell"

    def __init__(self, table: Table, row: int, col: int, _com_cell: Any | None = None) -> None:
        super().__init__(table._doc, name=f"table:{table.index}:{row}:{col}")
        self._table = table
        self._row = row
        self._col = col
        # An already-resolved COM cell (e.g. from `Columns(C).Cells`), so callers
        # iterating a collection don't pay a `Table.Cell(r, c)` round-trip per cell.
        self._com_cell = _com_cell

    @property
    def anchor_id(self) -> str:
        return f"table:{self._table.index}:{self._row}:{self._col}"

    @property
    def row(self) -> int:
        return self._row

    @property
    def column(self) -> int:
        return self._col

    def _cell(self) -> Any:
        if self._com_cell is not None:
            return self._com_cell
        with _com.translate_com_errors():
            return self._table.com.Cell(self._row, self._col)

    def _range(self) -> Any:
        return self._cell().Range

    def _caption_object_range(self) -> Any:
        """Caption the *whole table*, not this cell.

        Returning the parent table's range lets `insert_caption` place a real
        standalone caption above / below the table (Word's native behaviour for
        a caption-able object), instead of fusing it into the cell — which also
        sidesteps the "end of a table row" COM error on the last cell.
        """
        return self._table.com.Range

    @property
    def text(self) -> str:
        with _com.translate_com_errors():
            return _strip_cell_text(range_text(self._cell().Range))

    def set_text(self, text: str) -> None:
        with _com.translate_com_errors():
            self._cell().Range.Text = text

    def set_vertical_alignment(self, align: str) -> None:
        """Set where this cell's content sits vertically — top, center, or bottom.

        `align` is ``"top"`` / ``"center"`` (``"centre"``) / ``"bottom"``, mapped
        onto `Cell.VerticalAlignment`. (Word shares this value space with page
        vertical alignment, whose ``2`` = *justify* slot a cell rejects, so only
        those three are offered.) Idempotent. Wrap in `doc.edit(...)` for atomic
        undo; bad input raises `OpError`.
        """
        value = _coerce_align(align, _CELL_VALIGN, "cell vertical alignment")
        with _com.translate_com_errors():
            self._cell().VerticalAlignment = value

    def merge(self, other: Cell) -> None:
        """Merge this cell with `other` into one cell spanning their rectangle.

        `other` must belong to the **same table**. Word joins the cells' text and
        collapses the rectangle into its **upper-left** cell, regardless of which
        corner is `self` vs `other` — so the merged cell is addressed by the
        upper-left coordinate of the spanned rectangle (e.g.
        ``cell(2, 2).merge(cell(1, 1))`` yields a cell at ``table:N:1:1``), and the
        other spanned coordinates stop resolving. The table becomes
        **non-uniform** (`Table.is_uniform` → `False`) — afterwards `table:N:R:C`
        indexes *physical* cells (a short row has fewer than `column_count`), so
        re-read the table to see the new shape. Wrap in `doc.edit(...)` for
        atomic undo; cross-table cells raise `OpError`.
        """
        if other._table.index != self._table.index:
            raise OpError(
                f"cannot merge {self.anchor_id} with {other.anchor_id}: "
                f"cells are in different tables"
            )
        with _com.translate_com_errors():
            # Positional MergeTo — the keyword is dropped under late binding.
            self._cell().Merge(other._cell())

    def split(self, rows: int = 1, cols: int = 2) -> None:
        """Split this cell into a `rows` × `cols` grid of cells.

        The inverse of `merge`; defaults to two side-by-side cells
        (`rows=1, cols=2`). Makes the table **non-uniform** (`Table.is_uniform`
        → `False`) — see `merge`. Wrap in `doc.edit(...)` for atomic undo; a
        count below 1 raises `OpError`.
        """
        if rows < 1 or cols < 1:
            raise OpError(f"split: rows and cols must be >= 1 (got rows={rows}, cols={cols})")
        with _com.translate_com_errors():
            # Positional NumRows, NumColumns — keywords are dropped under late binding.
            self._cell().Split(rows, cols)


class RowAnchor(Anchor):
    """A whole table row, addressed by ``table:N:row:R`` (1-based).

    Subclasses `Anchor` over the row's contiguous ``Rows(R).Range``, so the
    inherited styling verbs — `set_shading`, `set_borders`, `apply_style`,
    `format_run`, `format_paragraph` — restyle the *entire row* in one call
    (``shading --anchor-id table:1:row:1`` shades the header row;
    ``format-run --anchor-id table:1:row:1 --bold`` bolds it). `set_text` is
    refused — a row is a styling target, not a text slot; edit its cells via
    ``table:N:R:C``.
    """

    kind = "table-row"

    def __init__(self, table: Table, row: int) -> None:
        super().__init__(table._doc, name=f"table:{table.index}:row:{row}")
        self._table = table
        self._row = row

    @property
    def anchor_id(self) -> str:
        return f"table:{self._table.index}:row:{self._row}"

    @property
    def row(self) -> int:
        return self._row

    def _range(self) -> Any:
        with _com.translate_com_errors():
            return self._table.com.Rows(self._row).Range

    def set_text(self, text: str) -> None:
        raise OpError(
            f"{self.anchor_id} is a whole row; set text on its cells via "
            f"table:{self._table.index}:{self._row}:C"
        )


class ColumnAnchor(Anchor):
    """A whole table column, addressed by ``table:N:col:C`` (1-based).

    Unlike a row, a column is **not** a contiguous Word range — `Column.Range`
    isn't reachable under late binding — so this anchor styles the column by
    fanning each op out across its cells (`Columns(C).Cells`). The column-wide
    styling verbs (`set_shading`, `set_borders`, `apply_style`, `format_run`,
    `format_paragraph`) are overridden to loop the cells; range-only ops that need
    a single span (`set_text`, `replace`, `insert_*`) raise `OpError`.

    A table with **merged or mixed-width cells** has no per-column model — Word
    raises "mixed cell widths" — so any column op on such a table raises `OpError`
    pointing at per-cell ``table:N:R:C`` styling. (Rows are unaffected; use
    ``table:N:row:R``.)
    """

    kind = "table-column"

    def __init__(self, table: Table, col: int) -> None:
        super().__init__(table._doc, name=f"table:{table.index}:col:{col}")
        self._table = table
        self._col = col

    @property
    def anchor_id(self) -> str:
        return f"table:{self._table.index}:col:{self._col}"

    @property
    def column(self) -> int:
        return self._col

    def _range(self) -> Any:
        raise OpError(
            f"{self.anchor_id} is a whole column, not a single range; use the "
            f"column styling verbs, or address cells via table:{self._table.index}:R:{self._col}"
        )

    def _cells(self) -> list[Cell]:
        """The column's cells as `Cell` anchors — or `OpError` on a mixed-width table.

        Reading the column collection is what trips Word's "mixed cell widths"
        error on a merged / irregular table; we catch the resulting `ComError` and
        re-raise a clear `OpError` that points at per-cell styling. On a regular
        table this yields one `Cell` per body row, in row order.
        """
        try:
            with _com.translate_com_errors():
                com_cells = list(self._table.com.Columns(self._col).Cells)
        except ComError as e:
            raise OpError(
                f"cannot style column {self._col} of table {self._table.index} "
                f"({e}); the table has merged or mixed-width cells — style its "
                f"cells individually via table:{self._table.index}:R:{self._col}"
            ) from e
        # Wrap each already-read COM cell directly (passing it through as the
        # cached `_com_cell`) rather than re-resolving via `Table.cell()`, which
        # would re-round-trip COM per cell and bounds-check the physical row/col
        # against the *logical* row/column counts.
        return [Cell(self._table, int(c.RowIndex), self._col, _com_cell=c) for c in com_cells]

    def set_shading(self, *, fill: Any = None, pattern: Any = None) -> None:
        for cell in self._cells():
            cell.set_shading(fill=fill, pattern=pattern)

    def set_borders(
        self, *, sides: Any = "all", style: Any = "single", weight: Any = 0.5, color: Any = None
    ) -> None:
        for cell in self._cells():
            cell.set_borders(sides=sides, style=style, weight=weight, color=color)

    def apply_style(self, name: str) -> None:
        for cell in self._cells():
            cell.apply_style(name)

    def format_run(self, **kwargs: Any) -> None:
        for cell in self._cells():
            cell.format_run(**kwargs)

    def format_paragraph(self, **kwargs: Any) -> None:
        for cell in self._cells():
            cell.format_paragraph(**kwargs)

    def set_text(self, text: str) -> None:
        raise OpError(
            f"{self.anchor_id} is a whole column; set text on its cells via "
            f"table:{self._table.index}:R:{self._col}"
        )


class Table:
    """Wraps a Word `Table` COM object, located by its 1-based document position.

    The index is stored at construction (the collection knows it without a COM
    round-trip), so `anchor_id` and cell ids never have to re-scan the document.
    """

    def __init__(self, doc: Document, com: Any, index: int) -> None:
        self._doc = doc
        self._com = com
        self._index = index

    @property
    def com(self) -> Any:
        return self._com

    @property
    def index(self) -> int:
        return self._index

    @property
    def anchor_id(self) -> str:
        return f"table:{self._index}"

    @property
    def row_count(self) -> int:
        with _com.translate_com_errors():
            return int(self._com.Rows.Count)

    @property
    def column_count(self) -> int:
        with _com.translate_com_errors():
            return int(self._com.Columns.Count)

    @property
    def title(self) -> str:
        with _com.translate_com_errors():
            return str(self._com.Title or "")

    @property
    def is_uniform(self) -> bool:
        """Whether every row has the same physical cell count — a clean grid.

        `True` for a freshly built `R × C` table; `False` once a cell has been
        merged or split. On a non-uniform table `table:N:R:C` indexes *physical*
        cells (so an index can shift after a merge or fall off a short row),
        `delete_column` / column anchors raise "mixed cell widths", and
        `row_count` × `column_count` overstates the true cell count. Worth a
        check before addressing cells in a table you didn't build.
        """
        with _com.translate_com_errors():
            cols = int(self._com.Columns.Count)
            return all(
                int(self._com.Rows(r).Cells.Count) == cols
                for r in range(1, int(self._com.Rows.Count) + 1)
            )

    def cell(self, row: int, col: int) -> Cell:
        """Return the `Cell` at 1-based (row, col).

        Raises `AnchorNotFoundError` (kind `"table cell"`) if the coordinates
        fall outside the table's grid.
        """
        rows, cols = self.row_count, self.column_count
        if not (1 <= row <= rows and 1 <= col <= cols):
            raise AnchorNotFoundError("table cell", f"table:{self._index}:{row}:{col}")
        return Cell(self, row, col)

    def row(self, row: int) -> RowAnchor:
        """Return the `RowAnchor` for the 1-based `row` (`table:N:row:R`).

        A styling handle for the whole row — `table.row(1).set_shading(fill=…)`
        shades it, `.format_run(bold=True)` bolds it. Same object as
        `doc.anchor_by_id("table:N:row:R")`. Raises `AnchorNotFoundError` (kind
        `"table row"`) if out of range.
        """
        rows = self.row_count
        if not (1 <= row <= rows):
            raise AnchorNotFoundError("table row", f"table:{self._index}:row:{row}")
        return RowAnchor(self, row)

    def column(self, col: int) -> ColumnAnchor:
        """Return the `ColumnAnchor` for the 1-based `col` (`table:N:col:C`).

        The column counterpart of `row()` — `table.column(3).format_paragraph(
        alignment="right")` right-aligns a totals column. Same object as
        `doc.anchor_by_id("table:N:col:C")`. Raises `AnchorNotFoundError` (kind
        `"table column"`) if out of range. (A column op on a merged / mixed-width
        table raises `OpError` when applied — see `ColumnAnchor`.)
        """
        cols = self.column_count
        if not (1 <= col <= cols):
            raise AnchorNotFoundError("table column", f"table:{self._index}:col:{col}")
        return ColumnAnchor(self, col)

    def _row_cell_count(self, row: int) -> int:
        """The number of *physical* cells in a 1-based row (a merged row is shorter)."""
        with _com.translate_com_errors():
            return int(self._com.Rows(row).Cells.Count)

    def grid(self) -> list[list[str]]:
        """All cell text as a row-major `list[list[str]]`.

        Iterates each row's *physical* cells, so it stays safe on a merged /
        split table (a merged row simply yields fewer columns); on a uniform
        table it's the plain `row_count` × `column_count` grid.
        """
        return [
            [Cell(self, r, c).text for c in range(1, self._row_cell_count(r) + 1)]
            for r in range(1, self.row_count + 1)
        ]

    def read(self) -> dict[str, Any]:
        """Structured dump: metadata plus every cell with its addressable id.

        Each cell carries its `anchor_id` (`table:N:R:C`) so a caller can feed
        it straight back into `replace` / `style apply` / `format-paragraph`.
        Cells are walked **physically** per row (robust to merged / split
        tables); `uniform` reports whether `rows` × `columns` is the full grid.
        """
        rows, cols = self.row_count, self.column_count
        cells = [
            [
                {
                    "row": r,
                    "col": c,
                    "text": Cell(self, r, c).text,
                    "anchor_id": f"table:{self._index}:{r}:{c}",
                }
                for c in range(1, self._row_cell_count(r) + 1)
            ]
            for r in range(1, rows + 1)
        ]
        return {
            "index": self._index,
            "title": self.title,
            "rows": rows,
            "columns": cols,
            "uniform": self.is_uniform,
            "cells": cells,
        }

    def to_dict(self) -> dict[str, Any]:
        """Metadata only — `{index, title, rows, columns}`. Used by `table list`."""
        return {
            "index": self._index,
            "title": self.title,
            "rows": self.row_count,
            "columns": self.column_count,
        }

    def _header_names(self) -> list[str]:
        """Row 1's cell texts — the column labels the records API keys on."""
        return [self.cell(1, c).text for c in range(1, self.column_count + 1)]

    def records(self) -> list[dict[str, str]]:
        """Read the body rows as a list of dicts keyed by the header row.

        Row 1 is taken as the header (the exact inverse of building a table from
        `data=[{...}]` — see `insert_table`); each row below it becomes
        `{header: cell_text}`. A pure read — no `doc.edit()` needed.

        Edge cases mirror the write path: a **duplicate** header label collapses
        (the rightmost column wins), and a **blank** header cell yields an
        empty-string key — both the caller's responsibility.
        """
        headers = self._header_names()
        out: list[dict[str, str]] = []
        for r in range(2, self.row_count + 1):
            out.append({headers[c - 1]: self.cell(r, c).text for c in range(1, len(headers) + 1)})
        return out

    def add_row(self, values: list[Any] | None = None) -> None:
        """Append a row at the end of the table, optionally filling its cells.

        `values` are matched to columns left-to-right; extras past the column
        count are ignored, short lists leave trailing cells empty.
        """
        with _com.translate_com_errors():
            self._com.Rows.Add()
            if values:
                last = int(self._com.Rows.Count)
                cols = int(self._com.Columns.Count)
                for c, val in enumerate(values, start=1):
                    if c > cols:
                        break
                    self._com.Cell(last, c).Range.Text = str(val)

    def append_record(self, record: dict[str, Any]) -> None:
        """Append a row from a dict, mapping its keys to the header columns.

        Keys are matched against row 1's headers; a header with no matching key
        gets an empty cell and an extra key is ignored — the same lenient
        mapping `insert_table(data=[{...}])` uses. The new row inherits the
        table's existing formatting / banding (Word's `Rows.Add`). Wrap in
        `doc.edit(...)` for atomic undo.
        """
        headers = self._header_names()
        self.add_row([record.get(name, "") for name in headers])

    def update_row(self, key: Any, values: dict[str, Any], *, column: str | None = None) -> None:
        """Update the first row whose key-column cell equals `key`, by header name.

        The key column is the **first** column by default, or the header named
        by `column=`. Each item in `values` sets the cell under that header
        (`{header: new_text}`). First match wins when several rows share `key`.

        Validates against the header **before** mutating: an unknown `column`,
        or a `values` key that isn't a header, raises `OpError` (exit 1). If no
        row matches `key`, raises `AnchorNotFoundError` (exit 2). Wrap in
        `doc.edit(...)` for atomic undo.
        """
        headers = self._header_names()
        # Rightmost duplicate header wins, matching `records()`.
        col_of = {name: i + 1 for i, name in enumerate(headers)}
        if column is not None and column not in col_of:
            raise OpError(f"update_row: column {column!r} is not a header; have {headers}")
        unknown = [name for name in values if name not in col_of]
        if unknown:
            raise OpError(f"update_row: {unknown} not in the header row; have {headers}")
        key_col = col_of[column] if column is not None else 1
        target = str(key)
        for r in range(2, self.row_count + 1):
            if self.cell(r, key_col).text == target:
                for name, val in values.items():
                    self.cell(r, col_of[name]).set_text(str(val))
                return
        keyed = column if column is not None else (headers[0] if headers else "1")
        raise AnchorNotFoundError("table row", f"table:{self._index}:{keyed}={target!r}")

    def delete_row(self, index: int) -> None:
        """Delete the 1-based row `index`.

        Raises `AnchorNotFoundError` (kind `"table row"`) if out of range.
        """
        rows = self.row_count
        if not (1 <= index <= rows):
            raise AnchorNotFoundError("table row", f"table:{self._index}:row:{index}")
        with _com.translate_com_errors():
            self._com.Rows(index).Delete()

    def add_column(self, values: list[Any] | None = None) -> None:
        """Append a column at the right edge of the table, optionally filling it.

        The column mirror of `add_row`: `values` are matched to rows
        top-to-bottom; extras past the row count are ignored, a short list
        leaves trailing cells empty. (Word's `Columns.Add` tolerates a merged
        table, so this works where `delete_column` can't.) Wrap in
        `doc.edit(...)` for atomic undo.

        The new column lands at the right edge, so existing `table:N:R:C` ids are
        unchanged — but any cached `column_count` is now stale; re-read it.
        """
        with _com.translate_com_errors():
            self._com.Columns.Add()
            if values:
                last = int(self._com.Columns.Count)
                rows = int(self._com.Rows.Count)
                for r, val in enumerate(values, start=1):
                    if r > rows:
                        break
                    self._com.Cell(r, last).Range.Text = str(val)

    def delete_column(self, index: int) -> None:
        """Delete the 1-based column `index`.

        Raises `AnchorNotFoundError` (kind `"table column"`) if out of range.
        Word can't address an individual column on a table with merged /
        mixed-width cells ("mixed cell widths") — that `ComError` is re-raised as
        an `OpError` pointing at per-cell deletion via `table:N:R:C` (the same
        contract as a column-anchor style op). Wrap in `doc.edit(...)` for
        atomic undo.

        Every column to the right of `index` renumbers down by one, so any
        cached `table:N:R:C` ids past it are now stale — re-resolve through
        `doc.tables` before addressing another column.
        """
        cols = self.column_count
        if not (1 <= index <= cols):
            raise AnchorNotFoundError("table column", f"table:{self._index}:col:{index}")
        try:
            with _com.translate_com_errors():
                self._com.Columns(index).Delete()
        except ComError as e:
            raise OpError(
                f"cannot delete column {index} of table {self._index} ({e}); the "
                f"table has merged or mixed-width cells — delete its cells "
                f"individually via table:{self._index}:R:{index}"
            ) from e

    def set_heading_row(
        self, row: int = 1, *, heading: bool = True, allow_break: bool | None = None
    ) -> None:
        """Mark a 1-based row as a repeating table heading row.

        A heading row (`HeadingFormat`) repeats at the top of every page the
        table spans — set it on the header row of a multi-page table so the
        column labels carry over. `heading=False` clears the flag.

        `allow_break` controls `AllowBreakAcrossPages` (whether a row's content
        may split across a page boundary). It defaults to ``not heading`` — a
        repeating header shouldn't fracture — so the common
        `set_heading_row(1)` both repeats row 1 *and* keeps it intact; pass
        `allow_break` explicitly to override.

        Raises `AnchorNotFoundError` (kind `"table row"`) if `row` is out of
        range.
        """
        rows = self.row_count
        if not (1 <= row <= rows):
            raise AnchorNotFoundError("table row", f"table:{self._index}:row:{row}")
        keep_intact = (not heading) if allow_break is None else bool(allow_break)
        with _com.translate_com_errors():
            r = self._com.Rows(row)
            r.HeadingFormat = bool(heading)
            r.AllowBreakAcrossPages = keep_intact

    def autofit(self, mode: str = "content") -> None:
        """Resize the table's columns — fit to content, the window, or pin them.

        `mode` is one of:

        - ``"content"`` (default) — shrink/grow each column to fit its cells.
        - ``"window"`` — stretch the table to the page (container) width.
        - ``"fixed"`` — pin the current column widths so Word stops auto-sizing
          (sets `AllowAutoFit = False`).

        A clean way to tidy a table whose columns drifted after edits. An unknown
        `mode` raises `OpError`. Wrap in `doc.edit(...)` for atomic undo.
        """
        key = str(mode).lower()
        behavior = _AUTOFIT_MODES.get(key)
        if behavior is None:
            allowed = ", ".join(sorted(_AUTOFIT_MODES))
            raise OpError(f"autofit mode must be one of {allowed}; got {mode!r}")
        with _com.translate_com_errors():
            # "fixed" means "stop auto-sizing": AllowAutoFit off, then pin widths.
            # "content"/"window" need AllowAutoFit on for the behavior to take.
            self._com.AllowAutoFit = behavior != WdAutoFitBehavior.FIXED
            self._com.AutoFitBehavior(int(behavior))

    def set_style(self, name: str) -> None:
        """Restyle this **existing** table with a named table style.

        The post-creation counterpart of `insert_table(style=…)` — point a table
        at any built-in or custom table style (`"Grid Table 4 - Accent 1"`,
        `"Plain Table 3"`, …; discover them via `style list` filtered to
        `type=="table"`). Raises `StyleNotFoundError` (exit 2) if the style isn't
        defined in the document.

        **Direct cell shading is not preserved.** Applying a table style reapplies
        the style's conditional formatting (banding, header shading), which
        overwrites explicit per-cell `set_shading` colours (live-confirmed). So
        restyle **first**, then layer cell-level overrides on top — not the
        reverse. Wrap in `doc.edit(...)` for atomic undo.
        """
        style_obj = self._doc.styles[name]  # StyleNotFoundError (exit 2) if missing
        with _com.translate_com_errors():
            self._com.Style = style_obj.com

    def set_alignment(self, alignment: str) -> None:
        """Align the whole table across the page width — left, center, or right.

        `alignment` is ``"left"`` / ``"center"`` (``"centre"``) / ``"right"``,
        mapped onto `Table.Rows.Alignment`. This positions the *table* between the
        page margins (distinct from the text alignment *inside* cells, which is
        `format_paragraph`). Idempotent. Wrap in `doc.edit(...)`; bad input raises
        `OpError`.
        """
        value = _coerce_align(alignment, _ROW_ALIGN, "table alignment")
        with _com.translate_com_errors():
            self._com.Rows.Alignment = value

    def set_borders(
        self,
        *,
        sides: Any = "all",
        style: Any = "single",
        weight: Any = 0.5,
        color: Any = None,
    ) -> None:
        """Draw borders across the **whole table grid** in one call.

        The table-wide counterpart of the per-cell `set_borders` (a `Cell` is an
        `Anchor`). `sides` is ``"all"``/``"box"`` (the four outer edges — the
        default), a single outer edge (``"top"``/``"bottom"``/``"left"``/
        ``"right"``), the interior gridlines (``"horizontal"``/``"vertical"`` —
        the lines *between* cells), or a list (e.g. ``["box", "horizontal",
        "vertical"]`` to rule every line). `style` is a line style (``"single"``,
        ``"double"``, ``"dot"``, ``"dash"``, …, or ``"none"`` to clear). `weight`
        is the line width in points, snapped to Word's set (0.25/0.5/0.75/1/1.5/
        2.25/3). `color` is an optional name/hex/RGB. Idempotent. Bad input raises
        `OpError`. Wrap in `doc.edit(...)`.
        """
        try:
            with _com.translate_com_errors():
                apply_borders(
                    self._com.Borders, sides=sides, style=style, weight=weight, color=color
                )
        except (ValueError, TypeError) as e:
            raise OpError(str(e)) from e

    def set_banding(
        self,
        *,
        first_row: bool | None = None,
        last_row: bool | None = None,
        first_column: bool | None = None,
        last_column: bool | None = None,
        banded_rows: bool | None = None,
        banded_columns: bool | None = None,
    ) -> None:
        """Toggle the table-style options (Word's "Table Style Options" ribbon group).

        Each flag turns one conditional-formatting band of the **applied table
        style** on or off — `first_row` (a distinct header row), `last_row` (a
        total row), `first_column` / `last_column`, and `banded_rows` /
        `banded_columns` (the alternating stripes). All are tri-state: ``True`` /
        ``False`` set, ``None`` (the default) leaves that flag untouched.

        These only show once a **real table style** is applied — a styleless table
        or plain `"Table Grid"` ignores band conditions. Pair with `set_style`.
        Idempotent. Wrap in `doc.edit(...)` for atomic undo.
        """
        flags: dict[str, bool | None] = {
            "ApplyStyleHeadingRows": first_row,
            "ApplyStyleLastRow": last_row,
            "ApplyStyleFirstColumn": first_column,
            "ApplyStyleLastColumn": last_column,
            "ApplyStyleRowBands": banded_rows,
            "ApplyStyleColumnBands": banded_columns,
        }
        with _com.translate_com_errors():
            for prop, val in flags.items():
                if val is not None:
                    setattr(self._com, prop, bool(val))

    def delete(self) -> None:
        """Delete this entire table — the structural mirror of `add_row`.

        Removes the table and all its cells from the document. Afterwards this
        `Table` (and any `Cell` anchors derived from it) is stale; the indices
        of any tables that followed it shift down by one, so re-resolve through
        `doc.tables` before addressing another.
        """
        with _com.translate_com_errors():
            self._com.Delete()

    def __iter__(self) -> Iterator[Cell]:
        """Iterate cells row-major."""
        rows, cols = self.row_count, self.column_count
        for r in range(1, rows + 1):
            for c in range(1, cols + 1):
                yield Cell(self, r, c)

    def __repr__(self) -> str:
        return f"<Table {self._index} {self.title!r}>"


class TableCollection:
    """Indexable, iterable view over a document's tables.

    Index by 1-based position (`doc.tables[1]`) or by the table's `Title`
    (`doc.tables["Budget"]`). Positions match Word's own `Tables(n)` ordering —
    document order, top to bottom.
    """

    def __init__(self, doc: Document) -> None:
        self._doc = doc

    def __len__(self) -> int:
        with _com.translate_com_errors():
            return int(self._doc.com.Tables.Count)

    def __getitem__(self, key: int | str) -> Table:
        if isinstance(key, bool):
            # bool is an int subclass; reject before the int branch matches.
            raise TypeError(f"table key must be int or str, got {type(key).__name__}")
        if isinstance(key, int):
            n = len(self)
            if not (1 <= key <= n):
                raise AnchorNotFoundError("table", str(key))
            with _com.translate_com_errors():
                return Table(self._doc, self._doc.com.Tables(key), key)
        if isinstance(key, str):
            for table in self:
                if table.title == key:
                    return table
            raise AnchorNotFoundError("table", key)
        raise TypeError(f"table key must be int or str, got {type(key).__name__}")

    def __contains__(self, key: object) -> bool:
        if isinstance(key, bool):
            return False
        if isinstance(key, int):
            return 1 <= key <= len(self)
        if isinstance(key, str):
            return any(t.title == key for t in self)
        return False

    def __iter__(self) -> Iterator[Table]:
        with _com.translate_com_errors():
            count = int(self._doc.com.Tables.Count)
        for i in range(1, count + 1):
            with _com.translate_com_errors():
                com = self._doc.com.Tables(i)
            yield Table(self._doc, com, i)

    def list(self) -> list[dict[str, Any]]:
        """All tables as `{index, title, rows, columns}` dicts."""
        return [t.to_dict() for t in self]
