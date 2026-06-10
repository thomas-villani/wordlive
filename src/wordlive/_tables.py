"""Tables — collection, table wrapper, and cell anchors.

Tables are document-scoped collections (`doc.tables`), so they live here rather
than in `_anchors.py`. A `Cell` *is* an `Anchor`, though: it targets the cell's
COM `Range`, so the inherited `apply_style` / `format_paragraph` / `set_text`
machinery works on cells with no special-casing, and `replace --anchor-id
table:N:R:C` resolves through `Document.anchor_by_id` like any other anchor.

The anchor-id scheme is `table:N:R:C` (1-based table index, row, column). The
bare `table:N` form is *not* an anchor — a whole table is a structural
collection, not a single range — so it's addressed via `doc.tables[N]` and the
`table` CLI group instead.

Limitation: cell addressing assumes a rectangular grid. Tables with merged or
split cells have a non-uniform COM cell model; `Table.cell(r, c)` follows Word's
own `Table.Cell(row, col)` indexing and may raise inside merged regions.
"""

from __future__ import annotations

from collections.abc import Iterator
from typing import TYPE_CHECKING, Any

from . import _com
from ._anchors import Anchor, range_text
from .exceptions import AnchorNotFoundError, OpError

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

    def __init__(self, table: Table, row: int, col: int) -> None:
        super().__init__(table._doc, name=f"table:{table.index}:{row}:{col}")
        self._table = table
        self._row = row
        self._col = col

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

    def cell(self, row: int, col: int) -> Cell:
        """Return the `Cell` at 1-based (row, col).

        Raises `AnchorNotFoundError` (kind `"table cell"`) if the coordinates
        fall outside the table's grid.
        """
        rows, cols = self.row_count, self.column_count
        if not (1 <= row <= rows and 1 <= col <= cols):
            raise AnchorNotFoundError("table cell", f"table:{self._index}:{row}:{col}")
        return Cell(self, row, col)

    def grid(self) -> list[list[str]]:
        """All cell text as a row-major `list[list[str]]`."""
        rows, cols = self.row_count, self.column_count
        return [[self.cell(r, c).text for c in range(1, cols + 1)] for r in range(1, rows + 1)]

    def read(self) -> dict[str, Any]:
        """Structured dump: metadata plus every cell with its addressable id.

        Each cell carries its `anchor_id` (`table:N:R:C`) so a caller can feed
        it straight back into `replace` / `style apply` / `format-paragraph`.
        """
        rows, cols = self.row_count, self.column_count
        cells = [
            [
                {
                    "row": r,
                    "col": c,
                    "text": self.cell(r, c).text,
                    "anchor_id": f"table:{self._index}:{r}:{c}",
                }
                for c in range(1, cols + 1)
            ]
            for r in range(1, rows + 1)
        ]
        return {
            "index": self._index,
            "title": self.title,
            "rows": rows,
            "columns": cols,
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
