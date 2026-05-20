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

from typing import Any, Iterator, TYPE_CHECKING

from . import _com
from ._anchors import Anchor
from .exceptions import AnchorNotFoundError

if TYPE_CHECKING:
    from ._document import Document


def _strip_cell_text(raw: Any) -> str:
    """Cell text minus Word's trailing end-of-cell markers.

    A cell's `Range.Text` ends with CR + the cell mark (BEL, `\\x07`); a cell
    that contains multiple paragraphs repeats the pattern. `rstrip` of those
    code points gives the human-visible text, mirroring `paragraph_text`.
    """
    return str(raw or "").rstrip("\r\n\x07")


class Cell(Anchor):
    """A single table cell, addressed by 1-based (row, column).

    Subclasses `Anchor`, so it inherits `insert_before` / `insert_after` /
    `delete` / `apply_style` / `format_paragraph` unchanged. Only the bits that
    differ for cells — the COM range, text read/write, and the anchor id — are
    overridden here.
    """

    kind = "cell"

    def __init__(self, table: "Table", row: int, col: int) -> None:
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

    @property
    def text(self) -> str:
        with _com.translate_com_errors():
            return _strip_cell_text(self._cell().Range.Text)

    def set_text(self, text: str) -> None:
        with _com.translate_com_errors():
            self._cell().Range.Text = text


class Table:
    """Wraps a Word `Table` COM object, located by its 1-based document position.

    The index is stored at construction (the collection knows it without a COM
    round-trip), so `anchor_id` and cell ids never have to re-scan the document.
    """

    def __init__(self, doc: "Document", com: Any, index: int) -> None:
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
            raise AnchorNotFoundError(
                "table cell", f"table:{self._index}:{row}:{col}"
            )
        return Cell(self, row, col)

    def grid(self) -> list[list[str]]:
        """All cell text as a row-major `list[list[str]]`."""
        rows, cols = self.row_count, self.column_count
        return [
            [self.cell(r, c).text for c in range(1, cols + 1)]
            for r in range(1, rows + 1)
        ]

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

    def delete_row(self, index: int) -> None:
        """Delete the 1-based row `index`.

        Raises `AnchorNotFoundError` (kind `"table row"`) if out of range.
        """
        rows = self.row_count
        if not (1 <= index <= rows):
            raise AnchorNotFoundError("table row", f"table:{self._index}:row:{index}")
        with _com.translate_com_errors():
            self._com.Rows(index).Delete()

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

    def __init__(self, doc: "Document") -> None:
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
