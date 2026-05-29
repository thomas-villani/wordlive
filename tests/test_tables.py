"""Tables — collection, cell anchors, structural edits, anchor-id resolution."""

from __future__ import annotations

import pytest

import wordlive
from wordlive._ops import run_batch
from wordlive.constants import WdParagraphAlignment
from wordlive.exceptions import AnchorNotFoundError, OpError, StyleNotFoundError

# ---------------------------------------------------------------------------
# TableCollection
# ---------------------------------------------------------------------------


def test_tables_list_returns_metadata(fake_word):
    with wordlive.attach() as word:
        rows = word.documents.active.tables.list()
    assert len(rows) == 1
    assert rows[0] == {"index": 1, "title": "Grid", "rows": 2, "columns": 2}


def test_tables_len(fake_word):
    with wordlive.attach() as word:
        assert len(word.documents.active.tables) == 1


def test_tables_contains(fake_word):
    with wordlive.attach() as word:
        tables = word.documents.active.tables
        assert 1 in tables
        assert 2 not in tables
        assert "Grid" in tables
        assert "Nope" not in tables
        assert True not in tables  # bool rejected before the int branch


def test_tables_getitem_by_index(fake_word):
    with wordlive.attach() as word:
        t = word.documents.active.tables[1]
    assert t.index == 1
    assert t.row_count == 2
    assert t.column_count == 2
    assert t.title == "Grid"


def test_tables_getitem_by_title(fake_word):
    with wordlive.attach() as word:
        t = word.documents.active.tables["Grid"]
    assert t.index == 1


def test_tables_getitem_out_of_range_raises(fake_word):
    with wordlive.attach() as word:
        with pytest.raises(AnchorNotFoundError) as exc_info:
            _ = word.documents.active.tables[5]
    assert exc_info.value.kind == "table"


def test_tables_getitem_bad_title_raises(fake_word):
    with wordlive.attach() as word:
        with pytest.raises(AnchorNotFoundError):
            _ = word.documents.active.tables["Missing"]


def test_tables_getitem_bool_rejected(fake_word):
    with wordlive.attach() as word:
        with pytest.raises(TypeError):
            _ = word.documents.active.tables[True]


def test_tables_iter_yields_tables(fake_word):
    with wordlive.attach() as word:
        tables = list(word.documents.active.tables)
    assert [t.index for t in tables] == [1]


# ---------------------------------------------------------------------------
# Table / Cell reads
# ---------------------------------------------------------------------------


def test_cell_text_strips_markers(fake_word):
    with wordlive.attach() as word:
        cell = word.documents.active.tables[1].cell(1, 1)
    # Fake seeds "A1\r\x07"; the trailing cell markers must be stripped.
    assert cell.text == "A1"


def test_cell_anchor_id(fake_word):
    with wordlive.attach() as word:
        cell = word.documents.active.tables[1].cell(2, 2)
    assert cell.anchor_id == "table:1:2:2"
    assert cell.kind == "cell"
    assert cell.row == 2 and cell.column == 2


def test_cell_out_of_range_raises(fake_word):
    with wordlive.attach() as word:
        t = word.documents.active.tables[1]
        with pytest.raises(AnchorNotFoundError) as exc_info:
            t.cell(3, 1)
    assert exc_info.value.kind == "table cell"


def test_table_grid(fake_word):
    with wordlive.attach() as word:
        grid = word.documents.active.tables[1].grid()
    assert grid == [["A1", "B1"], ["A2", "B2"]]


def test_table_read_shape(fake_word):
    with wordlive.attach() as word:
        data = word.documents.active.tables[1].read()
    assert data["index"] == 1
    assert data["rows"] == 2 and data["columns"] == 2
    assert data["cells"][0][0] == {
        "row": 1,
        "col": 1,
        "text": "A1",
        "anchor_id": "table:1:1:1",
    }
    assert data["cells"][1][1]["anchor_id"] == "table:1:2:2"


def test_table_iter_yields_cells_row_major(fake_word):
    with wordlive.attach() as word:
        cells = list(word.documents.active.tables[1])
    assert [c.anchor_id for c in cells] == [
        "table:1:1:1",
        "table:1:1:2",
        "table:1:2:1",
        "table:1:2:2",
    ]


# ---------------------------------------------------------------------------
# Cell writes
# ---------------------------------------------------------------------------


def test_cell_set_text_round_trip(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        with doc.edit("set cell"):
            doc.tables[1].cell(1, 2).set_text("changed")
        assert doc.tables[1].cell(1, 2).text == "changed"


def test_cell_apply_style_writes_through(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        with doc.edit("style cell"):
            doc.tables[1].cell(1, 1).apply_style("Heading 2")
    cell_range = fake_word.ActiveDocument.Tables(1).Cell(1, 1).Range
    assert cell_range.Style.NameLocal == "Heading 2"


def test_cell_format_paragraph(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        doc.tables[1].cell(2, 1).format_paragraph(alignment="center", left_indent=12.0)
    cell_range = fake_word.ActiveDocument.Tables(1).Cell(2, 1).Range
    pf = cell_range.ParagraphFormat
    assert pf.Alignment == int(WdParagraphAlignment.CENTER)
    assert pf.LeftIndent == 12.0


# ---------------------------------------------------------------------------
# Structural edits
# ---------------------------------------------------------------------------


def test_add_row_increases_count(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        t = doc.tables[1]
        with doc.edit("add row"):
            t.add_row()
        assert t.row_count == 3


def test_add_row_with_values_fills_cells(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        t = doc.tables[1]
        with doc.edit("add row with values"):
            t.add_row(["X", "Y"])
        assert t.cell(3, 1).text == "X"
        assert t.cell(3, 2).text == "Y"


def test_delete_row_decreases_count(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        t = doc.tables[1]
        with doc.edit("delete row"):
            t.delete_row(1)
        assert t.row_count == 1
        # Row 2 ("A2"/"B2") is now the only row.
        assert t.cell(1, 1).text == "A2"


def test_delete_row_out_of_range_raises(fake_word):
    with wordlive.attach() as word:
        t = word.documents.active.tables[1]
        with pytest.raises(AnchorNotFoundError) as exc_info:
            t.delete_row(9)
    assert exc_info.value.kind == "table row"


# ---------------------------------------------------------------------------
# Table creation / deletion
# ---------------------------------------------------------------------------


def test_add_table_appends_and_returns(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        with doc.edit("add table"):
            t = doc.add_table(2, 3)
        assert t.index == 2  # appended after the seeded "Grid" table
        assert t.row_count == 2
        assert t.column_count == 3
        assert len(doc.tables) == 2


def test_add_table_populates_data(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        with doc.edit("add table with data"):
            t = doc.add_table(2, 2, data=[["Name", "Qty"], ["Widget", "3"]])
        assert t.grid() == [["Name", "Qty"], ["Widget", "3"]]


def test_add_table_partial_data_leaves_cells_empty(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        with doc.edit("add table partial"):
            t = doc.add_table(2, 2, data=[["only-one"]])
        assert t.grid() == [["only-one", ""], ["", ""]]


def test_add_table_data_overflow_raises(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        with pytest.raises(OpError):
            doc.add_table(1, 2, data=[["a", "b", "c"]])  # 3 cells > 2 columns
        with pytest.raises(OpError):
            doc.add_table(1, 2, data=[["a"], ["b"]])  # 2 rows > 1 row


def test_add_table_bad_dimensions_raise(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        with pytest.raises(OpError):
            doc.add_table(0, 2)
        with pytest.raises(OpError):
            doc.add_table(2, -1)


def test_insert_table_at_heading(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        with doc.edit("table under heading"):
            t = doc.headings["Introduction"].insert_table(2, 2)
        assert t.row_count == 2 and t.column_count == 2
        assert len(doc.tables) == 2


def test_insert_table_unknown_style_raises(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        with pytest.raises(StyleNotFoundError):
            doc.add_table(2, 2, style="No Such Table Style")
        # The bad style is rejected before any table is created.
        assert len(doc.tables) == 1


def test_insert_table_applies_explicit_style(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        with doc.edit("styled table"):
            t = doc.add_table(2, 2, style="Heading 1")  # any style defined in the doc
        applied = fake_word.ActiveDocument.Tables(t.index).Style
        assert applied is not None
        assert applied.NameLocal == "Heading 1"


def test_table_delete_removes_table(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        with doc.edit("add then delete"):
            t = doc.add_table(2, 2)
        assert len(doc.tables) == 2
        with doc.edit("delete table"):
            doc.tables[t.index].delete()
        assert len(doc.tables) == 1


# ---------------------------------------------------------------------------
# exec ops: create_table / delete_table
# ---------------------------------------------------------------------------


def test_exec_create_table_reports_index(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        result, exc = run_batch(
            doc,
            [
                {
                    "op": "create_table",
                    "anchor_id": "end",
                    "rows": 2,
                    "cols": 2,
                    "data": [["A", "B"], ["C", "D"]],
                }
            ],
            label="test",
        )
    assert exc is None
    assert result["ok"] is True
    assert result["outputs"] == [
        {"index": 0, "op": "create_table", "table": 2, "rows": 2, "columns": 2}
    ]
    assert doc.tables[2].grid() == [["A", "B"], ["C", "D"]]


def test_exec_delete_table(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        run_batch(
            doc, [{"op": "create_table", "anchor_id": "end", "rows": 1, "cols": 1}], label="c"
        )
        assert len(doc.tables) == 2
        result, exc = run_batch(doc, [{"op": "delete_table", "table": 2}], label="d")
    assert exc is None
    assert result["ok"] is True
    assert len(doc.tables) == 1


def test_exec_create_table_missing_field_fails_cleanly(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        result, exc = run_batch(
            doc, [{"op": "create_table", "anchor_id": "end", "rows": 2}], label="bad"
        )
    assert exc is not None
    assert result["ok"] is False
    assert "cols" in result["failure"]["error"]


# ---------------------------------------------------------------------------
# anchor_by_id resolution
# ---------------------------------------------------------------------------


def test_anchor_by_id_resolves_cell(fake_word):
    with wordlive.attach() as word:
        anchor = word.documents.active.anchor_by_id("table:1:2:2")
    assert isinstance(anchor, wordlive.Cell)
    assert anchor.text == "B2"
    assert anchor.anchor_id == "table:1:2:2"


def test_anchor_by_id_cell_is_writable(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        with doc.edit("replace cell via id"):
            doc.anchor_by_id("table:1:1:1").set_text("Z")
        assert doc.tables[1].cell(1, 1).text == "Z"


def test_anchor_by_id_bare_table_raises(fake_word):
    with wordlive.attach() as word:
        with pytest.raises(AnchorNotFoundError) as exc_info:
            word.documents.active.anchor_by_id("table:1")
    assert exc_info.value.kind == "table cell"


def test_anchor_by_id_partial_table_id_raises(fake_word):
    with wordlive.attach() as word:
        with pytest.raises(AnchorNotFoundError):
            word.documents.active.anchor_by_id("table:1:2")


def test_anchor_by_id_non_numeric_table_id_raises(fake_word):
    with wordlive.attach() as word:
        with pytest.raises(AnchorNotFoundError):
            word.documents.active.anchor_by_id("table:a:b:c")
