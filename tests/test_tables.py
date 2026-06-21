"""Tables — collection, cell anchors, structural edits, anchor-id resolution."""

from __future__ import annotations

import pytest

import wordlive
from wordlive._format import to_bgr
from wordlive._ops import run_batch
from wordlive.constants import (
    WdBorderType,
    WdCellVerticalAlignment,
    WdLineStyle,
    WdParagraphAlignment,
    WdRowAlignment,
)
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
# Heading-row repeat — Rows(n).HeadingFormat / AllowBreakAcrossPages
# ---------------------------------------------------------------------------


def test_set_heading_row_marks_repeating(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        t = doc.tables[1]
        with doc.edit("heading row"):
            t.set_heading_row(1)
        row = fake_word.ActiveDocument.Tables(1).Rows(1)
        assert row.HeadingFormat is True
        # A repeating header shouldn't fracture across a page by default.
        assert row.AllowBreakAcrossPages is False


def test_set_heading_row_clear(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        t = doc.tables[1]
        with doc.edit("clear heading"):
            t.set_heading_row(1, heading=False)
        row = fake_word.ActiveDocument.Tables(1).Rows(1)
        assert row.HeadingFormat is False
        assert row.AllowBreakAcrossPages is True


def test_set_heading_row_allow_break_override(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        t = doc.tables[1]
        with doc.edit("heading allow break"):
            t.set_heading_row(1, heading=True, allow_break=True)
        row = fake_word.ActiveDocument.Tables(1).Rows(1)
        assert row.HeadingFormat is True
        assert row.AllowBreakAcrossPages is True


def test_set_heading_row_out_of_range_raises(fake_word):
    with wordlive.attach() as word:
        t = word.documents.active.tables[1]
        with pytest.raises(AnchorNotFoundError) as exc_info:
            t.set_heading_row(9)
    assert exc_info.value.kind == "table row"


def test_exec_set_heading_row(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        result, exc = run_batch(
            doc,
            [{"op": "set_heading_row", "table": 1, "row": 1}],
            label="t",
        )
    assert exc is None and result["ok"] is True
    row = fake_word.ActiveDocument.Tables(1).Rows(1)
    assert row.HeadingFormat is True


# ---------------------------------------------------------------------------
# Table-as-records — read (records) / write (append_record / update_row)
# ---------------------------------------------------------------------------
# The seeded "Grid" table is [["A1","B1"],["A2","B2"]] — row 1 ("A1"/"B1") is the
# header, so the lone body row reads back keyed by it.


def test_records_keys_body_rows_by_header(fake_word):
    with wordlive.attach() as word:
        recs = word.documents.active.tables[1].records()
    assert recs == [{"A1": "A2", "B1": "B2"}]


def test_append_record_maps_keys_to_header_columns(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        t = doc.tables[1]
        with doc.edit("append record"):
            t.append_record({"B1": "Y", "A1": "X"})  # order-independent
        assert t.row_count == 3
        assert t.cell(3, 1).text == "X" and t.cell(3, 2).text == "Y"


def test_append_record_missing_key_empty_extra_ignored(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        t = doc.tables[1]
        with doc.edit("append partial record"):
            t.append_record({"A1": "only", "Zzz": "ignored"})
        assert t.cell(3, 1).text == "only" and t.cell(3, 2).text == ""


def test_update_row_matches_first_column_sets_by_header(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        t = doc.tables[1]
        with doc.edit("update row"):
            t.update_row("A2", {"B1": "Z"})
        assert t.cell(2, 2).text == "Z"


def test_update_row_matches_named_column(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        t = doc.tables[1]
        with doc.edit("update by column"):
            t.update_row("B2", {"A1": "Q"}, column="B1")
        assert t.cell(2, 1).text == "Q"


def test_update_row_no_match_raises_anchor_not_found(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        t = doc.tables[1]
        with pytest.raises(AnchorNotFoundError) as exc_info:
            with doc.edit("update miss"):
                t.update_row("nope", {"B1": "Z"})
    assert exc_info.value.kind == "table row"


def test_update_row_unknown_header_raises_op_error(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        t = doc.tables[1]
        with pytest.raises(OpError):
            with doc.edit("update bad header"):
                t.update_row("A2", {"Nonexistent": "Z"})


def test_update_row_unknown_column_raises_op_error(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        t = doc.tables[1]
        with pytest.raises(OpError):
            with doc.edit("update bad column"):
                t.update_row("A2", {"B1": "Z"}, column="Nope")


def test_exec_append_record_and_update_row(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        result, exc = run_batch(
            doc,
            [
                {"op": "append_record", "table": 1, "record": {"A1": "X", "B1": "Y"}},
                {"op": "update_row", "table": 1, "key": "A2", "values": {"B1": "Z"}},
            ],
            label="t",
        )
        assert exc is None and result["ok"] is True
        t = doc.tables[1]
        assert t.cell(3, 1).text == "X" and t.cell(3, 2).text == "Y"
        assert t.cell(2, 2).text == "Z"


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


def test_add_table_at_document_end_opens_trailing_paragraph(fake_word):
    # Tables.Add at the document's final (undeletable) paragraph mark crashes
    # Word; insert_table must open a trailing paragraph first. Disable the
    # abutting-table separators so only the terminal-boundary guard can fire.
    fake_word.ActiveDocument.Range(33, 34).Information.return_value = 0
    fake_word.ActiveDocument.Range(34, 35).Information.return_value = 0
    with wordlive.attach() as word:
        doc = word.documents.active
        with doc.edit("add table at end"):
            doc.add_table(2, 2)
        assert len(doc.tables) == 2
    # Content.End is 35 in the fixture; the final mark sits at 34. The guard
    # opened a separator paragraph there so the table lands before it.
    assert fake_word.ActiveDocument.Range(34, 34).Text == "\r"


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


def test_insert_table_resets_cells_to_normal(fake_word):
    # New cells must default to the body style regardless of the anchor's
    # paragraph style — no inheriting Heading 2 from the heading above.
    with wordlive.attach() as word:
        doc = word.documents.active
        with doc.edit("table under heading"):
            t = doc.headings["Risks"].insert_table(2, 2)
        tbl = fake_word.ActiveDocument.Tables(t.index)
        for r in (1, 2):
            for c in (1, 2):
                style = tbl.Cell(r, c).Range.Style
                assert style is not None and style.NameLocal == "Normal"


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
# table from tabular data — records + dimension inference
# ---------------------------------------------------------------------------


def test_normalize_table_data_records_to_grid():
    from wordlive._anchors import _normalize_table_data

    grid, header = _normalize_table_data(
        [{"Item": "Travel", "Cost": "$400"}, {"Item": "Lodging", "Cost": "$600"}]
    )
    assert header is True
    assert grid == [["Item", "Cost"], ["Travel", "$400"], ["Lodging", "$600"]]


def test_normalize_table_data_array_passes_through():
    from wordlive._anchors import _normalize_table_data

    grid, header = _normalize_table_data([["a", "b"], ["c", "d"]])
    assert header is False
    assert grid == [["a", "b"], ["c", "d"]]


def test_normalize_table_data_records_missing_key_is_empty_cell():
    from wordlive._anchors import _normalize_table_data

    # First record fixes the columns; a later record missing a key → empty cell.
    grid, _ = _normalize_table_data([{"A": "1", "B": "2"}, {"A": "3"}])
    assert grid == [["A", "B"], ["1", "2"], ["3", ""]]


def test_normalize_table_data_mixed_shape_raises():
    from wordlive._anchors import _normalize_table_data

    with pytest.raises(OpError, match="mixes"):
        _normalize_table_data([{"A": "1"}, ["B", "2"]])


def test_add_table_from_records_infers_dims_and_header(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        with doc.edit("records table"):
            t = doc.end.insert_table(
                data=[{"Item": "Travel", "Cost": "$400"}, {"Item": "Lodging", "Cost": "$600"}]
            )
        assert t.row_count == 3 and t.column_count == 2
        assert t.grid() == [["Item", "Cost"], ["Travel", "$400"], ["Lodging", "$600"]]


def test_add_table_from_array_infers_dims(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        with doc.edit("array table"):
            t = doc.end.insert_table(data=[["Name", "Qty"], ["Widget", "3"]])
        assert t.row_count == 2 and t.column_count == 2
        assert t.grid() == [["Name", "Qty"], ["Widget", "3"]]


def test_insert_table_without_dims_or_data_raises(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        with pytest.raises(OpError, match="rows and cols"):
            doc.end.insert_table()


def test_insert_table_explicit_dims_pad_beyond_data(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        with doc.edit("padded"):
            t = doc.end.insert_table(4, 2, data=[["a", "b"]])
        assert t.row_count == 4 and t.column_count == 2


def test_exec_create_table_from_records(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        result, exc = run_batch(
            doc,
            [{"op": "create_table", "anchor_id": "end", "data": [{"X": "1", "Y": "2"}]}],
            label="rec",
        )
        assert exc is None
        out = result["outputs"][0]
        assert (out["rows"], out["columns"]) == (2, 2)
        assert doc.tables[out["table"]].grid() == [["X", "Y"], ["1", "2"]]


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


# ---------------------------------------------------------------------------
# Table styling — set_style / set_alignment / set_borders / set_banding
# ---------------------------------------------------------------------------


def test_set_style_restyles_existing_table(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        t = doc.tables[1]
        with doc.edit("set style"):
            t.set_style("Heading 1")
        assert fake_word.ActiveDocument.Tables(1).Style is doc.styles["Heading 1"].com


def test_set_style_unknown_raises(fake_word):
    with wordlive.attach() as word:
        t = word.documents.active.tables[1]
        with pytest.raises(StyleNotFoundError):
            t.set_style("No Such Table Style")


def test_exec_set_table_style(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        result, exc = run_batch(
            doc, [{"op": "set_table_style", "table": 1, "style": "Heading 1"}], label="t"
        )
    assert exc is None and result["ok"] is True
    assert fake_word.ActiveDocument.Tables(1).Style is doc.styles["Heading 1"].com


def test_set_alignment_centers_table(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        with doc.edit("align"):
            doc.tables[1].set_alignment("center")
        assert fake_word.ActiveDocument.Tables(1).Rows.Alignment == int(WdRowAlignment.CENTER)


def test_set_alignment_bad_raises(fake_word):
    with wordlive.attach() as word:
        with pytest.raises(OpError):
            word.documents.active.tables[1].set_alignment("middle")


def test_exec_set_table_alignment(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        result, exc = run_batch(
            doc, [{"op": "set_table_alignment", "table": 1, "alignment": "right"}], label="t"
        )
    assert exc is None and result["ok"] is True
    assert fake_word.ActiveDocument.Tables(1).Rows.Alignment == int(WdRowAlignment.RIGHT)


def test_set_table_borders_whole_grid(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        with doc.edit("table borders"):
            doc.tables[1].set_borders(sides="box", style="double", weight=1.5)
    borders = fake_word.ActiveDocument.Tables(1).Borders
    top = borders(int(WdBorderType.TOP))
    assert top.LineStyle == int(WdLineStyle.DOUBLE)
    # 1.5pt snaps to WdLineWidth 12 (points x 8).
    assert top.LineWidth == 12


def test_set_table_borders_bad_side_raises(fake_word):
    with wordlive.attach() as word:
        with pytest.raises(OpError):
            word.documents.active.tables[1].set_borders(sides="diagonal")


def test_exec_set_table_borders_line_style_alias(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        result, exc = run_batch(
            doc,
            [{"op": "set_table_borders", "table": 1, "sides": "top", "line_style": "dot"}],
            label="t",
        )
    assert exc is None and result["ok"] is True
    borders = fake_word.ActiveDocument.Tables(1).Borders
    assert borders(int(WdBorderType.TOP)).LineStyle == int(WdLineStyle.DOT)


def test_set_banding_flips_only_passed_flags(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        com = fake_word.ActiveDocument.Tables(1)
        com.ApplyStyleLastRow = True  # pre-existing flag must survive
        with doc.edit("banding"):
            doc.tables[1].set_banding(first_row=True, banded_rows=False)
    assert com.ApplyStyleHeadingRows is True
    assert com.ApplyStyleRowBands is False
    assert com.ApplyStyleLastRow is True  # untouched (None) flags unchanged
    assert com.ApplyStyleFirstColumn is False


def test_exec_set_table_banding(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        result, exc = run_batch(
            doc,
            [{"op": "set_table_banding", "table": 1, "banded_columns": True}],
            label="t",
        )
    assert exc is None and result["ok"] is True
    assert fake_word.ActiveDocument.Tables(1).ApplyStyleColumnBands is True


# ---------------------------------------------------------------------------
# Cell vertical alignment
# ---------------------------------------------------------------------------


def test_cell_set_vertical_alignment_bottom(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        with doc.edit("valign"):
            doc.anchor_by_id("table:1:1:1").set_vertical_alignment("bottom")
    # BOTTOM is 3 (the gap at 2 — wdAlignVerticalJustify — a cell rejects).
    assert fake_word.ActiveDocument.Tables(1).Cell(1, 1).VerticalAlignment == int(
        WdCellVerticalAlignment.BOTTOM
    )


def test_cell_vertical_alignment_bad_raises(fake_word):
    with wordlive.attach() as word:
        cell = word.documents.active.tables[1].cell(1, 1)
        with pytest.raises(OpError):
            cell.set_vertical_alignment("justify")


def test_exec_set_cell_vertical_alignment(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        result, exc = run_batch(
            doc,
            [{"op": "set_cell_vertical_alignment", "anchor_id": "table:1:2:1", "align": "center"}],
            label="t",
        )
    assert exc is None and result["ok"] is True
    assert fake_word.ActiveDocument.Tables(1).Cell(2, 1).VerticalAlignment == int(
        WdCellVerticalAlignment.CENTER
    )


def test_exec_set_cell_vertical_alignment_non_cell_raises(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        result, exc = run_batch(
            doc,
            [{"op": "set_cell_vertical_alignment", "anchor_id": "start", "align": "top"}],
            label="t",
        )
    assert exc is not None and isinstance(exc, OpError)


# ---------------------------------------------------------------------------
# Row / column anchors — table:N:row:R  /  table:N:col:C
# ---------------------------------------------------------------------------


def test_anchor_by_id_resolves_row(fake_word):
    with wordlive.attach() as word:
        anchor = word.documents.active.anchor_by_id("table:1:row:1")
    assert isinstance(anchor, wordlive.RowAnchor)
    assert anchor.anchor_id == "table:1:row:1"
    assert anchor.row == 1


def test_anchor_by_id_resolves_column(fake_word):
    with wordlive.attach() as word:
        anchor = word.documents.active.anchor_by_id("table:1:col:2")
    assert isinstance(anchor, wordlive.ColumnAnchor)
    assert anchor.anchor_id == "table:1:col:2"
    assert anchor.column == 2


def test_table_row_and_column_accessors(fake_word):
    with wordlive.attach() as word:
        t = word.documents.active.tables[1]
        assert isinstance(t.row(2), wordlive.RowAnchor)
        assert isinstance(t.column(1), wordlive.ColumnAnchor)


def test_row_accessor_out_of_range_raises(fake_word):
    with wordlive.attach() as word:
        with pytest.raises(AnchorNotFoundError) as exc_info:
            word.documents.active.tables[1].row(9)
    assert exc_info.value.kind == "table row"


def test_column_accessor_out_of_range_raises(fake_word):
    with wordlive.attach() as word:
        with pytest.raises(AnchorNotFoundError) as exc_info:
            word.documents.active.tables[1].column(9)
    assert exc_info.value.kind == "table column"


def test_row_anchor_shades_whole_row(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        with doc.edit("row shade"):
            doc.anchor_by_id("table:1:row:1").set_shading(fill="yellow")
    row_range = fake_word.ActiveDocument.Tables(1).Rows(1).Range
    assert row_range.Shading.BackgroundPatternColor == to_bgr("yellow")


def test_row_anchor_set_text_refused(fake_word):
    with wordlive.attach() as word:
        with pytest.raises(OpError):
            word.documents.active.anchor_by_id("table:1:row:1").set_text("nope")


def test_column_anchor_shades_each_cell(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        with doc.edit("col shade"):
            doc.anchor_by_id("table:1:col:1").set_shading(fill="red")
    com = fake_word.ActiveDocument.Tables(1)
    for r in (1, 2):  # the 2x2 "Grid" table's column 1
        assert com.Cell(r, 1).Range.Shading.BackgroundPatternColor == to_bgr("red")


def test_column_anchor_right_aligns_each_cell(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        with doc.edit("col align"):
            doc.anchor_by_id("table:1:col:2").format_paragraph(alignment="right")
    com = fake_word.ActiveDocument.Tables(1)
    for r in (1, 2):
        assert com.Cell(r, 2).Range.ParagraphFormat.Alignment == int(WdParagraphAlignment.RIGHT)


def test_column_anchor_set_text_refused(fake_word):
    with wordlive.attach() as word:
        with pytest.raises(OpError):
            word.documents.active.anchor_by_id("table:1:col:1").set_text("nope")
