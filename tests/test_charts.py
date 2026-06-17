"""Unit tests for the Excel-backed chart surface (off-Windows, fake COM).

The *live* AddChart2 path (selection-based insert, embedded-Excel data grid,
BreakLink-to-static) needs real Word + Excel and lives in the smoke suite. Here
we cover everything that doesn't: data validation, the Excel-availability gate,
insertion against the fake (cells written, SERIES formula, BreakLink, workbook
closed), the read collection (`doc.charts`), anchor resolution (`chart:N`), and
the ops / CLI / MCP request surfaces.
"""

from __future__ import annotations

import json

import pytest
from click.testing import CliRunner

import wordlive
from wordlive import _charts
from wordlive._ops import OP_OPTIONAL_FIELDS, OP_REQUIRED_FIELDS, run_batch
from wordlive.cli.main import main
from wordlive.exceptions import AnchorNotFoundError, ExcelNotAvailableError, OpError


@pytest.fixture
def excel_available(monkeypatch: pytest.MonkeyPatch) -> None:
    """Force the Excel probe True so insertion tests run regardless of platform."""
    monkeypatch.setattr(_charts, "probe_excel_available", lambda: True)


# ---------------------------------------------------------------------------
# normalize_chart_data — pure validation (no COM)
# ---------------------------------------------------------------------------


def test_normalize_mapping_categorical():
    xs, ys = _charts.normalize_chart_data("bar", {"Q1": 10, "Q2": 25, "Q3": 18})
    assert xs == ["Q1", "Q2", "Q3"]
    assert ys == [10.0, 25.0, 18.0]


def test_normalize_pairs_scatter_numeric_x():
    xs, ys = _charts.normalize_chart_data("scatter", [[1.2, 3.4], [1.2, 3.9], [2.5, 6.1]])
    # scatter keeps duplicate x and coerces both axes to float
    assert xs == [1.2, 1.2, 2.5]
    assert ys == [3.4, 3.9, 6.1]
    assert all(isinstance(x, float) for x in xs)


def test_normalize_line_accepts_pairs():
    xs, ys = _charts.normalize_chart_data("line", [[0, 1.0], [1, 2.3]])
    assert xs == [0, 1]  # categorical kind keeps x as given
    assert ys == [1.0, 2.3]


def test_normalize_scatter_rejects_non_numeric_x():
    with pytest.raises(OpError, match="scatter x"):
        _charts.normalize_chart_data("scatter", [["a", 1.0]])


def test_normalize_rejects_non_numeric_value():
    with pytest.raises(OpError, match="value"):
        _charts.normalize_chart_data("bar", {"Q1": "lots"})


def test_normalize_rejects_bool_value():
    with pytest.raises(OpError, match="value"):
        _charts.normalize_chart_data("bar", {"Q1": True})


def test_normalize_rejects_empty():
    with pytest.raises(OpError, match="empty"):
        _charts.normalize_chart_data("bar", {})
    with pytest.raises(OpError, match="empty"):
        _charts.normalize_chart_data("bar", [])


def test_normalize_rejects_bad_pair_shape():
    with pytest.raises(OpError, match=r"\[x, y\] pairs"):
        _charts.normalize_chart_data("scatter", [[1, 2, 3]])
    with pytest.raises(OpError, match=r"\[x, y\] pairs"):
        _charts.normalize_chart_data("scatter", ["nope"])


def test_normalize_rejects_unknown_shape():
    with pytest.raises(OpError, match="object .* or an array"):
        _charts.normalize_chart_data("bar", 42)


def test_kind_maps_are_consistent():
    assert _charts.CHART_KINDS == ("bar", "pie", "line", "scatter")
    # every kind round-trips kind -> XlChartType int -> kind
    for kind, xl in _charts.KIND_TO_XL.items():
        assert _charts.XL_TO_KIND[int(xl)] == kind


# ---------------------------------------------------------------------------
# probe_excel_available — real probe + the gate it drives
# ---------------------------------------------------------------------------


def test_probe_returns_bool():
    assert isinstance(_charts.probe_excel_available(), bool)


def test_insert_chart_raises_when_excel_missing(fake_word, monkeypatch):
    monkeypatch.setattr(_charts, "probe_excel_available", lambda: False)
    with wordlive.attach() as word:
        doc = word.documents.active
        with pytest.raises(ExcelNotAvailableError):
            doc.start.insert_chart("bar", {"a": 1})
        # the document is untouched — no chart was inserted
        assert len(doc.charts) == 0


def test_insert_chart_validation_before_excel_probe(fake_word, monkeypatch):
    # bad kind / where / data must raise without even probing Excel
    def boom() -> bool:  # pragma: no cover - must never be called
        raise AssertionError("probe should not run on a validation failure")

    monkeypatch.setattr(_charts, "probe_excel_available", boom)
    with wordlive.attach() as word:
        doc = word.documents.active
        with pytest.raises(ValueError, match="unknown chart kind"):
            doc.start.insert_chart("donut", {"a": 1})
        with pytest.raises(ValueError, match="where must be"):
            doc.start.insert_chart("bar", {"a": 1}, where="sideways")
        with pytest.raises(OpError):
            doc.start.insert_chart("bar", {})


# ---------------------------------------------------------------------------
# insert_chart — happy path against the fake (cells, SERIES formula, BreakLink)
# ---------------------------------------------------------------------------


def test_insert_chart_returns_anchor_and_populates(fake_word, excel_available):
    with wordlive.attach() as word:
        doc = word.documents.active
        chart = doc.start.insert_chart("bar", {"Q1": 10, "Q2": 25, "Q3": 18}, title="Quarterly")
        assert isinstance(chart, wordlive.ChartAnchor)
        assert chart.anchor_id == "chart:1"

        shape = doc.charts[1]._shape()
        fake_chart = shape.Chart
        # placeholder series collapsed to one, pointed at the cell ranges
        assert fake_chart.SeriesCollection().Count == 1
        formula = fake_chart.SeriesCollection()(1).Formula
        assert formula == "=SERIES(Sheet1!$B$1,Sheet1!$A$2:$A$4,Sheet1!$B$2:$B$4,1)"
        # data written into the workbook cells (header row + 3 points)
        cells = fake_chart.ChartData.Workbook._ws.Cells.values
        assert cells[(1, 2)] == "Quarterly"  # series name / header
        assert [cells[(r, 1)] for r in (2, 3, 4)] == ["Q1", "Q2", "Q3"]
        assert [cells[(r, 2)] for r in (2, 3, 4)] == [10.0, 25.0, 18.0]
        # finalised: title set, link broken, workbook closed (no orphan Excel)
        assert fake_chart.HasTitle is True
        assert fake_chart.ChartTitle.Text == "Quarterly"
        assert fake_chart.ChartData.linked is False
        assert fake_chart.ChartData.Workbook.closed is True


def test_insert_chart_scatter_writes_numeric_x(fake_word, excel_available):
    with wordlive.attach() as word:
        doc = word.documents.active
        doc.start.insert_chart("scatter", [[1.2, 3.4], [1.2, 3.9], [2.5, 6.1]])
        cells = doc.charts[1]._shape().Chart.ChartData.Workbook._ws.Cells.values
        assert [cells[(r, 1)] for r in (2, 3, 4)] == [1.2, 1.2, 2.5]  # numeric x, dup kept
        assert [cells[(r, 2)] for r in (2, 3, 4)] == [3.4, 3.9, 6.1]


def test_insert_chart_no_title_clears_title(fake_word, excel_available):
    with wordlive.attach() as word:
        doc = word.documents.active
        doc.start.insert_chart("pie", {"A": 1, "B": 2})
        assert doc.charts[1]._shape().Chart.HasTitle is False


def test_insert_chart_before_and_after(fake_word, excel_available):
    with wordlive.attach() as word:
        doc = word.documents.active
        doc.start.insert_chart("bar", {"a": 1})
        doc.start.insert_chart("pie", {"b": 2}, where="before")
        assert len(doc.charts) == 2


# ---------------------------------------------------------------------------
# doc.charts collection + chart:N resolution + metadata reads
# ---------------------------------------------------------------------------


def test_charts_collection_and_metadata(fake_word, excel_available):
    with wordlive.attach() as word:
        doc = word.documents.active
        doc.start.insert_chart("bar", {"Q1": 10}, title="Sales")
        doc.start.insert_chart("scatter", [[1.0, 2.0]])
        assert len(doc.charts) == 2
        ids = [c.anchor_id for c in doc.charts]
        assert ids == ["chart:1", "chart:2"]

        rows = doc.charts.list()
        assert [r["anchor_id"] for r in rows] == ["chart:1", "chart:2"]
        assert rows[0]["kind"] == "bar"
        assert rows[0]["title"] == "Sales"
        assert rows[1]["kind"] == "scatter"
        assert rows[1]["title"] is None
        assert all("para" in r for r in rows)


def test_chart_anchor_type_and_title(fake_word, excel_available):
    with wordlive.attach() as word:
        doc = word.documents.active
        doc.start.insert_chart("line", {"Jan": 3.1}, title="Trend")
        chart = doc.charts[1]
        assert chart.chart_type == "line"
        assert chart.title == "Trend"


def test_chart_anchor_set_text_raises(fake_word, excel_available):
    with wordlive.attach() as word:
        doc = word.documents.active
        chart = doc.start.insert_chart("bar", {"a": 1})
        with pytest.raises(OpError, match="no plain text"):
            chart.set_text("x")


def test_anchor_by_id_resolves_chart(fake_word, excel_available):
    with wordlive.attach() as word:
        doc = word.documents.active
        doc.start.insert_chart("bar", {"a": 1})
        anchor = doc.anchor_by_id("chart:1")
        assert isinstance(anchor, wordlive.ChartAnchor)
        assert anchor.index == 1


def test_anchor_by_id_chart_bad_index(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        with pytest.raises(AnchorNotFoundError):
            doc.anchor_by_id("chart:banana")
        with pytest.raises(AnchorNotFoundError):
            doc.anchor_by_id("chart:7")  # out of range


def test_stats_counts_charts(fake_word, excel_available):
    with wordlive.attach() as word:
        doc = word.documents.active
        doc.start.insert_chart("bar", {"a": 1})
        assert doc.stats()["charts"] == 1


# ---------------------------------------------------------------------------
# exec op — registries + apply_op outputs
# ---------------------------------------------------------------------------


def test_insert_chart_in_op_registries():
    assert OP_REQUIRED_FIELDS["insert_chart"] == ("anchor_id", "kind", "data")
    optional = OP_OPTIONAL_FIELDS["insert_chart"]
    for field in ("title", "before", "after", "where"):
        assert field in optional


def test_insert_chart_op_reports_outputs(fake_word, excel_available):
    with wordlive.attach() as word:
        doc = word.documents.active
        result, exc = run_batch(
            doc,
            [
                {
                    "op": "insert_chart",
                    "anchor_id": "end",
                    "kind": "bar",
                    "data": {"Q1": 10, "Q2": 25},
                    "title": "Quarterly",
                }
            ],
            label="chart",
        )
    assert exc is None
    out = result["outputs"][0]
    assert out["op"] == "insert_chart"
    assert out["chart"] == 1
    assert out["anchor_id"] == "chart:1"


# ---------------------------------------------------------------------------
# MCP request builder
# ---------------------------------------------------------------------------


def test_mcp_build_write_op_insert_chart():
    from wordlive.mcp.server import _build_write_op

    op = _build_write_op(
        "insert_chart",
        {
            "anchor_id": "end",
            "kind": "scatter",
            "data": [[1.2, 3.4], [2.5, 6.1]],
            "title": "signal",
            "before": True,
        },
    )
    assert op == {
        "op": "insert_chart",
        "anchor_id": "end",
        "kind": "scatter",
        "data": [[1.2, 3.4], [2.5, 6.1]],
        "before": True,
        "title": "signal",
    }


def test_mcp_build_write_op_insert_chart_requires_fields():
    from wordlive.mcp.server import _build_write_op

    with pytest.raises(OpError):
        _build_write_op("insert_chart", {"anchor_id": "end", "kind": "bar"})  # no data
    with pytest.raises(OpError):
        _build_write_op("insert_chart", {"anchor_id": "end", "data": {"a": 1}})  # no kind


# ---------------------------------------------------------------------------
# CLI — insert-chart + charts list
# ---------------------------------------------------------------------------


def _invoke(args: list[str], *, input: str | None = None) -> tuple[int, str, str]:
    runner = CliRunner()
    result = runner.invoke(main, args, input=input, catch_exceptions=False)
    return result.exit_code, result.stdout, result.stderr


def test_cli_insert_chart(fake_word, monkeypatch):
    monkeypatch.setattr(_charts, "probe_excel_available", lambda: True)
    code, out, _ = _invoke(
        ["insert-chart", "--anchor-id", "end", "--kind", "bar", "--data", '{"Q1": 10, "Q2": 25}']
    )
    assert code == 0
    data = json.loads(out)
    assert data["ok"] is True
    assert data["chart_anchor_id"] == "chart:1"
    assert data["kind"] == "bar"


def test_cli_insert_chart_data_from_stdin(fake_word, monkeypatch):
    monkeypatch.setattr(_charts, "probe_excel_available", lambda: True)
    code, out, _ = _invoke(
        ["insert-chart", "--anchor-id", "end", "--kind", "scatter", "--data", "-"],
        input="[[1.2, 3.4], [2.5, 6.1]]",
    )
    assert code == 0
    assert json.loads(out)["chart_anchor_id"] == "chart:1"


def test_cli_insert_chart_bad_json(fake_word):
    code, _, err = _invoke(
        ["insert-chart", "--anchor-id", "end", "--kind", "bar", "--data", "{not json"]
    )
    assert code != 0
    assert "not valid JSON" in err


def test_cli_insert_chart_excel_missing_exit_6(fake_word, monkeypatch):
    from wordlive.cli.main import EXIT_EXCEL_NOT_AVAILABLE

    monkeypatch.setattr(_charts, "probe_excel_available", lambda: False)
    code, _, err = _invoke(
        ["insert-chart", "--anchor-id", "end", "--kind", "bar", "--data", '{"a": 1}']
    )
    assert code == EXIT_EXCEL_NOT_AVAILABLE
    assert "Excel" in err


def test_cli_charts_list_empty(fake_word):
    code, out, _ = _invoke(["charts"])
    assert code == 0
    assert json.loads(out) == []
