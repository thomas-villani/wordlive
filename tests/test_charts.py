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


# ===========================================================================
# Chart formatting & design — format / set_axis / add_trendline /
# set_series_color, across the Python API, exec ops, MCP builder, and CLI.
# ===========================================================================


def _insert(doc, kind="bar", data=None, title=None):
    """Insert a chart and return (ChartAnchor, the fake `Chart` COM object)."""
    data = {"Q1": 10, "Q2": 25, "Q3": 18} if data is None else data
    chart = doc.start.insert_chart(kind, data, title=title)
    return chart, doc.charts[chart.index]._shape().Chart


def test_formatting_vocab_maps_consistent():
    assert set(_charts.LEGEND_POSITIONS) == {"right", "left", "top", "bottom", "corner"}
    assert _charts.AXIS_WHICH["y"] == _charts.AXIS_WHICH["value"]
    assert _charts.AXIS_WHICH["x"] == _charts.AXIS_WHICH["category"]
    assert _charts.SCALE_TYPES["log"] == _charts.SCALE_TYPES["logarithmic"]
    assert set(_charts.TRENDLINE_KINDS) >= {"linear", "power", "exponential", "moving_average"}


# --- ChartAnchor.format (whole-chart / design) -----------------------------


def test_format_writes_passed_fields_only(fake_word, excel_available):
    from wordlive._format import to_bgr

    with wordlive.attach() as word:
        doc = word.documents.active
        chart, fake = _insert(doc)
        before_style = fake.ChartStyle
        ret = chart.format(
            title="Revenue",
            legend=True,
            legend_position="bottom",
            chart_style=242,
            background="#F4F6F7",
            plot_background="white",
            font="Calibri",
            font_size=10,
            font_color="#333333",
            data_labels=True,
            data_label_format="0.0",
        )
        assert ret is chart  # chainable
        assert fake.HasTitle is True and fake.ChartTitle.Text == "Revenue"
        assert fake.HasLegend is True
        assert fake.Legend.Position == int(_charts.XlLegendPosition.BOTTOM)
        assert fake.ChartStyle == 242 and before_style != 242
        assert fake.ChartArea.Format.Fill.ForeColor.RGB == to_bgr("#F4F6F7")
        assert fake.PlotArea.Format.Fill.ForeColor.RGB == to_bgr("white")
        assert fake.ChartArea.Font.Name == "Calibri"
        assert fake.ChartArea.Font.Size == 10.0
        assert fake.ChartArea.Font.Color == to_bgr("#333333")
        series = fake.SeriesCollection()(1)
        assert series.HasDataLabels is True
        assert series.DataLabels().NumberFormat == "0.0"


def test_format_no_args_leaves_chart_untouched(fake_word, excel_available):
    with wordlive.attach() as word:
        doc = word.documents.active
        chart, fake = _insert(doc, title="keep")
        chart.format()  # nothing passed
        assert fake.HasTitle is True and fake.ChartTitle.Text == "keep"
        assert fake.HasLegend is False


def test_format_title_none_clears(fake_word, excel_available):
    with wordlive.attach() as word:
        doc = word.documents.active
        chart, fake = _insert(doc, title="temp")
        assert fake.HasTitle is True
        chart.format(title=None)
        assert fake.HasTitle is False


def test_format_legend_position_implies_shown(fake_word, excel_available):
    with wordlive.attach() as word:
        doc = word.documents.active
        chart, fake = _insert(doc)
        chart.format(legend_position="right")
        assert fake.HasLegend is True
        assert fake.Legend.Position == int(_charts.XlLegendPosition.RIGHT)


def test_format_chart_type_retypes(fake_word, excel_available):
    with wordlive.attach() as word:
        doc = word.documents.active
        chart, fake = _insert(doc, "pie", {"a": 1, "b": 2})
        chart.format(chart_type="line")
        assert fake.ChartType == int(_charts.KIND_TO_XL["line"])
        assert chart.chart_type == "line"


def test_format_bad_inputs_raise_operror(fake_word, excel_available):
    with wordlive.attach() as word:
        doc = word.documents.active
        chart, _ = _insert(doc)
        with pytest.raises(OpError, match="legend_position"):
            chart.format(legend_position="sideways")
        with pytest.raises(OpError, match="chart_type"):
            chart.format(chart_type="donut")
        with pytest.raises(OpError, match="colour|color"):
            chart.format(background="not-a-colour")


# --- ChartAnchor.set_axis --------------------------------------------------


def test_set_axis_value_fields(fake_word, excel_available):
    with wordlive.attach() as word:
        doc = word.documents.active
        chart, fake = _insert(doc)
        chart.set_axis(
            "value",
            title="USD",
            minimum=0,
            maximum=30,
            number_format="#,##0",
            gridlines=True,
        )
        axis = fake.Axes(int(_charts.XlAxisType.VALUE), 1)
        assert axis.HasTitle is True and axis.AxisTitle.Text == "USD"
        assert axis.MinimumScale == 0.0 and axis.MaximumScale == 30.0
        assert axis.TickLabels.NumberFormat == "#,##0"
        assert axis.HasMajorGridlines is True


def test_set_axis_log_scale_and_which_aliases(fake_word, excel_available):
    with wordlive.attach() as word:
        doc = word.documents.active
        chart, fake = _insert(doc, "scatter", [[1.0, 2.0]])
        chart.set_axis("y", scale="log")  # y is an alias for value
        axis = fake.Axes(int(_charts.XlAxisType.VALUE), 1)
        assert axis.ScaleType == int(_charts.XlScaleType.LOGARITHMIC)
        chart.set_axis("x", title="t")  # x is an alias for category
        cat = fake.Axes(int(_charts.XlAxisType.CATEGORY), 1)
        assert cat.AxisTitle.Text == "t"


def test_set_axis_title_none_clears(fake_word, excel_available):
    with wordlive.attach() as word:
        doc = word.documents.active
        chart, fake = _insert(doc)
        chart.set_axis("value", title="x")
        chart.set_axis("value", title=None)
        assert fake.Axes(int(_charts.XlAxisType.VALUE), 1).HasTitle is False


def test_set_axis_bad_inputs_raise(fake_word, excel_available):
    with wordlive.attach() as word:
        doc = word.documents.active
        chart, _ = _insert(doc)
        with pytest.raises(OpError, match="axis"):
            chart.set_axis("diagonal")
        with pytest.raises(OpError, match="scale"):
            chart.set_axis("value", scale="bendy")


# --- ChartAnchor.add_trendline ---------------------------------------------


def test_add_trendline(fake_word, excel_available):
    with wordlive.attach() as word:
        doc = word.documents.active
        chart, fake = _insert(doc, "scatter", [[1.0, 2.0], [2.0, 5.0]])
        chart.add_trendline(kind="power", display_equation=True, display_r_squared=True, forward=2)
        tl = fake.SeriesCollection()(1).trendlines[-1]
        assert tl.Type == int(_charts.XlTrendlineType.POWER)
        assert tl.DisplayEquation is True and tl.DisplayRSquared is True
        assert tl.Forward == 2.0


def test_add_trendline_bad_kind_raises(fake_word, excel_available):
    with wordlive.attach() as word:
        doc = word.documents.active
        chart, _ = _insert(doc)
        with pytest.raises(OpError, match="trendline kind"):
            chart.add_trendline(kind="magic")


# --- ChartAnchor.set_series_color ------------------------------------------


def test_set_series_color_series_and_point(fake_word, excel_available):
    from wordlive._format import to_bgr

    with wordlive.attach() as word:
        doc = word.documents.active
        chart, fake = _insert(doc)
        chart.set_series_color("#2E86C1")
        series = fake.SeriesCollection()(1)
        assert series.Format.Fill.ForeColor.RGB == to_bgr("#2E86C1")
        assert series.Format.Line.ForeColor.RGB == to_bgr("#2E86C1")  # line set too
        chart.set_series_color((231, 76, 60), point=2)
        assert series.Points(2).Format.Fill.ForeColor.RGB == to_bgr((231, 76, 60))


def test_set_series_color_bad_colour_raises(fake_word, excel_available):
    with wordlive.attach() as word:
        doc = word.documents.active
        chart, _ = _insert(doc)
        with pytest.raises(OpError, match="colour|color"):
            chart.set_series_color("chartreusey")


# --- read props -------------------------------------------------------------


def test_chart_read_props_and_list(fake_word, excel_available):
    with wordlive.attach() as word:
        doc = word.documents.active
        chart, _ = _insert(doc)
        chart.format(chart_style=242, legend=True)
        assert chart.chart_style == 242
        assert chart.has_legend is True
        row = doc.charts.list()[0]
        assert row["chart_style"] == 242 and row["has_legend"] is True


# --- exec ops ---------------------------------------------------------------


def test_format_ops_in_registries():
    assert OP_REQUIRED_FIELDS["format_chart"] == ("anchor_id",)
    assert OP_REQUIRED_FIELDS["format_axis"] == ("anchor_id", "which")
    assert OP_REQUIRED_FIELDS["set_series_color"] == ("anchor_id", "color")
    for f in ("chart_style", "legend", "chart_type"):
        assert f in OP_OPTIONAL_FIELDS["format_chart"]
    assert "scale" in OP_OPTIONAL_FIELDS["format_axis"]
    assert "kind" in OP_OPTIONAL_FIELDS["add_trendline"]
    assert "point" in OP_OPTIONAL_FIELDS["set_series_color"]


def test_exec_format_chart_applies(fake_word, excel_available):
    with wordlive.attach() as word:
        doc = word.documents.active
        _, fake = _insert(doc)
        result, exc = run_batch(
            doc,
            [{"op": "format_chart", "anchor_id": "chart:1", "title": "T", "chart_style": 240}],
            label="fmt",
        )
        assert exc is None
        assert fake.ChartTitle.Text == "T" and fake.ChartStyle == 240


def test_exec_format_axis_and_trendline_and_color(fake_word, excel_available):
    with wordlive.attach() as word:
        doc = word.documents.active
        _, fake = _insert(doc, "scatter", [[1.0, 2.0]])
        result, exc = run_batch(
            doc,
            [
                {"op": "format_axis", "anchor_id": "chart:1", "which": "value", "scale": "log"},
                {"op": "add_trendline", "anchor_id": "chart:1", "kind": "linear"},
                {"op": "set_series_color", "anchor_id": "chart:1", "color": "navy", "point": 1},
            ],
            label="fmt",
        )
        assert exc is None
        assert fake.Axes(int(_charts.XlAxisType.VALUE), 1).ScaleType == int(
            _charts.XlScaleType.LOGARITHMIC
        )
        assert len(fake.SeriesCollection()(1).trendlines) == 1


def test_exec_format_chart_rejects_non_chart_anchor(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        result, exc = run_batch(
            doc, [{"op": "format_chart", "anchor_id": "start", "title": "x"}], label="fmt"
        )
        assert isinstance(exc, OpError)
        assert "not a chart" in str(exc)


# --- MCP builder ------------------------------------------------------------


def test_mcp_build_format_ops():
    from wordlive.mcp.server import _build_write_op

    assert _build_write_op(
        "format_chart", {"anchor_id": "chart:1", "chart_style": 240, "legend": True}
    ) == {"op": "format_chart", "anchor_id": "chart:1", "chart_style": 240, "legend": True}
    assert _build_write_op(
        "format_axis", {"anchor_id": "chart:1", "which": "value", "scale": "log"}
    ) == {"op": "format_axis", "anchor_id": "chart:1", "which": "value", "scale": "log"}
    assert _build_write_op(
        "set_series_color", {"anchor_id": "chart:1", "color": "red", "point": 2}
    ) == {"op": "set_series_color", "anchor_id": "chart:1", "color": "red", "point": 2}


def test_mcp_build_format_ops_require_fields():
    from wordlive.mcp.server import _build_write_op

    with pytest.raises(OpError):
        _build_write_op("format_axis", {"anchor_id": "chart:1"})  # no which
    with pytest.raises(OpError):
        _build_write_op("set_series_color", {"anchor_id": "chart:1"})  # no color


# --- CLI --------------------------------------------------------------------


def test_cli_format_chart(fake_word, excel_available):
    with wordlive.attach() as word:
        word.documents.active.start.insert_chart("bar", {"a": 1})
    code, out, _ = _invoke(
        [
            "format-chart",
            "--anchor-id",
            "chart:1",
            "--chart-style",
            "240",
            "--legend",
            "--title",
            "Q",
        ]
    )
    assert code == 0
    data = json.loads(out)
    assert data["ok"] is True and data["applied"]["chart_style"] == 240


def test_cli_format_axis(fake_word, excel_available):
    with wordlive.attach() as word:
        word.documents.active.start.insert_chart("scatter", [[1.0, 2.0]])
    code, out, _ = _invoke(
        ["format-axis", "--anchor-id", "chart:1", "--which", "value", "--scale", "log"]
    )
    assert code == 0
    assert json.loads(out)["which"] == "value"


def test_cli_add_trendline(fake_word, excel_available):
    with wordlive.attach() as word:
        word.documents.active.start.insert_chart("scatter", [[1.0, 2.0]])
    code, out, _ = _invoke(
        ["add-trendline", "--anchor-id", "chart:1", "--kind", "power", "--display-equation"]
    )
    assert code == 0
    assert json.loads(out)["applied"]["kind"] == "power"


def test_cli_set_series_color(fake_word, excel_available):
    with wordlive.attach() as word:
        word.documents.active.start.insert_chart("pie", {"a": 1, "b": 2})
    code, out, _ = _invoke(
        ["set-series-color", "--anchor-id", "chart:1", "--color", "#2E86C1", "--point", "1"]
    )
    assert code == 0
    assert json.loads(out)["point"] == 1


def test_cli_format_chart_rejects_non_chart(fake_word):
    code, _, err = _invoke(["format-chart", "--anchor-id", "start", "--title", "x"])
    assert code != 0
    assert "not a chart" in err


# --- PR C: chart depth ------------------------------------------------------
# format-chart gap/overlap/data-table, format_series, add_error_bars,
# trendline order/period. Probed live 2026-06-21; the fake mirrors the contract.


def test_format_chart_bar_spacing_and_data_table(fake_word, excel_available):
    with wordlive.attach() as word:
        doc = word.documents.active
        chart, fake = _insert(doc)
        chart.format(gap_width=40, overlap=20, data_table=True)
        grp = fake.ChartGroups(1)
        assert grp.GapWidth == 40 and grp.Overlap == 20
        assert fake.HasDataTable is True


def test_format_series_markers_and_smooth(fake_word, excel_available):
    with wordlive.attach() as word:
        doc = word.documents.active
        chart, fake = _insert(doc, "line", {"a": 1, "b": 2})
        chart.format_series(series=1, marker="circle", marker_size=8, smooth=True)
        s = fake.SeriesCollection()(1)
        assert s.MarkerStyle == int(_charts.XlMarkerStyle.CIRCLE)
        assert s.MarkerSize == 8 and s.Smooth is True


def test_format_series_marker_int_passthrough(fake_word, excel_available):
    with wordlive.attach() as word:
        doc = word.documents.active
        chart, fake = _insert(doc, "scatter", [[1.0, 2.0]])
        chart.format_series(series=1, marker=2)  # raw XlMarkerStyle int
        assert fake.SeriesCollection()(1).MarkerStyle == 2


def test_format_series_point_explosion_and_label_font(fake_word, excel_available):
    from wordlive._format import to_bgr

    with wordlive.attach() as word:
        doc = word.documents.active
        chart, fake = _insert(doc, "pie", {"a": 1, "b": 2})
        chart.format_series(
            series=1, point=1, explosion=25, data_label_size=12, data_label_color="red"
        )
        pt = fake.SeriesCollection()(1).Points(1)
        assert pt.Explosion == 25
        assert pt.DataLabel.Font.Size == 12.0
        assert pt.DataLabel.Font.Color == to_bgr("red")


def test_format_series_data_labels_toggle_then_font(fake_word, excel_available):
    with wordlive.attach() as word:
        doc = word.documents.active
        chart, fake = _insert(doc)
        chart.format_series(series=1, data_labels=True, data_label_size=14)
        s = fake.SeriesCollection()(1)
        assert s.HasDataLabels is True
        assert s.DataLabels().Font.Size == 14.0


def test_format_series_bad_marker_raises(fake_word, excel_available):
    with wordlive.attach() as word:
        doc = word.documents.active
        chart, _ = _insert(doc, "line", {"a": 1})
        with pytest.raises(OpError, match="marker"):
            chart.format_series(series=1, marker="hexagram")


def test_add_error_bars_fixed(fake_word, excel_available):
    with wordlive.attach() as word:
        doc = word.documents.active
        chart, fake = _insert(doc)
        chart.add_error_bars(series=1, kind="percent", amount=5, include="plus", axis="y")
        s = fake.SeriesCollection()(1)
        assert s.HasErrorBars is True
        direction, include, type_, amount = s.error_bars[-1]
        assert direction == int(_charts.XlErrorBarDirection.Y)
        assert include == int(_charts.XlErrorBarInclude.PLUS_VALUES)
        assert type_ == int(_charts.XlErrorBarType.PERCENT)
        assert amount == 5.0


def test_add_error_bars_sterror_needs_no_amount(fake_word, excel_available):
    with wordlive.attach() as word:
        doc = word.documents.active
        chart, fake = _insert(doc)
        chart.add_error_bars(series=1, kind="sterror")  # Word computes the amount
        assert fake.SeriesCollection()(1).HasErrorBars is True


def test_add_error_bars_amount_required_and_bad_kind(fake_word, excel_available):
    with wordlive.attach() as word:
        doc = word.documents.active
        chart, _ = _insert(doc)
        with pytest.raises(OpError, match="amount"):
            chart.add_error_bars(series=1, kind="fixed")  # fixed needs an amount
        with pytest.raises(OpError, match="error-bar kind"):
            chart.add_error_bars(series=1, kind="bogus", amount=1)


def test_add_trendline_order_and_period(fake_word, excel_available):
    with wordlive.attach() as word:
        doc = word.documents.active
        chart, fake = _insert(doc, "scatter", [[1.0, 2.0], [2.0, 5.0]])
        chart.add_trendline(kind="polynomial", order=3)
        chart.add_trendline(kind="moving_average", period=2)
        tls = fake.SeriesCollection()(1).trendlines
        assert tls[-2].Order == 3
        assert tls[-1].Period == 2


def test_pr_c_ops_in_registries():
    assert OP_REQUIRED_FIELDS["format_series"] == ("anchor_id",)
    assert OP_REQUIRED_FIELDS["add_error_bars"] == ("anchor_id",)
    for f in ("marker", "marker_size", "smooth", "explosion", "data_label_color"):
        assert f in OP_OPTIONAL_FIELDS["format_series"]
    for f in ("kind", "amount", "include", "axis"):
        assert f in OP_OPTIONAL_FIELDS["add_error_bars"]
    for f in ("gap_width", "overlap", "data_table"):
        assert f in OP_OPTIONAL_FIELDS["format_chart"]
    for f in ("order", "period"):
        assert f in OP_OPTIONAL_FIELDS["add_trendline"]


def test_exec_format_series_and_error_bars(fake_word, excel_available):
    with wordlive.attach() as word:
        doc = word.documents.active
        _, fake = _insert(doc, "line", {"a": 1, "b": 2})
        result, exc = run_batch(
            doc,
            [
                {"op": "format_series", "anchor_id": "chart:1", "series": 1, "marker": "square"},
                {
                    "op": "add_error_bars",
                    "anchor_id": "chart:1",
                    "series": 1,
                    "kind": "fixed",
                    "amount": 2,
                },
            ],
            label="depth",
        )
        assert exc is None
        s = fake.SeriesCollection()(1)
        assert s.MarkerStyle == int(_charts.XlMarkerStyle.SQUARE)
        assert s.HasErrorBars is True and s.error_bars[-1][3] == 2.0


def test_mcp_build_pr_c_ops():
    from wordlive.mcp.server import _build_write_op

    assert _build_write_op(
        "format_series", {"anchor_id": "chart:1", "marker": "circle", "marker_size": 8}
    ) == {"op": "format_series", "anchor_id": "chart:1", "marker": "circle", "marker_size": 8}
    assert _build_write_op(
        "add_error_bars", {"anchor_id": "chart:1", "kind": "percent", "amount": 5}
    ) == {"op": "add_error_bars", "anchor_id": "chart:1", "kind": "percent", "amount": 5}


def test_cli_format_series(fake_word, excel_available):
    with wordlive.attach() as word:
        word.documents.active.start.insert_chart("line", {"a": 1, "b": 2})
    code, out, _ = _invoke(
        ["format-series", "--anchor-id", "chart:1", "--marker", "circle", "--marker-size", "8"]
    )
    assert code == 0
    assert json.loads(out)["ok"] is True


def test_cli_format_series_needs_an_option(fake_word, excel_available):
    with wordlive.attach() as word:
        word.documents.active.start.insert_chart("line", {"a": 1})
    code, _, err = _invoke(["format-series", "--anchor-id", "chart:1"])
    assert code != 0
    assert "at least one" in err


def test_cli_add_error_bars(fake_word, excel_available):
    with wordlive.attach() as word:
        word.documents.active.start.insert_chart("bar", {"a": 1, "b": 2})
    code, out, _ = _invoke(
        ["add-error-bars", "--anchor-id", "chart:1", "--kind", "percent", "--amount", "5"]
    )
    assert code == 0
    assert json.loads(out)["applied"]["kind"] == "percent"
