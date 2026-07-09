"""Chart insertion and formatting."""

from __future__ import annotations

import json
from typing import Any

import click

from ... import attach
from ..._ops import pick_doc as _pick_doc
from ..main import _run, emit
from ._common import (
    _chart_anchor,
    _fmt_charts,
    _parse_color,
)


@click.command(name="charts")
@click.pass_context
def charts_cmd(ctx: click.Context) -> None:
    """List the document's charts (chart:N id, kind, title, para:N).

    The discovery half of charting: see what charts are in the document and
    address them by `chart:N`. Metadata only (the series data is static) and
    non-mutating.
    """

    def go() -> None:
        with attach() as word:
            doc = _pick_doc(word, ctx.obj["doc_name"])
            rows = doc.charts.list()
            emit(rows, as_text=not ctx.obj["as_json"], text=_fmt_charts(rows))

    _run(ctx, go)


@click.command(name="insert-chart")
@click.option(
    "--anchor-id", "anchor_id", required=True, help="Anchor to insert the chart relative to."
)
@click.option(
    "--kind",
    "kind",
    required=True,
    type=click.Choice(["bar", "pie", "line", "scatter"]),
    help="Chart kind: bar (clustered columns), pie, line, or scatter.",
)
@click.option(
    "--data",
    "data",
    required=True,
    help="Chart data as JSON, or '-' to read it from stdin: an object "
    '{"label": value} (bar/pie/line) or an array of [x, y] pairs (scatter/line).',
)
@click.option("--title", "title", default=None, help="Chart title (and series name).")
@click.option(
    "--before/--after",
    "before",
    default=False,
    show_default="--after",
    help="Insert before the anchor instead of after it.",
)
@click.pass_context
def insert_chart_cmd(
    ctx: click.Context,
    anchor_id: str,
    kind: str,
    data: str,
    title: str | None,
    before: bool,
) -> None:
    """Insert an Excel-backed chart at any anchor (atomic-undo).

    --data is JSON (or '-' for stdin): an object {"Q1": 10, "Q2": 25} for
    bar/pie/line, or an array of [x, y] pairs [[1.2, 3.4], [2.5, 6.1]] for
    scatter (both axes numeric; duplicate x preserved). Charts embed a hidden
    Excel workbook, so Excel must be installed (exit 6 if not); the data link is
    then broken, so the chart's data is static.
    """
    raw = click.get_text_stream("stdin").read() if data == "-" else data
    try:
        parsed = json.loads(raw)
    except json.JSONDecodeError as e:
        raise click.UsageError(f"--data is not valid JSON: {e}") from e
    where = "before" if before else "after"

    def go() -> None:
        with attach() as word:
            doc = _pick_doc(word, ctx.obj["doc_name"])
            anchor = doc.anchor_by_id(anchor_id)
            with doc.edit(f"CLI: insert chart {where} {anchor_id}"):
                chart = anchor.insert_chart(kind, parsed, title=title, where=where)
            emit(
                {
                    "ok": True,
                    "anchor_id": anchor_id,
                    "chart": chart.index,
                    "chart_anchor_id": chart.anchor_id,
                    "kind": kind,
                    "where": where,
                },
                as_text=not ctx.obj["as_json"],
                text=f"inserted {chart.anchor_id} ({kind}) {where} {anchor_id}",
            )

    _run(ctx, go)


@click.command(name="format-chart")
@click.option("--anchor-id", "anchor_id", required=True, help="Chart anchor (chart:N) to format.")
@click.option("--title", "title", default=None, help='Chart title (pass "" to clear it).')
@click.option("--legend/--no-legend", "legend", default=None, help="Show or hide the legend.")
@click.option(
    "--legend-position",
    "legend_position",
    type=click.Choice(["right", "left", "top", "bottom", "corner"]),
    default=None,
    help="Where the legend sits (implies it is shown).",
)
@click.option(
    "--chart-style", "chart_style", type=int, default=None, help="Design-gallery style id."
)
@click.option("--background", "background", default=None, help="Chart-area fill colour.")
@click.option("--plot-background", "plot_background", default=None, help="Plot-area fill colour.")
@click.option("--font", "font", default=None, help="Whole-chart font family.")
@click.option("--font-size", "font_size", default=None, help="Whole-chart font size (pt or unit).")
@click.option("--font-color", "font_color", default=None, help="Whole-chart font colour.")
@click.option(
    "--data-labels/--no-data-labels", "data_labels", default=None, help="Show point data labels."
)
@click.option(
    "--data-label-format", "data_label_format", default=None, help="Data-label number format."
)
@click.option(
    "--chart-type",
    "chart_type",
    type=click.Choice(["bar", "pie", "line", "scatter"]),
    default=None,
    help="Re-type the chart in place.",
)
@click.option(
    "--gap-width", "gap_width", type=int, default=None, help="Bar gap width (bar/column charts)."
)
@click.option(
    "--overlap", "overlap", type=int, default=None, help="Bar overlap (bar/column charts)."
)
@click.option(
    "--data-table/--no-data-table",
    "data_table",
    default=None,
    help="Show the data-table grid beneath the plot.",
)
@click.pass_context
def format_chart_cmd(
    ctx: click.Context,
    anchor_id: str,
    title: str | None,
    legend: bool | None,
    legend_position: str | None,
    chart_style: int | None,
    background: str | None,
    plot_background: str | None,
    font: str | None,
    font_size: str | None,
    font_color: str | None,
    data_labels: bool | None,
    data_label_format: str | None,
    chart_type: str | None,
    gap_width: int | None,
    overlap: int | None,
    data_table: bool | None,
) -> None:
    """Apply whole-chart / design formatting to a chart (atomic-undo).

    Operates on the static chart — no Excel needed. Colours are a name, hex
    (#2E86C1), or comma-separated r,g,b. Pass at least one option.
    """
    raw: dict[str, Any] = {
        "title": title,
        "legend": legend,
        "legend_position": legend_position,
        "chart_style": chart_style,
        "background": _parse_color(background),
        "plot_background": _parse_color(plot_background),
        "font": font,
        "font_size": font_size,
        "font_color": _parse_color(font_color),
        "data_labels": data_labels,
        "data_label_format": data_label_format,
        "chart_type": chart_type,
        "gap_width": gap_width,
        "overlap": overlap,
        "data_table": data_table,
    }
    kwargs = {k: v for k, v in raw.items() if v is not None}
    if not kwargs:
        raise click.UsageError("pass at least one formatting option")

    def go() -> None:
        with attach() as word:
            doc, anchor = _chart_anchor(word, ctx.obj["doc_name"], anchor_id)
            with doc.edit(f"CLI: format chart {anchor_id}"):
                anchor.format(**kwargs)
            emit(
                {"ok": True, "anchor_id": anchor_id, "applied": kwargs},
                as_text=not ctx.obj["as_json"],
                text=f"formatted {anchor_id}: {kwargs}",
            )

    _run(ctx, go)


@click.command(name="format-axis")
@click.option("--anchor-id", "anchor_id", required=True, help="Chart anchor (chart:N) to format.")
@click.option(
    "--which",
    "which",
    required=True,
    type=click.Choice(["value", "y", "category", "x"]),
    help="Which axis: value/y or category/x.",
)
@click.option("--title", "title", default=None, help='Axis title (pass "" to clear it).')
@click.option("--minimum", "minimum", type=float, default=None, help="Axis minimum.")
@click.option("--maximum", "maximum", type=float, default=None, help="Axis maximum.")
@click.option(
    "--scale",
    "scale",
    type=click.Choice(["linear", "log"]),
    default=None,
    help="Axis scale type.",
)
@click.option("--number-format", "number_format", default=None, help="Tick-label number format.")
@click.option("--gridlines/--no-gridlines", "gridlines", default=None, help="Show major gridlines.")
@click.pass_context
def format_axis_cmd(
    ctx: click.Context,
    anchor_id: str,
    which: str,
    title: str | None,
    minimum: float | None,
    maximum: float | None,
    scale: str | None,
    number_format: str | None,
    gridlines: bool | None,
) -> None:
    """Format one axis of a chart (atomic-undo). `log` scale suits order-of-magnitude data."""
    raw: dict[str, Any] = {
        "title": title,
        "minimum": minimum,
        "maximum": maximum,
        "scale": scale,
        "number_format": number_format,
        "gridlines": gridlines,
    }
    kwargs = {k: v for k, v in raw.items() if v is not None}
    if not kwargs:
        raise click.UsageError("pass at least one axis option")

    def go() -> None:
        with attach() as word:
            doc, anchor = _chart_anchor(word, ctx.obj["doc_name"], anchor_id)
            with doc.edit(f"CLI: format axis {anchor_id}"):
                anchor.set_axis(which, **kwargs)
            emit(
                {"ok": True, "anchor_id": anchor_id, "which": which, "applied": kwargs},
                as_text=not ctx.obj["as_json"],
                text=f"formatted {which} axis of {anchor_id}: {kwargs}",
            )

    _run(ctx, go)


@click.command(name="add-trendline")
@click.option("--anchor-id", "anchor_id", required=True, help="Chart anchor (chart:N).")
@click.option("--series", "series", type=int, default=1, show_default=True, help="1-based series.")
@click.option(
    "--kind",
    "kind",
    type=click.Choice(
        ["linear", "exponential", "logarithmic", "moving_average", "polynomial", "power"]
    ),
    default="linear",
    show_default=True,
    help="Trendline curve type.",
)
@click.option(
    "--display-equation", "display_equation", is_flag=True, default=False, help="Show the equation."
)
@click.option(
    "--display-r-squared", "display_r_squared", is_flag=True, default=False, help="Show R²."
)
@click.option("--forward", "forward", type=float, default=None, help="Forecast forward N units.")
@click.option("--backward", "backward", type=float, default=None, help="Forecast backward N units.")
@click.option(
    "--order", "order", type=int, default=None, help="Polynomial degree 2–6 (kind=polynomial)."
)
@click.option(
    "--period",
    "period",
    type=int,
    default=None,
    help="Moving-average window (kind=moving_average).",
)
@click.pass_context
def add_trendline_cmd(
    ctx: click.Context,
    anchor_id: str,
    series: int,
    kind: str,
    display_equation: bool,
    display_r_squared: bool,
    forward: float | None,
    backward: float | None,
    order: int | None,
    period: int | None,
) -> None:
    """Fit a trendline to a chart series (atomic-undo).

    A power/exponential fit with --display-equation draws the law of best fit.
    """
    kwargs: dict[str, Any] = {
        "series": series,
        "kind": kind,
        "display_equation": display_equation,
        "display_r_squared": display_r_squared,
    }
    if forward is not None:
        kwargs["forward"] = forward
    if backward is not None:
        kwargs["backward"] = backward
    if order is not None:
        kwargs["order"] = order
    if period is not None:
        kwargs["period"] = period

    def go() -> None:
        with attach() as word:
            doc, anchor = _chart_anchor(word, ctx.obj["doc_name"], anchor_id)
            with doc.edit(f"CLI: add trendline {anchor_id}"):
                anchor.add_trendline(**kwargs)
            emit(
                {"ok": True, "anchor_id": anchor_id, "applied": kwargs},
                as_text=not ctx.obj["as_json"],
                text=f"added {kind} trendline to {anchor_id} series {series}",
            )

    _run(ctx, go)


@click.command(name="set-series-color")
@click.option("--anchor-id", "anchor_id", required=True, help="Chart anchor (chart:N).")
@click.option(
    "--color", "color", required=True, help="Colour: name, hex (#2E86C1), or comma-separated r,g,b."
)
@click.option("--series", "series", type=int, default=1, show_default=True, help="1-based series.")
@click.option(
    "--point",
    "point",
    type=int,
    default=None,
    help="1-based point/slice to recolour (omit to colour the whole series).",
)
@click.pass_context
def set_series_color_cmd(
    ctx: click.Context, anchor_id: str, color: str, series: int, point: int | None
) -> None:
    """Recolour a chart series, or a single point / pie slice (atomic-undo)."""
    parsed = _parse_color(color)

    def go() -> None:
        with attach() as word:
            doc, anchor = _chart_anchor(word, ctx.obj["doc_name"], anchor_id)
            with doc.edit(f"CLI: set series color {anchor_id}"):
                anchor.set_series_color(parsed, series=series, point=point)
            target = f"point {point}" if point is not None else f"series {series}"
            emit(
                {
                    "ok": True,
                    "anchor_id": anchor_id,
                    "series": series,
                    "point": point,
                    "color": color,
                },
                as_text=not ctx.obj["as_json"],
                text=f"recoloured {target} of {anchor_id} -> {color}",
            )

    _run(ctx, go)


@click.command(name="format-series")
@click.option("--anchor-id", "anchor_id", required=True, help="Chart anchor (chart:N).")
@click.option("--series", "series", type=int, default=1, show_default=True, help="1-based series.")
@click.option(
    "--point",
    "point",
    type=int,
    default=None,
    help="1-based point/slice to target (omit for the whole series).",
)
@click.option(
    "--marker",
    "marker",
    default=None,
    help="Marker glyph: circle/square/diamond/triangle/x/star/dot/dash/plus/none/auto.",
)
@click.option("--marker-size", "marker_size", type=int, default=None, help="Marker size 2–72.")
@click.option(
    "--smooth/--no-smooth", "smooth", default=None, help="Curve a line/scatter through its points."
)
@click.option(
    "--explosion", "explosion", type=int, default=None, help="Pull a pie slice out 0–400."
)
@click.option(
    "--data-labels/--no-data-labels",
    "data_labels",
    default=None,
    help="Show this series' point labels.",
)
@click.option(
    "--data-label-size", "data_label_size", type=float, default=None, help="Data-label font size."
)
@click.option(
    "--data-label-color", "data_label_color", default=None, help="Data-label font colour."
)
@click.pass_context
def format_series_cmd(
    ctx: click.Context,
    anchor_id: str,
    series: int,
    point: int | None,
    marker: str | None,
    marker_size: int | None,
    smooth: bool | None,
    explosion: int | None,
    data_labels: bool | None,
    data_label_size: float | None,
    data_label_color: str | None,
) -> None:
    """Format a chart series, or a single point / slice (atomic-undo).

    Markers and --smooth suit line/scatter; --explosion a pie slice. Colours are
    a name, hex (#2E86C1), or comma-separated r,g,b. Pass at least one option.
    """
    raw: dict[str, Any] = {
        "marker": marker,
        "marker_size": marker_size,
        "smooth": smooth,
        "explosion": explosion,
        "data_labels": data_labels,
        "data_label_size": data_label_size,
        "data_label_color": _parse_color(data_label_color),
    }
    kwargs = {k: v for k, v in raw.items() if v is not None}
    if not kwargs:
        raise click.UsageError("pass at least one formatting option")
    kwargs["series"] = series
    kwargs["point"] = point

    def go() -> None:
        with attach() as word:
            doc, anchor = _chart_anchor(word, ctx.obj["doc_name"], anchor_id)
            with doc.edit(f"CLI: format series {anchor_id}"):
                anchor.format_series(**kwargs)
            target = f"point {point}" if point is not None else f"series {series}"
            emit(
                {"ok": True, "anchor_id": anchor_id, "series": series, "point": point},
                as_text=not ctx.obj["as_json"],
                text=f"formatted {target} of {anchor_id}",
            )

    _run(ctx, go)


@click.command(name="add-error-bars")
@click.option("--anchor-id", "anchor_id", required=True, help="Chart anchor (chart:N).")
@click.option("--series", "series", type=int, default=1, show_default=True, help="1-based series.")
@click.option(
    "--kind",
    "kind",
    type=click.Choice(["fixed", "percent", "stdev", "sterror"]),
    default="fixed",
    show_default=True,
    help="How the error amount is computed.",
)
@click.option(
    "--amount",
    "amount",
    type=float,
    default=None,
    help="Error magnitude (required unless kind=sterror).",
)
@click.option(
    "--include",
    "include",
    type=click.Choice(["both", "plus", "minus"]),
    default="both",
    show_default=True,
    help="Which side(s) to draw.",
)
@click.option(
    "--axis",
    "axis",
    type=click.Choice(["y", "value", "x", "category"]),
    default="y",
    show_default=True,
    help="Which axis the bars run along.",
)
@click.pass_context
def add_error_bars_cmd(
    ctx: click.Context,
    anchor_id: str,
    series: int,
    kind: str,
    amount: float | None,
    include: str,
    axis: str,
) -> None:
    """Draw error bars on a chart series (atomic-undo)."""
    kwargs: dict[str, Any] = {"series": series, "kind": kind, "include": include, "axis": axis}
    if amount is not None:
        kwargs["amount"] = amount

    def go() -> None:
        with attach() as word:
            doc, anchor = _chart_anchor(word, ctx.obj["doc_name"], anchor_id)
            with doc.edit(f"CLI: add error bars {anchor_id}"):
                anchor.add_error_bars(**kwargs)
            emit(
                {"ok": True, "anchor_id": anchor_id, "series": series, "applied": kwargs},
                as_text=not ctx.obj["as_json"],
                text=f"added {kind} error bars to {anchor_id} series {series}",
            )

    _run(ctx, go)
