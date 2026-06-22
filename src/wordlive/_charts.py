"""Excel-backed charts — `InlineShapes.AddChart2` + embedded-workbook population.

Charts are the one wordlive feature that reaches a *second* Office app: Word's
`Range.InlineShapes.AddChart2` embeds a chart whose data lives in a hidden Excel
workbook, so an `Excel.Application` COM server must be installed. The live-Word
gotchas this module encodes (hard-won by live probing, 2026-06-16):

- **AddChart2 only works off the `Selection`**, never an arbitrary `Range`
  (`Range.InlineShapes.AddChart2` raises "Requested object is not available"). The
  caller selects the insertion point first; `doc.edit()` restores the user's
  selection afterwards (politeness).
- **Populate via workbook cells + a `=SERIES(...)` formula string**, not the
  `Series.XValues` / `.Values` array setters — those are unreliable under pywin32
  late binding ("Property XValues can not be set"). Writing numbers into cells
  also keeps a scatter's x numeric: a literal `{...}` array stores the x values as
  *text*, which an XY chart then plots at category positions 1, 2, 3… instead.
- **`ChartData.BreakLink()` before closing the workbook**, so the chart's data
  goes static and Word drops its reference to the embedded Excel — otherwise that
  hidden Excel never terminates and orphans accumulate until Word locks up ("The
  chart data grid is already open in ..."). The close order is BreakLink →
  Workbook.Close → Application.Quit (only when it has no workbooks left).
- Reading a chart's *series* data back destabilises Word, so metadata reads
  (`doc.charts`) touch only `ChartType` / `ChartTitle` — never `SeriesCollection`.
"""

from __future__ import annotations

from collections.abc import Mapping
from typing import Any

from ._format import to_bgr
from .constants import (
    XlAxisGroup,
    XlAxisType,
    XlChartType,
    XlErrorBarDirection,
    XlErrorBarInclude,
    XlErrorBarType,
    XlLegendPosition,
    XlMarkerStyle,
    XlScaleType,
    XlTrendlineType,
)
from .exceptions import OpError

# Sentinel for "argument not supplied" — lets a formatting field tell "leave
# untouched" apart from an explicit `None` that means "clear" (e.g. a title).
_UNSET: Any = object()

# `insert_chart(kind=...)` string → Excel chart-type enum. Keep narrow.
KIND_TO_XL: dict[str, XlChartType] = {
    "bar": XlChartType.COLUMN_CLUSTERED,
    "pie": XlChartType.PIE,
    "line": XlChartType.LINE,
    "scatter": XlChartType.XY_SCATTER_MARKERS,
}

# Reverse map for metadata reads (chart-type int → our kind string).
XL_TO_KIND: dict[int, str] = {int(v): k for k, v in KIND_TO_XL.items()}

CHART_KINDS: tuple[str, ...] = tuple(KIND_TO_XL)

# --- formatting vocab: caller-facing strings → Excel enums (kept narrow) -------

LEGEND_POSITIONS: dict[str, XlLegendPosition] = {
    "right": XlLegendPosition.RIGHT,
    "left": XlLegendPosition.LEFT,
    "top": XlLegendPosition.TOP,
    "bottom": XlLegendPosition.BOTTOM,
    "corner": XlLegendPosition.CORNER,
}

# `value`/`y` is the dependent axis, `category`/`x` the independent one.
AXIS_WHICH: dict[str, XlAxisType] = {
    "value": XlAxisType.VALUE,
    "y": XlAxisType.VALUE,
    "category": XlAxisType.CATEGORY,
    "x": XlAxisType.CATEGORY,
}

SCALE_TYPES: dict[str, XlScaleType] = {
    "linear": XlScaleType.LINEAR,
    "log": XlScaleType.LOGARITHMIC,
    "logarithmic": XlScaleType.LOGARITHMIC,
}

TRENDLINE_KINDS: dict[str, XlTrendlineType] = {
    "linear": XlTrendlineType.LINEAR,
    "exponential": XlTrendlineType.EXPONENTIAL,
    "logarithmic": XlTrendlineType.LOGARITHMIC,
    "log": XlTrendlineType.LOGARITHMIC,
    "moving_average": XlTrendlineType.MOVING_AVERAGE,
    "moving_avg": XlTrendlineType.MOVING_AVERAGE,
    "polynomial": XlTrendlineType.POLYNOMIAL,
    "power": XlTrendlineType.POWER,
}

# Data-point glyph names → `Series.MarkerStyle` (line/scatter only).
MARKER_STYLES: dict[str, XlMarkerStyle] = {
    "none": XlMarkerStyle.NONE,
    "auto": XlMarkerStyle.AUTOMATIC,
    "automatic": XlMarkerStyle.AUTOMATIC,
    "square": XlMarkerStyle.SQUARE,
    "diamond": XlMarkerStyle.DIAMOND,
    "triangle": XlMarkerStyle.TRIANGLE,
    "x": XlMarkerStyle.X,
    "star": XlMarkerStyle.STAR,
    "dot": XlMarkerStyle.DOT,
    "dash": XlMarkerStyle.DASH,
    "circle": XlMarkerStyle.CIRCLE,
    "plus": XlMarkerStyle.PLUS,
}

# Error-bar `include` (which sides) → `Series.ErrorBar(Include=...)`.
ERRORBAR_INCLUDE: dict[str, XlErrorBarInclude] = {
    "both": XlErrorBarInclude.BOTH,
    "plus": XlErrorBarInclude.PLUS_VALUES,
    "minus": XlErrorBarInclude.MINUS_VALUES,
}

# Error-bar `kind` (how the amount is computed) → `Series.ErrorBar(Type=...)`.
ERRORBAR_KINDS: dict[str, XlErrorBarType] = {
    "fixed": XlErrorBarType.FIXED_VALUE,
    "percent": XlErrorBarType.PERCENT,
    "percentage": XlErrorBarType.PERCENT,
    "stdev": XlErrorBarType.STANDARD_DEVIATION,
    "standard_deviation": XlErrorBarType.STANDARD_DEVIATION,
    "sterror": XlErrorBarType.STANDARD_ERROR,
    "standard_error": XlErrorBarType.STANDARD_ERROR,
}

# Error-bar `axis` (which axis the bars run along) → `Series.ErrorBar(Direction=...)`.
ERRORBAR_AXIS: dict[str, XlErrorBarDirection] = {
    "y": XlErrorBarDirection.Y,
    "value": XlErrorBarDirection.Y,
    "x": XlErrorBarDirection.X,
    "category": XlErrorBarDirection.X,
}

# Standard-error error bars compute their own amount; the others use the supplied one.
_ERRORBAR_NEEDS_AMOUNT: frozenset[str] = frozenset({"fixed", "percent", "stdev"})


def _marker_value(marker: Any) -> int:
    """Map a marker name (or raw int) onto an `XlMarkerStyle` int."""
    if isinstance(marker, bool):  # bool is an int subclass — reject it explicitly
        raise OpError(f"marker must be a glyph name or int, got {marker!r}")
    if isinstance(marker, int):
        return int(marker)
    return int(_lookup(marker, MARKER_STYLES, label="marker"))


def _lookup(value: str, table: dict[str, Any], *, label: str) -> Any:
    """Map a caller string onto an enum, raising `OpError` (bad input) on a miss."""
    try:
        key = value.strip().lower()
    except AttributeError as e:
        raise OpError(f"{label} must be a string, got {value!r}") from e
    if key not in table:
        raise OpError(f"unknown {label} {value!r}; expected one of {sorted(table)}")
    return table[key]


def probe_excel_available() -> bool:
    """True if the `Excel.Application` COM server is registered (non-invasive).

    Looks up the ProgID under ``HKEY_CLASSES_ROOT`` — present iff Excel is
    installed — *without* launching Excel or touching any running instance (so a
    user's open Excel is never disturbed, and no hidden Excel is spawned just to
    check). `insert_chart` calls this before inserting and raises
    `ExcelNotAvailableError` on False, leaving the document untouched.
    """
    try:
        import winreg  # Windows-only; absent off-Windows where charts can't run anyway
    except ImportError:
        return False
    try:
        # typeshed guards winreg's members behind sys.platform == "win32", so
        # OpenKey/HKEY_CLASSES_ROOT read as missing under Linux-CI mypy.
        winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, "Excel.Application").Close()  # type: ignore[attr-defined]
        return True
    except OSError:
        return False


def _coerce_float(value: Any, *, ctx: str) -> float:
    if isinstance(value, bool):  # bool is an int subclass — reject it explicitly
        raise OpError(f"chart {ctx} must be a number, got {value!r}")
    try:
        return float(value)
    except (TypeError, ValueError) as e:
        raise OpError(f"chart {ctx} must be a number, got {value!r}") from e


def normalize_chart_data(kind: str, data: Any) -> tuple[list[Any], list[float]]:
    """Validate `data` for `kind` and split it into (x_values, y_values).

    Two shapes are accepted (JSON tells them apart by type):

    - a mapping ``{label: value}`` — categorical, for bar / pie / line;
    - a sequence of ``[x, y]`` pairs — numeric x/y, for scatter (and line, which
      also accepts the mapping form).

    `y` values are always numeric. For ``scatter`` the `x` values must be numeric
    too — so the plot uses a real value axis, with duplicate/clustered x preserved
    as distinct points, instead of category positions — while for the categorical
    kinds the `x` values are labels, kept as given. Raises `OpError` (bad input →
    exit 1) for an empty/unknown shape or a non-numeric value.
    """
    if isinstance(data, Mapping):
        items: list[tuple[Any, Any]] = list(data.items())
    elif isinstance(data, (list, tuple)):
        items = []
        for i, pair in enumerate(data):
            if (
                isinstance(pair, Mapping)
                or isinstance(pair, (str, bytes))
                or not isinstance(pair, (list, tuple))
                or len(pair) != 2
            ):
                raise OpError(f"chart data array must hold [x, y] pairs; item {i} is {pair!r}")
            items.append((pair[0], pair[1]))
    else:
        raise OpError(
            "chart data must be an object {label: value} or an array of [x, y] pairs; "
            f"got {type(data).__name__}"
        )
    if not items:
        raise OpError("chart data is empty")
    numeric_x = kind == "scatter"
    xs: list[Any] = []
    ys: list[float] = []
    for i, (x, y) in enumerate(items):
        ys.append(_coerce_float(y, ctx=f"value (point {i})"))
        xs.append(_coerce_float(x, ctx=f"scatter x (point {i})") if numeric_x else x)
    return xs, ys


def chart_shapes(doc_com: Any) -> list[Any]:
    """The document's chart inline shapes (those with `HasChart`), document order."""
    out: list[Any] = []
    shapes = doc_com.InlineShapes
    for i in range(1, int(shapes.Count) + 1):
        shape = shapes.Item(i)
        try:
            has_chart = bool(shape.HasChart)
        except Exception:
            has_chart = False
        if has_chart:
            out.append(shape)
    return out


def _close_workbook(wb: Any, xlapp: Any) -> None:
    """Best-effort: close the embedded workbook, quit its Excel if now empty.

    Paired with `ChartData.BreakLink()` (run before this) so the workbook can
    actually close and the hidden Excel terminate — leaving no orphan process.
    """
    if wb is not None:
        try:
            wb.Close(False)
        except Exception:
            pass
    if xlapp is not None:
        try:
            if int(xlapp.Workbooks.Count) == 0:
                xlapp.Quit()
        except Exception:
            pass


def populate_chart(
    chart_com: Any, kind: str, xs: list[Any], ys: list[float], title: str | None
) -> None:
    """Write data into the chart's embedded workbook and finalise it as static.

    Writes x → column A and y → column B (with a header row so the series name
    comes from B1), points the chart's single series at those cells via a
    `=SERIES(...)` formula, sets the title, then `BreakLink`s and closes the
    embedded workbook so no orphan Excel is left running. Always closes the
    workbook — even on failure — and re-raises; the caller removes the
    half-built chart.
    """
    wb = None
    xlapp = None
    try:
        wb = chart_com.ChartData.Workbook
        try:
            xlapp = wb.Application
        except Exception:
            xlapp = None
        ws = wb.Worksheets(1)
        ws.UsedRange.Clear()
        ws.Cells(1, 1).Value = "x"
        ws.Cells(1, 2).Value = title or "series"
        n = len(xs)
        for i, (x, y) in enumerate(zip(xs, ys, strict=True), start=2):
            ws.Cells(i, 1).Value = x
            ws.Cells(i, 2).Value = y
        series = chart_com.SeriesCollection()
        while int(series.Count) > 1:
            series(int(series.Count)).Delete()
        if int(series.Count) == 0:
            series.NewSeries()
        sheet = ws.Name
        name_ref = f"{sheet}!{ws.Cells(1, 2).Address}"
        x_ref = f"{sheet}!{ws.Range(ws.Cells(2, 1), ws.Cells(n + 1, 1)).Address}"
        y_ref = f"{sheet}!{ws.Range(ws.Cells(2, 2), ws.Cells(n + 1, 2)).Address}"
        series(1).Formula = f"=SERIES({name_ref},{x_ref},{y_ref},1)"
        if title:
            chart_com.HasTitle = True
            chart_com.ChartTitle.Text = title
        else:
            chart_com.HasTitle = False
        # Sever the link so the chart's data is static and Word releases the
        # embedded Excel — required so the hidden Excel can actually terminate.
        chart_com.ChartData.BreakLink()
    finally:
        _close_workbook(wb, xlapp)


# --- post-insert formatting (operates on the BreakLink-static Chart) -----------
#
# Every helper below mutates a chart's *presentation* (style, legend, axes,
# trendlines, colours) on the static, post-insert `Chart` COM object. Live-probed
# 2026-06-16/17: none of these reopen the embedded-Excel data grid or respin a
# hidden Excel (verified 0 orphan EXCEL.EXE across the full surface), so they need
# no Excel gate. The standing rule still holds: they never read *series data*
# back. Bad input raises `OpError`; COM failures surface via the caller's
# `translate_com_errors`.


def apply_chart_format(
    chart_com: Any,
    *,
    title: Any = _UNSET,
    legend: bool | None = None,
    legend_position: str | None = None,
    chart_style: int | None = None,
    background: Any = None,
    plot_background: Any = None,
    font: str | None = None,
    font_size: Any = None,
    font_color: Any = None,
    data_labels: bool | None = None,
    data_label_format: str | None = None,
    chart_type: str | None = None,
    gap_width: int | None = None,
    overlap: int | None = None,
    data_table: bool | None = None,
) -> None:
    """Apply whole-chart / design formatting. Tri-state: only passed fields write.

    `title=None` clears the title (omit it to leave it alone — that's what the
    `_UNSET` sentinel distinguishes). `legend_position` implies the legend is
    shown. `chart_style` is the built-in design-gallery int. `background` /
    `plot_background` and `font_color` accept any [`to_bgr`][wordlive._format.to_bgr]
    colour; `font` / `font_size` set the whole-chart font. `data_labels` toggles
    point labels on every series; `chart_type` re-types the chart in place.
    `gap_width` / `overlap` tune bar spacing (bar/column charts only); `data_table`
    toggles the data-table grid beneath the plot.
    """
    if title is not _UNSET:
        if title:
            chart_com.HasTitle = True
            chart_com.ChartTitle.Text = str(title)
        else:  # None or "" → remove the title
            chart_com.HasTitle = False
    if legend is not None:
        chart_com.HasLegend = bool(legend)
    if legend_position is not None:
        chart_com.HasLegend = True
        chart_com.Legend.Position = int(
            _lookup(legend_position, LEGEND_POSITIONS, label="legend_position")
        )
    if chart_style is not None:
        chart_com.ChartStyle = int(chart_style)
    if background is not None:
        chart_com.ChartArea.Format.Fill.ForeColor.RGB = to_bgr(background)
    if plot_background is not None:
        chart_com.PlotArea.Format.Fill.ForeColor.RGB = to_bgr(plot_background)
    if font is not None or font_size is not None or font_color is not None:
        from ._anchors import _apply_font  # lazy: _anchors imports this module

        _apply_font(chart_com.ChartArea.Font, font_name=font, size=font_size, color=font_color)
    if data_labels is not None:
        sc = chart_com.SeriesCollection()
        for i in range(1, int(sc.Count) + 1):
            sc(i).HasDataLabels = bool(data_labels)
    if data_label_format is not None:
        sc = chart_com.SeriesCollection()
        for i in range(1, int(sc.Count) + 1):
            series = sc(i)
            series.HasDataLabels = True
            series.DataLabels().NumberFormat = str(data_label_format)
    if chart_type is not None:
        if chart_type not in KIND_TO_XL:
            raise OpError(f"unknown chart_type {chart_type!r}; expected one of {list(CHART_KINDS)}")
        chart_com.ChartType = int(KIND_TO_XL[chart_type])
    if gap_width is not None or overlap is not None:
        try:
            grp = chart_com.ChartGroups(1)
            if gap_width is not None:
                grp.GapWidth = int(gap_width)
            if overlap is not None:
                grp.Overlap = int(overlap)
        except Exception as e:
            raise OpError("gap_width / overlap apply only to bar/column charts") from e
    if data_table is not None:
        chart_com.HasDataTable = bool(data_table)


def apply_axis_format(
    chart_com: Any,
    which: str,
    *,
    title: Any = _UNSET,
    minimum: Any = None,
    maximum: Any = None,
    scale: str | None = None,
    number_format: str | None = None,
    gridlines: bool | None = None,
) -> None:
    """Format one axis of the chart. `which` is ``"value"``/``"y"`` or
    ``"category"``/``"x"``; tri-state otherwise. `scale` is ``"linear"`` or
    ``"log"``; `title=None` clears the axis title."""
    atype = _lookup(which, AXIS_WHICH, label="axis")
    axis = chart_com.Axes(int(atype), int(XlAxisGroup.PRIMARY))
    if title is not _UNSET:
        if title:
            axis.HasTitle = True
            axis.AxisTitle.Text = str(title)
        else:
            axis.HasTitle = False
    if minimum is not None:
        axis.MinimumScale = float(minimum)
    if maximum is not None:
        axis.MaximumScale = float(maximum)
    if scale is not None:
        axis.ScaleType = int(_lookup(scale, SCALE_TYPES, label="scale"))
    if number_format is not None:
        axis.TickLabels.NumberFormat = str(number_format)
    if gridlines is not None:
        axis.HasMajorGridlines = bool(gridlines)


def add_trendline(
    chart_com: Any,
    *,
    series: int = 1,
    kind: str = "linear",
    display_equation: bool = False,
    display_r_squared: bool = False,
    forward: Any = None,
    backward: Any = None,
    order: int | None = None,
    period: int | None = None,
) -> None:
    """Fit a trendline to a series. `kind` is one of `TRENDLINE_KINDS`; the
    optional `forward`/`backward` extend the fit that many units past the data.
    `order` is the polynomial degree (2–6, `kind="polynomial"`); `period` is the
    moving-average window (`kind="moving_average"`)."""
    xl = _lookup(kind, TRENDLINE_KINDS, label="trendline kind")
    sc = chart_com.SeriesCollection(int(series))
    tl = sc.Trendlines().Add(int(xl))  # positional Type — keywords drop under late binding
    if order is not None:
        tl.Order = int(order)
    if period is not None:
        tl.Period = int(period)
    if display_equation:
        tl.DisplayEquation = True
    if display_r_squared:
        tl.DisplayRSquared = True
    if forward is not None:
        tl.Forward = float(forward)
    if backward is not None:
        tl.Backward = float(backward)


def format_series(
    chart_com: Any,
    *,
    series: int = 1,
    point: int | None = None,
    marker: Any = None,
    marker_size: int | None = None,
    smooth: bool | None = None,
    explosion: int | None = None,
    data_labels: bool | None = None,
    data_label_size: Any = None,
    data_label_color: Any = None,
) -> None:
    """Format one series (or a single 1-based `point` within it). Tri-state.

    `marker` is a glyph name (`MARKER_STYLES`) or raw `XlMarkerStyle` int, and
    `marker_size` (2–72) sizes them — both for line/scatter series. `smooth`
    curves a line/scatter through its points. `explosion` (0–400) pulls a pie
    slice out. `data_labels` toggles this series' point labels; `data_label_size`
    / `data_label_color` style their font (a colour goes through
    [`to_bgr`][wordlive._format.to_bgr]). When `point` is given, `marker` /
    `explosion` / the data-label font target that single point; `marker_size` /
    `smooth` / the `data_labels` toggle stay series-wide.
    """
    sc = chart_com.SeriesCollection(int(series))
    pt = sc.Points(int(point)) if point is not None else None
    if marker is not None:
        (pt if pt is not None else sc).MarkerStyle = _marker_value(marker)
    if marker_size is not None:
        sc.MarkerSize = int(marker_size)
    if smooth is not None:
        sc.Smooth = bool(smooth)
    if explosion is not None:
        (pt if pt is not None else sc).Explosion = int(explosion)
    if data_labels is not None:
        sc.HasDataLabels = bool(data_labels)
    if data_label_size is not None or data_label_color is not None:
        sc.HasDataLabels = True  # DataLabels() raises until labels are shown
        font = pt.DataLabel.Font if pt is not None else sc.DataLabels().Font
        if data_label_size is not None:
            font.Size = float(data_label_size)
        if data_label_color is not None:
            font.Color = to_bgr(data_label_color)


def add_error_bars(
    chart_com: Any,
    *,
    series: int = 1,
    kind: str = "fixed",
    amount: Any = None,
    include: str = "both",
    axis: str = "y",
) -> None:
    """Draw error bars on a series. `kind` is ``"fixed"`` / ``"percent"`` /
    ``"stdev"`` / ``"sterror"``; `amount` is the magnitude (required for all but
    ``"sterror"``, which Word computes from the data). `include` is which side(s)
    to draw (``"both"`` / ``"plus"`` / ``"minus"``); `axis` is ``"y"``/``"value"``
    (the usual) or ``"x"``/``"category"`` for scatter x-uncertainty."""
    etype = _lookup(kind, ERRORBAR_KINDS, label="error-bar kind")
    incl = _lookup(include, ERRORBAR_INCLUDE, label="error-bar include")
    axdir = _lookup(axis, ERRORBAR_AXIS, label="error-bar axis")
    needs_amount = kind.strip().lower() in _ERRORBAR_NEEDS_AMOUNT
    if amount is None:
        if needs_amount:
            raise OpError(f"error-bar kind {kind!r} needs an amount")
        amt = 1.0  # ignored by Word for standard-error bars
    else:
        amt = _coerce_float(amount, ctx="error-bar amount")
    sc = chart_com.SeriesCollection(int(series))
    sc.HasErrorBars = True
    # ErrorBar(Direction, Include, Type, Amount) — positional under late binding.
    sc.ErrorBar(int(axdir), int(incl), int(etype), amt)


def set_series_color(
    chart_com: Any, color: Any, *, series: int = 1, point: int | None = None
) -> None:
    """Recolour a whole series, or a single 1-based `point` (a bar / pie slice /
    marker). Sets the fill, and the line/marker colour too where the series has
    one (line/scatter). `color` is any [`to_bgr`][wordlive._format.to_bgr] colour."""
    bgr = to_bgr(color)
    sc = chart_com.SeriesCollection(int(series))
    target = sc.Points(int(point)) if point is not None else sc
    target.Format.Fill.ForeColor.RGB = bgr
    try:
        target.Format.Line.ForeColor.RGB = bgr
    except Exception:
        pass  # bar/pie points have no separate line — the fill is the colour
