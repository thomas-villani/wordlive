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

from .constants import XlChartType
from .exceptions import OpError

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
        winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, "Excel.Application").Close()
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
