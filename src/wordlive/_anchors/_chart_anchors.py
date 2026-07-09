"""`chart:N` anchors and the chart collection (see also the `_charts` feature module)."""

from __future__ import annotations

from collections.abc import Iterator
from typing import TYPE_CHECKING, Any

from .. import _charts, _com
from ..exceptions import AnchorNotFoundError, OpError

if TYPE_CHECKING:
    from .._document import Document

from ._base import Anchor


class ChartAnchor(Anchor):
    """An Excel-backed chart located by 1-based index — `chart:N`.

    Indexes the document's chart inline shapes in document order. The anchor
    resolves to the chart's range, so it inherits `apply_style` / formatting like
    any anchor. `chart_type` reports the kind string (``"bar"`` / ``"pie"`` /
    ``"line"`` / ``"scatter"``) and `title` the chart title — metadata only:
    charts are inserted with a broken data link, so the underlying series data is
    static and isn't read back. Create charts with
    [`Anchor.insert_chart`][wordlive.Anchor.insert_chart]; discover them via
    [`doc.charts`][wordlive.Document.charts]. A chart isn't plain text, so
    `set_text` raises.
    """

    kind = "chart"

    def __init__(self, doc: Document, index: int) -> None:
        super().__init__(doc, name=f"chart:{index}")
        self._index = index

    @property
    def index(self) -> int:
        return self._index

    @property
    def anchor_id(self) -> str:
        return f"chart:{self._index}"

    def _shape(self) -> Any:
        shapes = _charts.chart_shapes(self._doc.com)
        if not (1 <= self._index <= len(shapes)):
            raise AnchorNotFoundError("chart", f"chart:{self._index}")
        return shapes[self._index - 1]

    def _range(self) -> Any:
        return self._shape().Range

    @property
    def chart_type(self) -> str:
        """The chart kind — ``"bar"`` / ``"pie"`` / ``"line"`` / ``"scatter"`` (or the raw int)."""
        with _com.translate_com_errors():
            ctype = int(self._shape().Chart.ChartType)
        return _charts.XL_TO_KIND.get(ctype, str(ctype))

    @property
    def title(self) -> str | None:
        """The chart's title, or ``None`` if it has none."""
        with _com.translate_com_errors():
            chart = self._shape().Chart
            if not bool(chart.HasTitle):
                return None
            return str(chart.ChartTitle.Text or "")

    @property
    def chart_style(self) -> int:
        """The built-in design-gallery style id (`Chart.ChartStyle`)."""
        with _com.translate_com_errors():
            return int(self._shape().Chart.ChartStyle)

    @property
    def has_legend(self) -> bool:
        """Whether the chart currently shows a legend."""
        with _com.translate_com_errors():
            return bool(self._shape().Chart.HasLegend)

    def format(
        self,
        *,
        title: Any = _charts._UNSET,
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
    ) -> ChartAnchor:
        """Apply whole-chart / design formatting — Word's chart "Design" tab.

        All kwargs are optional and tri-state; only the ones you pass are written.
        `title=None` clears the chart title (omit it to leave it). `legend`
        toggles the legend; `legend_position` (`"right"`/`"left"`/`"top"`/
        `"bottom"`/`"corner"`) implies it's shown. `chart_style` is the built-in
        design-gallery int. `background` / `plot_background` fill the chart and
        plot areas; `font` / `font_size` / `font_color` set the whole-chart font.
        `data_labels` toggles point labels on every series, `data_label_format`
        is their number format. `chart_type` (`"bar"`/`"pie"`/`"line"`/
        `"scatter"`) re-types the chart in place. `gap_width` / `overlap` tune bar
        spacing (bar/column charts only); `data_table` toggles the data-table grid
        beneath the plot. Operates on the static chart — no Excel needed. Returns
        `self` (chainable); wrap in `doc.edit(...)` for atomic undo. Bad input
        raises `OpError`.
        """
        self._apply(
            _charts.apply_chart_format,
            title=title,
            legend=legend,
            legend_position=legend_position,
            chart_style=chart_style,
            background=background,
            plot_background=plot_background,
            font=font,
            font_size=font_size,
            font_color=font_color,
            data_labels=data_labels,
            data_label_format=data_label_format,
            chart_type=chart_type,
            gap_width=gap_width,
            overlap=overlap,
            data_table=data_table,
        )
        return self

    def set_axis(
        self,
        which: str,
        *,
        title: Any = _charts._UNSET,
        minimum: Any = None,
        maximum: Any = None,
        scale: str | None = None,
        number_format: str | None = None,
        gridlines: bool | None = None,
    ) -> ChartAnchor:
        """Format one axis. `which` is ``"value"``/``"y"`` or ``"category"``/``"x"``.

        Tri-state: `title=None` clears the axis title; `minimum`/`maximum` set the
        scale bounds; `scale` is ``"linear"`` or ``"log"`` (log is ideal for
        order-of-magnitude data); `number_format` is the tick-label format string;
        `gridlines` toggles major gridlines. Returns `self`. Bad input raises
        `OpError`.
        """
        self._apply(
            _charts.apply_axis_format,
            which,
            title=title,
            minimum=minimum,
            maximum=maximum,
            scale=scale,
            number_format=number_format,
            gridlines=gridlines,
        )
        return self

    def add_trendline(
        self,
        *,
        series: int = 1,
        kind: str = "linear",
        display_equation: bool = False,
        display_r_squared: bool = False,
        forward: Any = None,
        backward: Any = None,
        order: int | None = None,
        period: int | None = None,
    ) -> ChartAnchor:
        """Fit a trendline to a series (1-based `series`).

        `kind` is ``"linear"``, ``"exponential"``, ``"logarithmic"``,
        ``"moving_average"``, ``"polynomial"``, or ``"power"``. `display_equation`
        / `display_r_squared` annotate the fit — a power/exponential fit with the
        equation literally draws the law of best fit. `forward` / `backward`
        extend the line that many units past the data. `order` is the polynomial
        degree (2–6, with `kind="polynomial"`); `period` is the moving-average
        window (with `kind="moving_average"`). Returns `self`. Bad input raises
        `OpError`.
        """
        self._apply(
            _charts.add_trendline,
            series=series,
            kind=kind,
            display_equation=display_equation,
            display_r_squared=display_r_squared,
            forward=forward,
            backward=backward,
            order=order,
            period=period,
        )
        return self

    def set_series_color(
        self, color: Any, *, series: int = 1, point: int | None = None
    ) -> ChartAnchor:
        """Recolour a whole series, or a single 1-based `point` (bar / pie slice).

        `color` is a named colour, hex (`"#2E86C1"`), or `(r, g, b)`. Omit `point`
        to colour the entire series; pass it to vary one bar / slice / marker.
        Sets the line/marker colour too where the series has one (line/scatter).
        Returns `self`. Bad input raises `OpError`.
        """
        self._apply(_charts.set_series_color, color, series=series, point=point)
        return self

    def format_series(
        self,
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
    ) -> ChartAnchor:
        """Format one series, or a single 1-based `point` within it.

        `marker` is a glyph name (``"circle"``/``"square"``/``"diamond"``/
        ``"triangle"``/``"x"``/``"star"``/``"dot"``/``"dash"``/``"plus"``/
        ``"none"``/``"auto"``) or a raw `XlMarkerStyle` int, with `marker_size`
        (2–72) — both for line/scatter. `smooth` curves a line/scatter through its
        points. `explosion` (0–400) pulls a pie slice out. `data_labels` toggles
        this series' point labels; `data_label_size` / `data_label_color` style
        their font. With `point` set, `marker` / `explosion` / the data-label font
        target that one point; `marker_size` / `smooth` / the `data_labels` toggle
        stay series-wide. Returns `self`. Bad input raises `OpError`.
        """
        self._apply(
            _charts.format_series,
            series=series,
            point=point,
            marker=marker,
            marker_size=marker_size,
            smooth=smooth,
            explosion=explosion,
            data_labels=data_labels,
            data_label_size=data_label_size,
            data_label_color=data_label_color,
        )
        return self

    def add_error_bars(
        self,
        *,
        series: int = 1,
        kind: str = "fixed",
        amount: Any = None,
        include: str = "both",
        axis: str = "y",
    ) -> ChartAnchor:
        """Draw error bars on a series (1-based `series`).

        `kind` is ``"fixed"`` (an absolute amount), ``"percent"`` (of each value),
        ``"stdev"`` (multiples of the standard deviation), or ``"sterror"`` (the
        standard error — Word computes it, so `amount` is ignored). `amount` is the
        magnitude (required for all kinds but ``"sterror"``). `include` is which
        side(s) to draw (``"both"`` / ``"plus"`` / ``"minus"``); `axis` is
        ``"y"``/``"value"`` (the usual) or ``"x"``/``"category"`` for scatter
        x-uncertainty. Returns `self`. Bad input raises `OpError`.
        """
        self._apply(
            _charts.add_error_bars,
            series=series,
            kind=kind,
            amount=amount,
            include=include,
            axis=axis,
        )
        return self

    def _apply(self, fn: Any, *args: Any, **kwargs: Any) -> None:
        """Run a `_charts` formatting helper on this chart's `Chart`, translating
        COM and bad-input errors into the wordlive hierarchy (`OpError`)."""
        chart = self._shape().Chart
        try:
            with _com.translate_com_errors():
                fn(chart, *args, **kwargs)
        except (ValueError, TypeError) as e:
            raise OpError(str(e)) from e

    def set_text(self, text: str) -> None:
        raise OpError(
            "a chart anchor has no plain text to set; delete it and "
            "insert_chart(...) again to change it"
        )


class ChartCollection:
    """Read-only, iterable view over the document's charts (`doc.charts`).

    Index a chart by 1-based position (`doc.charts[2]`) to get a
    [`ChartAnchor`][wordlive.ChartAnchor] (`chart:N`); `list()` summarises every
    chart — id, kind, title, and the `para:N` it sits in. Positions follow
    document order. Metadata only — charts are inserted with their data link
    broken (static data), so reading the series back is deferred. The write
    mirror is any anchor's [`insert_chart`][wordlive.Anchor.insert_chart].
    """

    def __init__(self, doc: Document) -> None:
        self._doc = doc

    def __len__(self) -> int:
        with _com.translate_com_errors():
            return len(_charts.chart_shapes(self._doc.com))

    def __getitem__(self, index: int) -> ChartAnchor:
        if isinstance(index, bool) or not isinstance(index, int):
            raise TypeError(f"chart index must be int, got {type(index).__name__}")
        n = len(self)
        if not (1 <= index <= n):
            raise AnchorNotFoundError("chart", str(index))
        return ChartAnchor(self._doc, index)

    def __iter__(self) -> Iterator[ChartAnchor]:
        for i in range(1, len(self) + 1):
            yield ChartAnchor(self._doc, i)

    def list(self) -> list[dict[str, Any]]:
        """Every chart as `{index, anchor_id, kind, title, chart_style, has_legend, para}`.

        `kind` is the chart-type string; `title` the chart title (or ``None``);
        `chart_style` the design-gallery id; `has_legend` whether a legend shows;
        `para` the `para:N` the chart sits in. Touches only `ChartType` /
        `ChartTitle` / `ChartStyle` / `HasLegend` (never the series data), so it's
        cheap and Word-stable.
        """
        out: list[dict[str, Any]] = []
        with _com.translate_com_errors():
            shapes = _charts.chart_shapes(self._doc.com)
            for i, shape in enumerate(shapes, start=1):
                chart = shape.Chart
                try:
                    kind: str | None = _charts.XL_TO_KIND.get(
                        int(chart.ChartType), str(int(chart.ChartType))
                    )
                except Exception:
                    kind = None
                try:
                    title = str(chart.ChartTitle.Text) if bool(chart.HasTitle) else None
                except Exception:
                    title = None
                try:
                    chart_style: int | None = int(chart.ChartStyle)
                except Exception:
                    chart_style = None
                try:
                    has_legend: bool | None = bool(chart.HasLegend)
                except Exception:
                    has_legend = None
                try:
                    start = int(shape.Range.Start)
                except Exception:
                    start = None
                para_id: str | None = None
                if start is not None:
                    para = self._doc.paragraphs.at(start)
                    para_id = para.anchor_id if para is not None else None
                out.append(
                    {
                        "index": i,
                        "anchor_id": f"chart:{i}",
                        "kind": kind,
                        "title": title,
                        "chart_style": chart_style,
                        "has_legend": has_legend,
                        "para": para_id,
                    }
                )
        return out
