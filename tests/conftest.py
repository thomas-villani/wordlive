"""Test fixtures.

`fake_word` builds a MagicMock that quacks like a Word.Application COM object —
enough surface to exercise the politeness logic, anchors, and CLI shape without
needing a real Word install.

`real_word` is for the @pytest.mark.smoke suite and is skipped if Word can't
be reached.
"""

from __future__ import annotations

import re
from collections.abc import Iterable
from types import SimpleNamespace
from typing import Any
from unittest.mock import MagicMock

import pytest

# ---------------------------------------------------------------------------
# Fake COM Application
# ---------------------------------------------------------------------------


class _FakeBookmarkRegistry:
    """Mimics ActiveDocument.Bookmarks: Exists, by-name lookup, Add, iteration.

    Bookmark mocks (and their `.Range`) are cached per name so a test that
    writes via `Bookmarks(name).Range.X = y` can read `X` back through the
    same handle. `Add` and the seeding helper invalidate the cache so the
    range offsets stay consistent with the registered tuple.
    """

    def __init__(self) -> None:
        self._items: dict[str, tuple[int, int]] = {}
        self._cache: dict[str, Any] = {}

    def add(self, name: str, start: int, end: int) -> None:
        self._items[name] = (start, end)
        self._cache.pop(name, None)

    # COM-style methods
    def Exists(self, name: str) -> bool:
        return name in self._items

    def Add(self, Name: str | None = None, Range: Any | None = None) -> None:
        if Name is None or Range is None:
            return
        self._items[Name] = (int(Range.Start), int(Range.End))
        self._cache.pop(Name, None)

    def __call__(self, name: str) -> Any:
        if name in self._cache:
            return self._cache[name]
        s, e = self._items[name]
        bm = MagicMock(name=f"Bookmark[{name}]")
        bm.Name = name
        bm.Range = _make_range(s, e)
        self._cache[name] = bm
        return bm

    def __iter__(self) -> Iterable[Any]:
        return iter([self(name) for name in self._items])


class _FakeDropdownEntry:
    def __init__(self, text: str, value: str) -> None:
        self.Text = text
        self.Value = value


class _FakeDropdownListEntries:
    """Mimics ContentControl.DropdownListEntries: Add/Clear/Count/Item/iterate.

    `Add` is a recording MagicMock (side-effect appends a real entry) so the
    positional (Text, Value) call args stay assertable while `set_items`
    (Clear + re-add) still round-trips against the readable entry list.
    """

    def __init__(self, items: list[tuple[str, str]] | None = None) -> None:
        self._entries = [_FakeDropdownEntry(t, v) for t, v in (items or [])]
        self.Add = MagicMock(name="DropdownListEntriesAdd", side_effect=self._add)

    def _add(self, Text: str, Value: str | None = None) -> _FakeDropdownEntry:  # positional
        entry = _FakeDropdownEntry(Text, Value if Value is not None else Text)
        self._entries.append(entry)
        return entry

    def Clear(self) -> None:
        self._entries.clear()

    @property
    def Count(self) -> int:
        return len(self._entries)

    def Item(self, index: int) -> _FakeDropdownEntry:  # 1-based, like Word
        return self._entries[index - 1]

    def __iter__(self) -> Iterable[Any]:
        return iter(self._entries)


class _FakeContentControls:
    def __init__(self, controls: list[dict[str, Any]]) -> None:
        self._items = []
        for cc in controls:
            mock = MagicMock(name=f"CC[{cc.get('title', '?')}]")
            mock.Title = cc.get("title", "")
            mock.Tag = cc.get("tag", "")
            mock.Type = cc.get("kind", 0)  # WdContentControlType int (0 = rich_text)
            mock.LockContents = cc.get("lock_contents", False)
            mock.LockContentControl = cc.get("lock_control", False)
            mock.Range = _make_range(cc.get("start", 0), cc.get("end", 0))
            mock.Range.Text = cc.get("text", "")
            raw_items = cc.get("items") or []
            pairs = [
                (i, i) if isinstance(i, str) else (i["text"], i.get("value", i["text"]))
                for i in raw_items
            ]
            mock.DropdownListEntries = _FakeDropdownListEntries(pairs)
            self._items.append(mock)
        # Add(Type, Range) is a recording MagicMock so tests can assert the
        # control type; it appends a settable control (Title/Tag default to "" so
        # an unnamed new control doesn't shadow an existing one in `_cc_by_name`).
        self.Add = MagicMock(name="CCAdd", side_effect=self._add)

    def _add(self, Type: int = 0, Range: Any = None) -> Any:
        mock = MagicMock(name=f"CC[new:{Type}]")
        mock.Type = Type
        mock.Title = ""
        mock.Tag = ""
        mock.LockContents = False
        mock.LockContentControl = False
        mock.DropdownListEntries = _FakeDropdownListEntries()
        rng = _make_range(0, 0)
        rng.Text = ""
        mock.Range = rng
        self._items.append(mock)
        return mock

    def __iter__(self) -> Iterable[Any]:
        return iter(self._items)


class _FakeParagraphs:
    def __init__(self, paragraphs: list[dict[str, Any]]) -> None:
        self._items = []
        for p in paragraphs:
            mock = MagicMock(name=f"Para[{p.get('text', '?')[:20]}]")
            mock.OutlineLevel = p.get("level", 10)
            mock.Range = _make_range(p.get("start", 0), p.get("end", 0))
            mock.Range.Text = p.get("text", "") + "\r"
            # Applied paragraph style name, surfaced by ParagraphCollection.list()
            # (via Style) and format_info() (via ParagraphStyle); keep both in sync.
            style_name = p.get("style", "Normal")
            mock.Range.Style.NameLocal = style_name
            mock.Range.ParagraphStyle.NameLocal = style_name
            self._items.append(mock)

    def __iter__(self) -> Iterable[Any]:
        return iter(self._items)


class _FakeListFormat:
    """Mimics Range.ListFormat: apply/remove templates, in/out-dent, read state.

    `ApplyListTemplate` reads the gallery off the template handed in by
    `_list_application` (1=bullet, 2=number, 3=outline) and updates the
    readable list state to match, so apply -> list_info round-trips. A template
    is `None` until a list is applied, so `restart_numbering` on a plain range
    raises (matching Word).
    """

    _GALLERY_TO_TYPE = {1: 2, 2: 3, 3: 4}  # bullet->BULLET, number->SIMPLE, outline->OUTLINE
    _MARKER = {2: "•", 3: "1.", 4: "1."}

    def __init__(self) -> None:
        self.ListType = 0
        self.ListLevelNumber = 1
        self.ListValue = 0
        self.ListString = ""
        self._gallery: int | None = None
        self._continue = False
        self._template: Any = None  # the applied template (gallery mock or custom)

    def ApplyListTemplate(
        self, ListTemplate=None, ContinuePreviousList=False, ApplyTo=0, DefaultListBehavior=2, **kw
    ):
        self._template = ListTemplate
        self._continue = bool(ContinuePreviousList)
        self.ListLevelNumber = 1
        gallery = getattr(ListTemplate, "_gallery", None)
        if gallery is not None:
            self._gallery = int(gallery)
            self.ListType = self._GALLERY_TO_TYPE.get(self._gallery, 3)
            self.ListValue = 1 if self.ListType in (3, 4) else 0
            self.ListString = self._MARKER.get(self.ListType, "")
        else:
            # A custom template (apply_list_format): infer kind from level 1's
            # format — a bare glyph (no %N) is a bullet, else numbered.
            self._gallery = None
            lvl1 = ListTemplate.ListLevels(1)
            fmt = str(getattr(lvl1, "NumberFormat", "") or "")
            is_bullet = bool(fmt) and "%" not in fmt
            self.ListType = 2 if is_bullet else 3
            self.ListValue = 0 if is_bullet else 1
            self.ListString = fmt if is_bullet else fmt.replace("%1", "1")

    def RemoveNumbers(self, NumberType=3):
        self.ListType = 0
        self.ListLevelNumber = 1
        self.ListValue = 0
        self.ListString = ""
        self._gallery = None
        self._template = None

    def ListIndent(self) -> None:
        self.ListLevelNumber = min(9, self.ListLevelNumber + 1)

    def ListOutdent(self) -> None:
        self.ListLevelNumber = max(1, self.ListLevelNumber - 1)

    @property
    def ListTemplate(self):
        if self._template is not None:
            return self._template
        if self._gallery is None:
            return None
        t = MagicMock(name="ListTemplate")
        t._gallery = self._gallery
        return t


class _FakeListLevel:
    """Mimics a `ListLevel` — the settable per-level format props author/read use."""

    def __init__(self) -> None:
        self.NumberFormat = ""
        self.NumberStyle = 0
        self.TrailingCharacter = 0
        self.StartAt = 1
        self.NumberPosition = 0.0
        self.TextPosition = 0.0
        self.Alignment = 0
        self.Font = MagicMock(name="LevelFont")
        self.Font.Name = ""
        self.Font.Bold = False
        self.Font.Italic = False
        self.Font.Color = 0


class _FakeListLevels:
    """Callable `ListLevels` view: `Count` and `ListLevels(i)` -> a `_FakeListLevel`."""

    def __init__(self, count: int) -> None:
        self._levels = {i: _FakeListLevel() for i in range(1, count + 1)}

    @property
    def Count(self) -> int:
        return len(self._levels)

    def __call__(self, index: int) -> Any:
        return self._levels[index]


class _FakeListTemplate:
    """A custom `ListTemplate` minted by `Document.ListTemplates.Add(outline)`."""

    def __init__(self, outline: bool) -> None:
        # OutlineNumbered=True → 9 levels, else a single level (matches Word).
        self.ListLevels = _FakeListLevels(9 if outline else 1)
        self.Name = ""


class _FakeListTemplates:
    """Mimics `Document.ListTemplates`: `Count` and `Add(OutlineNumbered, [Name])`."""

    def __init__(self) -> None:
        self._items: list[Any] = []

    @property
    def Count(self) -> int:
        return len(self._items)

    def Add(self, OutlineNumbered=False, Name=None):
        lt = _FakeListTemplate(bool(OutlineNumbered))
        self._items.append(lt)
        return lt


def _list_application() -> MagicMock:
    """A stand-in Application whose ListGalleries(n).ListTemplates(1) carries gallery n."""
    app = MagicMock(name="ListApplication")

    def galleries(gallery_type):
        gallery = MagicMock(name=f"Gallery[{gallery_type}]")

        def templates(n):
            t = MagicMock(name=f"Template[{gallery_type}:{n}]")
            t._gallery = int(gallery_type)
            return t

        gallery.ListTemplates.side_effect = templates
        return gallery

    app.ListGalleries.side_effect = galleries
    return app


class _FakeWrapFormat:
    """Mimics Shape.WrapFormat — `.Type`, `.Side`, `.Distance*`, `.AllowOverlap`."""

    def __init__(self) -> None:
        self.Type = 7  # wdWrapInline until ConvertToShape / set_wrap sets a real wrap
        self.AllowOverlap = False
        self.Side = 0
        self.DistanceTop = 0.0
        self.DistanceBottom = 0.0
        self.DistanceLeft = 0.0
        self.DistanceRight = 0.0


class _FakePictureFormat:
    """Mimics Shape/InlineShape.PictureFormat — the four crop insets (points)."""

    def __init__(self) -> None:
        self.CropLeft = 0.0
        self.CropTop = 0.0
        self.CropRight = 0.0
        self.CropBottom = 0.0


class _FakeInlineShape:
    """The InlineShape returned by Range.InlineShapes.AddPicture()."""

    def __init__(self, owner: _FakeInlineShapes) -> None:
        self._owner = owner
        self.Width = 100.0
        self.Height = 80.0
        self.AlternativeText = ""
        self.LockAspectRatio = -1
        self.PictureFormat = _FakePictureFormat()
        self.converted = None  # the floating _FakeFloatingShape, once ConvertToShape() runs

    def ConvertToShape(self) -> _FakeFloatingShape:
        # A floating picture shape — appended to the document's body Shapes (so
        # insert_image can hand back a shape:N handle), mirroring live Word.
        doc_shapes = getattr(self._owner, "_doc_shapes", None)
        anchor_start = int(getattr(self._owner, "_anchor_start", 0))
        owner = doc_shapes if doc_shapes is not None else _FakeFloatingShapes()
        count = len(owner._items) if doc_shapes is not None else 0
        shape = _FakeFloatingShape(
            owner,
            shape_type=13,  # msoPicture
            anchor_start=anchor_start,
            width=self.Width,
            height=self.Height,
            alt_text=self.AlternativeText,
            name=f"Picture {count + 1}",  # auto-named like live Word
        )
        if doc_shapes is not None:
            doc_shapes._items.append(shape)
        self.converted = shape
        return shape


class _FakeInlineShapes:
    """Mimics Range.InlineShapes: AddPicture records its call and returns a shape.

    `_doc_shapes` / `_anchor_start` are wired by the document's range factory so a
    `ConvertToShape()` lands the floating picture in `Document.Shapes`.
    """

    def __init__(self) -> None:
        self._doc_shapes: _FakeFloatingShapes | None = None
        self._anchor_start = 0
        self.shape = _FakeInlineShape(self)
        self.AddPicture = MagicMock(name="AddPicture", return_value=self.shape)

    @property
    def Count(self) -> int:
        # No picture present unless a test wires one up; range_text() reads this
        # to decide whether to tokenize inline shapes.
        return 0


class _FakeDocImage:
    """A document-level inline picture, as `Document.InlineShapes.Item(n)` returns.

    Its `Range.WordOpenXML` carries exactly one image part, so `read_image()`
    against the shape's own range round-trips the seeded bytes + MIME.
    """

    def __init__(
        self,
        *,
        mime: str,
        data: bytes,
        width: float,
        height: float,
        alt_text: str,
        start: int,
    ) -> None:
        rng = _make_range(start, start + 1)
        rng.WordOpenXML = _flat_opc([(mime, data)])
        self.Range = rng
        self.Width = width
        self.Height = height
        self.LockAspectRatio = -1
        self.AlternativeText = alt_text
        self.PictureFormat = _FakePictureFormat()
        self.Type = 3  # wdInlineShapePicture


class _FakeDocInlineShapes:
    """Mimics `Document.InlineShapes`: Count, 1-based Item()/call lookup, iteration."""

    def __init__(self, images: list[dict[str, Any]] | None) -> None:
        self._items: list[_FakeDocImage] = []
        for spec in images or []:
            self._items.append(
                _FakeDocImage(
                    mime=spec.get("mime", "image/png"),
                    data=spec.get("data", b"\x89PNG\r\n\x1a\nFAKE"),
                    width=spec.get("width", 100.0),
                    height=spec.get("height", 80.0),
                    alt_text=spec.get("alt_text", ""),
                    start=spec.get("start", 0),
                )
            )

    @property
    def Count(self) -> int:
        return len(self._items)

    def Item(self, index: int) -> _FakeDocImage:
        return self._items[index - 1]

    def __call__(self, index: int) -> _FakeDocImage:
        return self._items[index - 1]

    def __iter__(self) -> Iterable[Any]:
        return iter(self._items)


# ---------------------------------------------------------------------------
# Charts — the Excel-backed AddChart2 surface, faked enough for the wiring tests.
# Mirrors the live mechanics _charts.populate_chart drives: write data into the
# embedded workbook's cells, point a single series at them via a SERIES formula,
# title, BreakLink, then close the workbook. The fake captures all of it.
# ---------------------------------------------------------------------------


class _FakeCell:
    def __init__(self, store: dict[tuple[int, int], Any], row: int, col: int) -> None:
        self._store = store
        self._row = row
        self._col = col

    @property
    def Value(self) -> Any:
        return self._store.get((self._row, self._col))

    @Value.setter
    def Value(self, value: Any) -> None:
        self._store[(self._row, self._col)] = value

    @property
    def Address(self) -> str:
        return f"${chr(ord('A') + self._col - 1)}${self._row}"


class _FakeCells:
    """`Worksheet.Cells(r, c)` — a callable cell accessor over a shared value store."""

    def __init__(self) -> None:
        self.values: dict[tuple[int, int], Any] = {}

    def __call__(self, row: int, col: int) -> _FakeCell:
        return _FakeCell(self.values, row, col)


class _FakeCellRange:
    def __init__(self, a: _FakeCell, b: _FakeCell) -> None:
        self._a = a
        self._b = b

    @property
    def Address(self) -> str:
        return f"{self._a.Address}:{self._b.Address}"


class _FakeChartWorksheet:
    def __init__(self) -> None:
        self.Name = "Sheet1"
        self.Cells = _FakeCells()
        self.UsedRange = MagicMock(name="UsedRange")

    def Range(self, a: _FakeCell, b: _FakeCell) -> _FakeCellRange:
        return _FakeCellRange(a, b)


class _FakeChartWorkbook:
    """The chart's embedded Excel workbook (`ChartData.Workbook`)."""

    def __init__(self) -> None:
        self._ws = _FakeChartWorksheet()
        self.Application = MagicMock(name="EmbeddedExcel")
        # Empty after Close, so populate_chart's quit-if-empty branch fires.
        self.Application.Workbooks.Count = 0
        self.closed = False

    def Worksheets(self, index: int) -> _FakeChartWorksheet:
        return self._ws

    def Close(self, save_changes: bool = False) -> None:
        self.closed = True


class _FakeChartData:
    def __init__(self, workbook: _FakeChartWorkbook) -> None:
        self.Workbook = workbook
        self.linked = True

    def Activate(self) -> None:
        pass

    def BreakLink(self) -> None:
        self.linked = False


# --- formatting surface (post-insert, static chart) -------------------------


class _FakeColorTarget:
    """A `.Format.Fill` / `.Format.Line` ForeColor holder — captures `.RGB`."""

    def __init__(self) -> None:
        self.ForeColor = SimpleNamespace(RGB=None)


class _FakeShapeFormat:
    def __init__(self) -> None:
        self.Fill = _FakeColorTarget()
        self.Line = _FakeColorTarget()


class _FakeTrendline:
    def __init__(self, type_: int) -> None:
        self.Type = type_
        self.DisplayEquation = False
        self.DisplayRSquared = False
        self.Forward = None
        self.Backward = None
        self.Order = None
        self.Period = None


class _FakeTrendlines:
    def __init__(self, series: _FakeSeries) -> None:
        self._series = series

    def Add(self, type_: int) -> _FakeTrendline:
        tl = _FakeTrendline(int(type_))
        self._series.trendlines.append(tl)
        return tl


class _FakePoint:
    def __init__(self) -> None:
        self.Format = _FakeShapeFormat()
        self.MarkerStyle = None
        self.Explosion = None
        self.DataLabel = SimpleNamespace(Font=SimpleNamespace())


class _FakeAxis:
    def __init__(self) -> None:
        self.HasTitle = False
        self.AxisTitle = SimpleNamespace(Text="")
        self.MinimumScale = None
        self.MaximumScale = None
        self.ScaleType = None
        self.TickLabels = SimpleNamespace(NumberFormat="")
        self.HasMajorGridlines = False


class _FakeSeries:
    def __init__(self, collection: _FakeSeriesCollection) -> None:
        self._collection = collection
        self.Formula = ""
        self.Name = ""
        self.Format = _FakeShapeFormat()
        self.HasDataLabels = False
        self._data_labels = SimpleNamespace(NumberFormat="", Font=SimpleNamespace())
        self._points: dict[int, _FakePoint] = {}
        self.trendlines: list[_FakeTrendline] = []
        # Chart-depth surface (markers / line / pie / error bars).
        self.MarkerStyle = None
        self.MarkerSize = None
        self.Smooth = None
        self.Explosion = None
        self.HasErrorBars = False
        self.error_bars: list[tuple[int, int, int, float]] = []

    def Points(self, index: int) -> _FakePoint:
        return self._points.setdefault(int(index), _FakePoint())

    def DataLabels(self) -> SimpleNamespace:
        if not self.HasDataLabels:  # mirror Word: DataLabels() raises until shown
            raise RuntimeError("Unable to get the Count property of the DataLabels class")
        return self._data_labels

    def ErrorBar(self, direction: int, include: int, type_: int, amount: float) -> None:
        self.error_bars.append((int(direction), int(include), int(type_), float(amount)))

    def Trendlines(self) -> _FakeTrendlines:
        return _FakeTrendlines(self)

    def Delete(self) -> None:
        self._collection._items.remove(self)


class _FakeSeriesCollection:
    """`Chart.SeriesCollection()` — seeded with AddChart2's 3 placeholder series."""

    def __init__(self, count: int = 3) -> None:
        self._items: list[_FakeSeries] = [_FakeSeries(self) for _ in range(count)]

    @property
    def Count(self) -> int:
        return len(self._items)

    def __call__(self, index: int) -> _FakeSeries:
        return self._items[index - 1]

    def NewSeries(self) -> _FakeSeries:
        series = _FakeSeries(self)
        self._items.append(series)
        return series


class _FakeChart:
    def __init__(self, chart_type: int) -> None:
        self.ChartType = int(chart_type)
        self.HasTitle = False
        self.ChartTitle = SimpleNamespace(Text="")
        self._series = _FakeSeriesCollection()
        self.ChartData = _FakeChartData(_FakeChartWorkbook())
        # Formatting surface (mutated by the post-insert format verbs).
        self.ChartStyle = 201
        self.HasLegend = False
        self.Legend = SimpleNamespace(Position=None)
        self.ChartArea = SimpleNamespace(Format=_FakeShapeFormat(), Font=SimpleNamespace())
        self.PlotArea = SimpleNamespace(Format=_FakeShapeFormat(), Font=SimpleNamespace())
        self._axes: dict[tuple[int, int], _FakeAxis] = {}
        # Chart-depth surface.
        self.HasDataTable = False
        self.DataTable = SimpleNamespace(HasBorderOutline=False)
        self._groups: dict[int, SimpleNamespace] = {}

    def SeriesCollection(self, index: int | None = None) -> Any:
        # Real Word: no index → the collection; with an index → that Series.
        return self._series if index is None else self._series(int(index))

    def Axes(self, axis_type: int, group: int = 1) -> _FakeAxis:
        return self._axes.setdefault((int(axis_type), int(group)), _FakeAxis())

    def ChartGroups(self, index: int) -> SimpleNamespace:
        return self._groups.setdefault(
            int(index), SimpleNamespace(GapWidth=None, Overlap=None, FirstSliceAngle=None)
        )


class _FakeChartInlineShape:
    """A document chart inline shape — what AddChart2 returns and inserts."""

    def __init__(self, chart_type: int, start: int, owner: _FakeDocInlineShapes) -> None:
        self.HasChart = True
        self.Type = 12  # wdInlineShapeChart
        self.Chart = _FakeChart(chart_type)
        self.Range = _make_range(start, start + 1)
        self._owner = owner

    def Delete(self) -> None:
        if self in self._owner._items:
            self._owner._items.remove(self)


def _wire_chart_insertion(doc: MagicMock) -> None:
    """Make `Selection.InlineShapes.AddChart2` insert a chart into `doc.InlineShapes`.

    AddChart2 only works off the live Selection (a Range raises in real Word), so
    `insert_chart` selects the point then calls `Application.Selection.InlineShapes
    .AddChart2`. The fake appends the new chart shape to the document's inline
    shapes (incrementing positions) so `chart:N` / `doc.charts` resolve it.
    """

    def add_chart2(style: int, chart_type: int) -> _FakeChartInlineShape:
        items = doc.InlineShapes._items
        start = (len(items) + 1) * 100
        shape = _FakeChartInlineShape(int(chart_type), start, doc.InlineShapes)
        items.append(shape)
        return shape

    doc.Application.Selection.InlineShapes.AddChart2 = add_chart2


class _FakeOMath:
    """A single equation zone, as `Document.OMaths(n)` returns.

    `Type` is wdOMathType (1=display, 0=inline). `Range.Text` carries the
    built-up form (with the internal `\\r` structure markers the reader strips for
    its linear preview). `BuildUp`/`Linearize` are recorded but don't restructure
    the fake text. `Range.WordOpenXML` is left as a plain Flat OPC skeleton — the
    OMML→MathML read goes through MSXML, so `mathml` is exercised only by the
    smoke suite, not here.
    """

    def __init__(self, *, rng: Any, type_: int, text: str) -> None:
        rng.Text = text
        self.Range = rng
        self.Type = int(type_)
        self.built = True

    def BuildUp(self) -> None:
        self.built = True

    def Linearize(self) -> None:
        self.built = False


class _FakeDocOMaths:
    """Mimics `Document.OMaths`: Count, 1-based Item()/call lookup, iteration.

    Equations are kept in document order (sorted by range Start) so `equation:N`
    and `EquationCollection.list()` see Word's own ordering.
    """

    def __init__(self, equations: list[dict[str, Any]] | None, range_factory: Any) -> None:
        self._items: list[_FakeOMath] = []
        for spec in equations or []:
            start = int(spec.get("start", 0))
            end = int(spec.get("end", start + 1))
            self._items.append(
                _FakeOMath(
                    rng=range_factory(start, end),
                    type_=spec.get("type", 1),
                    text=spec.get("text", ""),
                )
            )
        self._items.sort(key=lambda o: int(o.Range.Start))

    @property
    def Count(self) -> int:
        return len(self._items)

    def Item(self, index: int) -> _FakeOMath:
        return self._items[index - 1]

    def __call__(self, index: int) -> _FakeOMath:
        return self._items[index - 1]

    def __iter__(self) -> Iterable[Any]:
        return iter(list(self._items))


class _FakeBorders:
    """Mimics `Range.Borders` — a callable vending one stable child per edge index.

    A bare MagicMock would return the *same* `return_value` for `Borders(-1)` and
    `Borders(-3)`, so a per-side assert (top vs. bottom) couldn't tell them apart.
    Memoising one child per index keeps each edge's `LineStyle`/`LineWidth`/`Color`
    distinct, which is what `set_borders` writes.
    """

    def __init__(self) -> None:
        self._edges: dict[int, MagicMock] = {}

    def __call__(self, index: int) -> MagicMock:
        return self._edges.setdefault(int(index), MagicMock(name=f"Border[{index}]"))


class _FakeVariable:
    """A single document variable with a settable Value and a Delete()."""

    def __init__(self, registry: _FakeVariables, name: str, value: str) -> None:
        self._registry = registry
        self.Name = name
        self.Value = value

    def Delete(self) -> None:
        self._registry._remove(self.Name)


class _FakeVariables:
    """Mimics doc.Variables: Count, by-name call lookup, iteration, Add(Name, Value)."""

    def __init__(self, variables: dict[str, str] | None) -> None:
        self._items: dict[str, _FakeVariable] = {
            k: _FakeVariable(self, k, str(v)) for k, v in (variables or {}).items()
        }

    @property
    def Count(self) -> int:
        return len(self._items)

    def __call__(self, name: str) -> _FakeVariable:
        return self._items[name]

    def __iter__(self) -> Iterable[Any]:
        return iter(list(self._items.values()))

    def Add(self, Name: str | None = None, Value: str = "") -> _FakeVariable:
        v = _FakeVariable(self, str(Name), str(Value))
        self._items[str(Name)] = v
        return v

    def _remove(self, name: str) -> None:
        self._items.pop(name, None)


class _FakeHyperlink:
    """A single hyperlink: Address / SubAddress / TextToDisplay / ScreenTip / Range."""

    def __init__(self, spec: dict[str, Any]) -> None:
        self.Address = spec.get("address", "")
        self.SubAddress = spec.get("sub_address", "")
        self.TextToDisplay = spec.get("text", "")
        self.ScreenTip = spec.get("screen_tip", "")
        self.Range = _make_range(spec.get("start", 0), spec.get("end", 0))


class _FakeHyperlinks:
    """Mimics doc.Hyperlinks: Count, 1-based call lookup, iteration, Add.

    `Add` is a MagicMock (side-effect appends a `_FakeHyperlink`) so `link_to`'s
    positional call args stay assertable while the reader still sees the new link.
    """

    def __init__(self, links: list[dict[str, Any]] | None) -> None:
        self._items = [_FakeHyperlink(spec) for spec in (links or [])]
        self.Add = MagicMock(name="HyperlinksAdd", side_effect=self._add)

    def _add(
        self,
        Anchor: Any = None,
        Address: str = "",
        SubAddress: str = "",
        ScreenTip: str = "",
        TextToDisplay: str = "",
        *args: Any,
    ) -> _FakeHyperlink:
        start = int(getattr(Anchor, "Start", 0)) if Anchor is not None else 0
        end = int(getattr(Anchor, "End", start)) if Anchor is not None else start
        h = _FakeHyperlink(
            {
                "address": str(Address or ""),
                "sub_address": str(SubAddress or ""),
                "screen_tip": str(ScreenTip or ""),
                "text": str(TextToDisplay or ""),
                "start": start,
                "end": end,
            }
        )
        self._items.append(h)
        return h

    @property
    def Count(self) -> int:
        return len(self._items)

    def __call__(self, index: int) -> _FakeHyperlink:
        return self._items[index - 1]

    def __iter__(self) -> Iterable[Any]:
        return iter(list(self._items))


class _FakeField:
    """A single field: Code (range w/ Text), Result (w/ Text), Type, Locked."""

    def __init__(self, spec: dict[str, Any]) -> None:
        start = int(spec.get("start", 0))
        code = _make_range(start, int(spec.get("end", start + 1)))
        code.Text = spec.get("code", "")
        self.Code = code
        result = MagicMock(name="FieldResult")
        result.Text = spec.get("result", "")
        self.Result = result
        self.Type = int(spec.get("type", 33))  # wdFieldPage
        self.Locked = bool(spec.get("locked", False))


class _FakeFields:
    """Mimics doc.Fields: Count, 1-based call lookup, iteration, Update().

    `Update` is a MagicMock so `Document.update_fields` (which calls
    `Fields.Update()`) stays assertable.
    """

    def __init__(self, fields: list[dict[str, Any]] | None) -> None:
        self._items = [_FakeField(spec) for spec in (fields or [])]
        self.Update = MagicMock(name="FieldsUpdate")

    @property
    def Count(self) -> int:
        return len(self._items)

    def __call__(self, index: int) -> _FakeField:
        return self._items[index - 1]

    def __iter__(self) -> Iterable[Any]:
        return iter(list(self._items))


class _FakeDocProperty:
    """A built-in / custom document property with name + (maybe read-only) value.

    `unreadable=True` makes `.Value` raise on read (an unset built-in, e.g. a
    date), which the reader skips. `readonly=True` makes `.Value` raise on write
    (a computed stat), which `set` surfaces as an OpError.
    """

    def __init__(
        self,
        registry: _FakeDocProperties,
        name: str,
        value: Any,
        *,
        readonly: bool = False,
        unreadable: bool = False,
    ) -> None:
        self._registry = registry
        self.Name = name
        self._value = value
        self._readonly = readonly
        self._unreadable = unreadable

    @property
    def Value(self) -> Any:
        if self._unreadable:
            raise RuntimeError(f"property {self.Name!r} has no value")
        return self._value

    @Value.setter
    def Value(self, new: Any) -> None:
        if self._readonly:
            raise RuntimeError(f"property {self.Name!r} is read-only")
        self._value = new

    def Delete(self) -> None:
        self._registry._remove(self.Name)


# The standard built-in property names real Word always exposes (a representative
# slice). The fake auto-vivifies any of these with an empty value on first access,
# so setting an unset-but-standard built-in (e.g. "Subject") works as in real Word;
# a name outside this set raises, mirroring Word's com_error for an unknown one.
_BUILTIN_PROP_NAMES = frozenset(
    {
        "Title",
        "Subject",
        "Author",
        "Keywords",
        "Comments",
        "Template",
        "Last author",
        "Revision number",
        "Application name",
        "Last print date",
        "Creation date",
        "Last save time",
        "Total editing time",
        "Number of pages",
        "Number of words",
        "Number of characters",
        "Security",
        "Category",
        "Format",
        "Manager",
        "Company",
        "Hyperlink base",
        "Content status",
        "Word count",
    }
)


class _FakeDocProperties:
    """Mimics BuiltIn/CustomDocumentProperties: Count, by-name lookup, iter, Add.

    `known` (the built-in bag's standard names) makes `__call__` auto-vivify an
    unset-but-standard property; `None` (the custom bag) raises for any name not
    already present, matching Word's com_error.
    """

    def __init__(
        self,
        props: dict[str, Any] | None,
        *,
        readonly: set[str] | None = None,
        unreadable: set[str] | None = None,
        known: frozenset[str] | None = None,
    ) -> None:
        ro, ur = readonly or set(), unreadable or set()
        self._known = known
        self._items: dict[str, _FakeDocProperty] = {
            name: _FakeDocProperty(self, name, value, readonly=name in ro, unreadable=name in ur)
            for name, value in (props or {}).items()
        }

    @property
    def Count(self) -> int:
        return len(self._items)

    def __call__(self, name: str) -> _FakeDocProperty:
        if name not in self._items:
            if self._known is not None and name in self._known:
                self._items[name] = _FakeDocProperty(self, name, "")
            else:
                # Real Word raises a com_error for an unknown name; KeyError is
                # close enough for the wrapper's catch-and-reraise-as-OpError path.
                raise KeyError(name)
        return self._items[name]

    def __iter__(self) -> Iterable[Any]:
        return iter(list(self._items.values()))

    def Add(
        self,
        Name: str | None = None,
        LinkToContent: bool = False,
        Type: int = 4,
        Value: Any = "",
        *args: Any,
    ) -> _FakeDocProperty:
        p = _FakeDocProperty(self, str(Name), Value)
        self._items[str(Name)] = p
        return p

    def _remove(self, name: str) -> None:
        self._items.pop(name, None)


class _FakeProofErrors:
    """Mimics doc.SpellingErrors / doc.GrammaticalErrors: a ProofreadingErrors collection.

    Each item is a Range over the flagged run (Start/End/Text). `raises=True`
    simulates a checker that's unavailable (proofing off / protected document),
    so `Count` access raises and the reader reports `count: None`.
    """

    def __init__(self, errors: list[dict[str, Any]] | None, *, raises: bool = False) -> None:
        self._raises = raises
        self._items: list[Any] = []
        for spec in errors or []:
            start = int(spec.get("start", 0))
            rng = _make_range(start, int(spec.get("end", start + 1)))
            rng.Text = spec.get("text", "")
            self._items.append(rng)

    @property
    def Count(self) -> int:
        if self._raises:
            raise RuntimeError("proofing tools unavailable")
        return len(self._items)

    def Item(self, index: int) -> Any:
        return self._items[index - 1]

    def __iter__(self) -> Iterable[Any]:
        return iter(list(self._items))


class _FakeReadabilityStat:
    def __init__(self, name: str, value: float) -> None:
        self.Name = name
        self.Value = value


class _FakeReadability:
    """Mimics doc.ReadabilityStatistics: iterable of {Name, Value} stats."""

    def __init__(self, stats: dict[str, float] | None, *, raises: bool = False) -> None:
        self._raises = raises
        self._items = [_FakeReadabilityStat(k, float(v)) for k, v in (stats or {}).items()]

    def __iter__(self) -> Iterable[Any]:
        if self._raises:
            raise RuntimeError("readability unavailable")
        return iter(list(self._items))


def _flat_opc(parts: list[tuple[str, bytes]]) -> str:
    """Build a minimal Flat OPC package string, as `Range.WordOpenXML` returns.

    `parts` is a list of `(content_type, raw_bytes)` image parts; each is emitted
    as a `<pkg:part>` with a base64 `<pkg:binaryData>` body. A namespace-clean
    document.xml skeleton part is always included so the parser sees the same
    "full package skeleton" shape real Word emits.
    """
    import base64 as _b64

    pieces = [
        '<?xml version="1.0" standalone="yes"?>',
        '<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">',
        '<pkg:part pkg:name="/word/document.xml" '
        'pkg:contentType="application/vnd.openxmlformats-officedocument.'
        'wordprocessingml.document.main+xml"><pkg:xmlData></pkg:xmlData></pkg:part>',
    ]
    for i, (ctype, data) in enumerate(parts, start=1):
        b64 = _b64.b64encode(data).decode("ascii")
        pieces.append(
            f'<pkg:part pkg:name="/word/media/image{i}.bin" pkg:contentType="{ctype}">'
            f"<pkg:binaryData>{b64}</pkg:binaryData></pkg:part>"
        )
    pieces.append("</pkg:package>")
    return "".join(pieces)


def _fake_font(**kw: Any) -> MagicMock:
    """A `Range.Font` / `Style.Font` mock pre-populated with Word-like defaults.

    Real values where set (so `format_info`'s `float(Font.Size)` etc. work and
    `format_run` writes round-trip through the same handle), MagicMock fallback
    for anything not modelled. Defaults match Aptos 12pt, no emphasis, black.
    """
    f = MagicMock(name="Font")
    f.Name = kw.get("Name", "Aptos")
    f.Size = kw.get("Size", 12.0)
    f.Bold = kw.get("Bold", 0)
    f.Italic = kw.get("Italic", 0)
    f.Underline = kw.get("Underline", 0)
    f.StrikeThrough = kw.get("StrikeThrough", 0)
    f.Color = kw.get("Color", 0)
    f.Subscript = kw.get("Subscript", 0)
    f.Superscript = kw.get("Superscript", 0)
    f.SmallCaps = kw.get("SmallCaps", 0)
    f.AllCaps = kw.get("AllCaps", 0)
    f.Spacing = kw.get("Spacing", 0.0)
    f.Hidden = kw.get("Hidden", 0)
    return f


def _fake_paragraph_format(**kw: Any) -> MagicMock:
    """A `Range.ParagraphFormat` / `Style.ParagraphFormat` mock with real defaults
    (single spacing, 8pt after, widow control on — Word's Normal baseline)."""
    pf = MagicMock(name="ParagraphFormat")
    pf.Alignment = kw.get("Alignment", 0)
    pf.LeftIndent = kw.get("LeftIndent", 0.0)
    pf.RightIndent = kw.get("RightIndent", 0.0)
    pf.FirstLineIndent = kw.get("FirstLineIndent", 0.0)
    pf.SpaceBefore = kw.get("SpaceBefore", 0.0)
    pf.SpaceAfter = kw.get("SpaceAfter", 8.0)
    pf.LineSpacingRule = kw.get("LineSpacingRule", 0)  # wdLineSpaceSingle
    pf.LineSpacing = kw.get("LineSpacing", 12.0)
    pf.PageBreakBefore = kw.get("PageBreakBefore", 0)
    pf.KeepTogether = kw.get("KeepTogether", 0)
    pf.KeepWithNext = kw.get("KeepWithNext", 0)
    pf.WidowControl = kw.get("WidowControl", -1)
    return pf


def _fake_para_style(name: str = "Normal") -> MagicMock:
    """A `Range.ParagraphStyle` mock — the applied style's baseline that
    `format_info` compares the effective formatting against."""
    st = MagicMock(name=f"ParagraphStyle[{name}]")
    st.NameLocal = name
    st.Font = _fake_font()
    st.ParagraphFormat = _fake_paragraph_format()
    return st


class _FakeWords:
    """Mimics Range.Words: word objects derived from the range's current `Text`.

    Iterated lazily so it reflects a `Text` set after `_make_range` (the paragraph
    builder overrides it). Each word carries `.Text`, an absolute `.Start`, and
    the range's `.Font` — enough for `_export`'s emphasis walk (the fake has no
    per-word formatting, so words read the paragraph's uniform font).
    """

    def __init__(self, rng: MagicMock) -> None:
        self._rng = rng

    def __iter__(self) -> Iterable[Any]:
        text = self._rng.Text or ""
        base = int(self._rng.Start)
        for m in re.finditer(r"\S+\s*|\s+", text):
            w = MagicMock(name="Word")
            w.Text = m.group(0)
            w.Start = base + m.start()
            w.End = base + m.end()
            w.Font = self._rng.Font
            yield w


def _make_range(start: int, end: int) -> MagicMock:
    rng = MagicMock(name=f"Range[{start},{end}]")
    rng.Start = start
    rng.End = end
    rng.Text = ""
    # Export support: Range.Words (lazy, derived from Text) for the to_markdown
    # emphasis walk, and an empty Range.Hyperlinks (tests seed their own).
    rng.Words = _FakeWords(rng)
    rng.Hyperlinks = []
    # Image-extraction support: WordOpenXML serialises the range as Flat OPC.
    # A plain range carries no media part, so read_image() on it reports "no
    # image"; image shape ranges (and tests) override this with a real part.
    rng.WordOpenXML = _flat_opc([])
    # List support: a stateful ListFormat plus an Application that vends
    # list-gallery templates, so apply_list / list_info / restart work.
    rng.ListFormat = _FakeListFormat()
    rng.Application = _list_application()
    # Image support: every range can take an inline picture and knows its
    # section's page geometry (Letter defaults) for the wrap="auto" heuristic.
    rng.InlineShapes = _FakeInlineShapes()
    rng.PageSetup = _FakePageSetup()
    # A plain range contains no tables; find/replace's segment fast-path checks
    # Range.Tables.Count to decide whether table-boundary segmentation is needed.
    # Tests that exercise the in-table path set a positive Count on their range.
    rng.Tables = MagicMock(name="RangeTables")
    rng.Tables.Count = 0
    # Snapshot support: Range.Information(wdActiveEndPageNumber=3) reports the
    # page a (collapsed) range sits on. The tiny fake document is all one page,
    # so report page 1 for any query; tests that need a specific page set
    # `rng.Information.return_value` on the cached range themselves.
    rng.Information = MagicMock(name="Information", return_value=1)
    # Border support: one stable child per edge so per-side asserts work.
    rng.Borders = _FakeBorders()
    # Revision support: a range scopes accept_all/reject_all(within=...). Empty by
    # default — tests that need a populated range seed rng.Revisions themselves.
    rng.Revisions = _FakeRevisions([])
    # Format-read support: real Font / ParagraphFormat values (so format_info and
    # the consistency linter rules can probe them) plus the applied ParagraphStyle
    # baseline they compare against. A clean range reads no overrides because the
    # effective values and the style baseline share the same defaults.
    rng.Font = _fake_font()
    rng.ParagraphFormat = _fake_paragraph_format()
    rng.ParagraphStyle = _fake_para_style()
    # Format-read support: highlight is a Range property (not on Font). Default 0
    # (wdNoHighlight -> "none"); tests that need a highlighted run set it themselves.
    rng.HighlightColorIndex = 0
    return rng


_DEFAULT_STYLES: tuple[dict[str, Any], ...] = (
    {"name": "Normal", "type": 1, "builtin": True, "in_use": True},
    {"name": "Body Text", "type": 1, "builtin": True, "in_use": True},
    {"name": "Heading 1", "type": 1, "builtin": True, "in_use": True},
    {"name": "Heading 2", "type": 1, "builtin": True, "in_use": True},
    {"name": "Heading 3", "type": 1, "builtin": True, "in_use": True},
    {"name": "List Bullet", "type": 1, "builtin": True, "in_use": True},
    {"name": "List Number", "type": 1, "builtin": True, "in_use": True},
)


class _FakeStyles:
    """Mimics ActiveDocument.Styles: iterable of objects exposing NameLocal/Type/BuiltIn/InUse.

    Note: real Word's `Styles(name)` raises a generic com_error for missing
    names. wordlive validates membership via iteration *first*, so the fake
    only needs to support iteration. We still implement __call__ for the rare
    direct lookup path; it raises KeyError if the style isn't there.
    """

    def __init__(self, styles: tuple[dict[str, Any], ...] | list[dict[str, Any]]) -> None:
        self._items: list[Any] = []
        for s in styles:
            mock = MagicMock(name=f"Style[{s['name']}]")
            mock.NameLocal = s["name"]
            mock.Type = s.get("type", 1)
            mock.BuiltIn = s.get("builtin", True)
            mock.InUse = s.get("in_use", True)
            self._items.append(mock)

    def __iter__(self) -> Iterable[Any]:
        return iter(self._items)

    def __call__(self, name: str) -> Any:
        for s in self._items:
            if s.NameLocal == name:
                return s
        raise KeyError(name)

    def Add(self, Name: str | None = None, Type: int = 1) -> Any:
        """Mimic Styles.Add(Name, Type): append a new style and return its mock.

        `StyleCollection.add` then re-resolves it by name via `Style.com`'s
        direct-then-iterate lookup, so the new style must be iterable here.
        """
        mock = MagicMock(name=f"Style[{Name}]")
        mock.NameLocal = Name
        mock.Type = int(Type)
        mock.BuiltIn = False
        mock.InUse = False
        self._items.append(mock)
        return mock


class _FakeTable:
    """Mimics a Word Table COM object: Rows/Columns counts, Cell(r,c), row add/delete.

    Cells are backed by persistent MagicMock ranges so writes (`Range.Text`,
    `Range.Style`, `Range.ParagraphFormat`) round-trip through the same handle.
    Cell text is seeded with Word's trailing `\\r\\x07` markers so the marker-
    stripping path in `Cell.text` gets exercised.
    """

    def __init__(
        self,
        grid: list[list[str]],
        title: str = "",
        *,
        start: int = 0,
        owner: Any | None = None,
        style: str | None = None,
    ) -> None:
        self.Title = title
        # Word's Table.Style is a Style object (its NameLocal names the table
        # style); None until one is applied. The `table-style-consistent` linter
        # rule reads `.Style.NameLocal`, so a test can seed it via `style=`.
        if style is not None:
            self.Style = MagicMock(name=f"TableStyle[{style}]")
            self.Style.NameLocal = style
        else:
            self.Style = None
        self._start = start
        self._owner = owner
        self._rows = [[self._mk_cell_range(text) for text in row] for row in grid]
        self._range_mock: Any | None = None
        self._rows_view: _FakeRows | None = None
        self._cell_mocks: dict[tuple[int, int], Any] = {}
        self._borders: Any | None = None
        # Autofit state: AllowAutoFit is settable; AutoFitBehavior records its arg.
        self.AllowAutoFit = True
        self.autofit_behavior: int | None = None
        # Banding / table-style options — settable booleans (set_banding), all off
        # by default so a tri-state set_banding flips only the flags it's given.
        self.ApplyStyleHeadingRows = False
        self.ApplyStyleLastRow = False
        self.ApplyStyleFirstColumn = False
        self.ApplyStyleLastColumn = False
        self.ApplyStyleRowBands = False
        self.ApplyStyleColumnBands = False

    def AutoFitBehavior(self, behavior: int) -> None:
        self.autofit_behavior = int(behavior)

    @property
    def Borders(self) -> Any:
        # Whole-table borders (Table.set_borders): one stable child per edge.
        if self._borders is None:
            self._borders = _FakeBorders()
        return self._borders

    @staticmethod
    def _mk_cell_range(text: str) -> MagicMock:
        # Build on _make_range so a cell has the full range surface (Information /
        # Font / ParagraphFormat), which the linter's table-repeat-header rule
        # probes via cell().location(). Default Information -> page 1.
        rng = _make_range(0, 0)
        # Word terminates cell text with CR + the cell mark; populated cells
        # (set via Range.Text) carry no markers, so only seed them for non-empty
        # seed text to mirror both states.
        rng.Text = (text + "\r\x07") if text else "\r\x07"
        return rng

    @property
    def Range(self) -> Any:
        if self._range_mock is None:
            rng = MagicMock(name="TableRange")
            rng.Start = self._start
            # The fake models a table as a standalone object, not via in-document
            # cell-paragraphs, so its range carries no document paragraphs — an
            # empty [Start, End) interval (End defaults to a MagicMock that ints
            # to 1, which would otherwise swallow the paragraph at offset Start in
            # _export's table-interleave walk).
            rng.End = self._start
            self._range_mock = rng
        return self._range_mock

    @property
    def Rows(self) -> Any:
        if self._rows_view is None:
            self._rows_view = _FakeRows(self._rows, self._mk_cell_range)
        return self._rows_view

    @property
    def Columns(self) -> Any:
        return _FakeColumns(self._rows, self._mk_cell_range)

    def Cell(self, row: int, col: int) -> Any:
        # Persist per-(row,col) so cell-level COM property writes (e.g.
        # VerticalAlignment) round-trip through the same handle; Range stays the
        # shared persistent cell range.
        cell = self._cell_mocks.get((row, col))
        if cell is None:
            cell = MagicMock(name=f"Cell[{row},{col}]")
            cell.Range = self._rows[row - 1][col - 1]
            cell.RowIndex = row
            cell.ColumnIndex = col
            rows = self._rows
            mk = self._mk_cell_range
            mocks = self._cell_mocks

            def _merge(other: Any, _r: int = row, _c: int = col) -> None:
                # Simplified horizontal merge within a row: drop the physical
                # cells between this cell and `other`, leaving the row shorter
                # (so Table.is_uniform reports False). Faithful enough for wiring
                # + the non-uniform signal; live smoke covers real COM merging.
                oc = int(getattr(other, "ColumnIndex", _c))
                lo, hi = sorted((_c, oc))
                target = rows[_r - 1]
                joined = "".join(str(getattr(target[i - 1], "Text", "")) for i in range(lo, hi + 1))
                for i in range(hi, lo, -1):
                    if i - 1 < len(target):
                        del target[i - 1]
                if lo - 1 < len(target):
                    target[lo - 1].Text = joined
                mocks.clear()

            def _split(nrows: int = 1, ncols: int = 2, _r: int = row, _c: int = col) -> None:
                # Insert (ncols-1) physical cells after this one in the row, so
                # the row grows and the table goes non-uniform.
                target = rows[_r - 1]
                for _ in range(max(0, int(ncols) - 1)):
                    target.insert(_c, mk(""))
                mocks.clear()

            cell.Merge = _merge
            cell.Split = _split
            self._cell_mocks[(row, col)] = cell
        return cell

    def Delete(self) -> None:
        if self._owner is not None:
            self._owner._remove(self)


class _FakeRows:
    """Rows view: Count, Add() (append at end), Rows(i).Delete().

    Shares the parent table's row list so structural edits persist.
    """

    def __init__(self, rows: list[list[Any]], mk: Any) -> None:
        self._rows = rows
        self._mk = mk
        # Persist row mocks so property writes (HeadingFormat,
        # AllowBreakAcrossPages) round-trip through the same handle.
        self._row_mocks: dict[int, Any] = {}
        # Whole-table alignment across the page (Table.set_alignment writes here).
        self.Alignment = 0

    @property
    def Count(self) -> int:
        return len(self._rows)

    def Add(self, BeforeRow: Any | None = None) -> None:
        ncols = len(self._rows[0]) if self._rows else 1
        self._rows.append([self._mk("") for _ in range(ncols)])

    def __call__(self, index: int) -> Any:
        row = self._row_mocks.get(index)
        if row is None:
            row = MagicMock(name=f"Row[{index}]")
            rows = self._rows
            row.Delete = lambda: rows.__delitem__(index - 1)
            # Real defaults (not truthy MagicMocks) so the linter's
            # table-repeat-header rule reads "no heading row" until set_heading_row
            # flips HeadingFormat; set_heading_row writes both back here.
            row.HeadingFormat = False
            row.AllowBreakAcrossPages = True
            # A whole-row Range (RowAnchor styles table:N:row:R through it) — full
            # range surface so Shading / Borders / Font writes round-trip.
            row.Range = _make_range(0, 0)
            # Live physical-cell count for that row (Table.is_uniform reads
            # Rows(r).Cells.Count; a merge/split changes the row's length).
            row.Cells = _FakeRowCells(self._rows, index)
            self._row_mocks[index] = row
        return row


class _FakeRowCells:
    """`Rows(r).Cells` — only `.Count`, read live off the shared row list."""

    def __init__(self, rows: list[list[Any]], index: int) -> None:
        self._rows = rows
        self._index = index

    @property
    def Count(self) -> int:
        return len(self._rows[self._index - 1]) if self._rows else 0


class _FakeColumns:
    """Callable Columns view: `Count`, `Add()`, and `Columns(c)` -> a `_FakeColumn`.

    Backs ColumnAnchor (`table:N:col:C`): the anchor reads `Columns(c).Cells`,
    each carrying a `RowIndex`, to fan styling out across the column's cells. A
    real Word table with merged / mixed-width cells raises on `Columns(c)`; the
    fake always succeeds (that error path is covered by the live smoke test).
    """

    def __init__(self, rows: list[list[Any]], mk: Any | None = None) -> None:
        self._rows = rows
        self._mk = mk

    @property
    def Count(self) -> int:
        return len(self._rows[0]) if self._rows else 0

    def Add(self, BeforeColumn: Any | None = None) -> Any:
        # Append a physical cell to every row (Table.add_column appends at end).
        mk = self._mk or (lambda _t: MagicMock(name="ColCellRange"))
        for row in self._rows:
            row.append(mk(""))
        return MagicMock(name="Column")

    def __call__(self, col: int) -> Any:
        return _FakeColumn(self._rows, col)


class _FakeColumn:
    def __init__(self, rows: list[list[Any]], col: int) -> None:
        self._rows = rows
        self._col = col

    @property
    def Cells(self) -> list[Any]:
        out: list[Any] = []
        for r in range(len(self._rows)):
            cell = MagicMock(name=f"ColCell[{r + 1},{self._col}]")
            cell.RowIndex = r + 1
            cell.ColumnIndex = self._col
            # Share the row's persistent cell range — in real Word `Columns(c).Cells`
            # and `Table.Cell(r, c)` are the *same* cell, so a style written through
            # one must be visible through the other.
            if self._col - 1 < len(self._rows[r]):
                cell.Range = self._rows[r][self._col - 1]
            out.append(cell)
        return out

    def Delete(self) -> None:
        # Drop this column's physical cell from each row (Table.delete_column).
        for row in self._rows:
            if self._col - 1 < len(row):
                del row[self._col - 1]


class _FakeTablesCollection:
    """Mimics doc.Tables: Count, 1-based call lookup, iteration, Add, Delete.

    `Add(Range, NumRows, NumColumns)` builds an empty `_FakeTable` seeded with
    the insertion range's `Start` and inserts it in document order (sorted by
    Start), so `index_of` can recover its 1-based position the same way it does
    against real Word.
    """

    def __init__(self, tables: list[_FakeTable]) -> None:
        self._tables = tables
        for t in tables:
            t._owner = self

    @property
    def Count(self) -> int:
        return len(self._tables)

    def __call__(self, index: int) -> _FakeTable:
        return self._tables[index - 1]

    def __iter__(self) -> Iterable[Any]:
        return iter(self._tables)

    def Add(
        self, Range: Any = None, NumRows: int = 1, NumColumns: int = 1, *args: Any, **kwargs: Any
    ) -> _FakeTable:
        start = int(getattr(Range, "Start", 0)) if Range is not None else 0
        grid = [["" for _ in range(int(NumColumns))] for _ in range(int(NumRows))]
        t = _FakeTable(grid, start=start, owner=self)
        self._tables.append(t)
        self._tables.sort(key=lambda x: x._start)
        return t

    def _remove(self, table: _FakeTable) -> None:
        self._tables.remove(table)


class _FakeComment:
    """Mimics a Word Comment: Author, Range (body), Scope (anchored range), Done."""

    def __init__(self, registry: _FakeComments, index: int, scope: Any, text: str) -> None:
        self._registry = registry
        self.Index = index
        self.Author = ""
        body = MagicMock(name=f"CommentBody[{index}]")
        body.Text = text
        self.Range = body
        self.Scope = scope if scope is not None else _make_range(0, 0)
        self.Done = False

    def Delete(self) -> None:
        self._registry._remove(self)


class _FakeComments:
    """Mimics doc.Comments: Count, Add(Range, Text), 1-based call lookup, iteration."""

    def __init__(self) -> None:
        self._items: list[_FakeComment] = []

    @property
    def Count(self) -> int:
        return len(self._items)

    def Add(self, Range: Any = None, Text: str = "") -> _FakeComment:
        c = _FakeComment(self, len(self._items) + 1, Range, str(Text))
        self._items.append(c)
        return c

    def __call__(self, index: int) -> _FakeComment:
        return self._items[index - 1]

    def __iter__(self) -> Iterable[Any]:
        return iter(list(self._items))

    def _remove(self, comment: _FakeComment) -> None:
        self._items.remove(comment)
        for i, c in enumerate(self._items, start=1):
            c.Index = i


class _FakeRevision:
    """Mimics a Word Revision: Type (int), Author, Range (with Text), Date, Accept/Reject."""

    def __init__(self, spec: dict[str, Any], owner: _FakeRevisions | None = None) -> None:
        self._owner = owner
        self.Type = int(spec.get("type", 1))  # 1=insert, 2=delete
        self.Author = spec.get("author", "")
        rng = MagicMock(name="RevisionRange")
        rng.Start = int(spec.get("start", 0))
        rng.End = int(spec.get("end", 0))
        rng.Text = spec.get("text", "")
        self.Range = rng
        # A plain ISO string (real Word returns a datetime; the reader accepts both).
        self.Date = spec.get("date", "2026-06-08T12:00:00")

    def Accept(self) -> None:
        if self._owner is not None:
            self._owner._remove(self)

    def Reject(self) -> None:
        if self._owner is not None:
            self._owner._remove(self)


class _FakeRevisions:
    """Mimics doc.Revisions: Count, 1-based call lookup, iteration, Accept/RejectAll."""

    def __init__(self, revisions: list[dict[str, Any]]) -> None:
        self._items = [_FakeRevision(r, owner=self) for r in revisions]

    @property
    def Count(self) -> int:
        return len(self._items)

    def __call__(self, index: int) -> _FakeRevision:
        return self._items[index - 1]

    def __iter__(self) -> Iterable[Any]:
        return iter(list(self._items))

    def _remove(self, revision: _FakeRevision) -> None:
        self._items.remove(revision)

    def AcceptAll(self) -> None:
        self._items.clear()

    def RejectAll(self) -> None:
        self._items.clear()


class _FakeNote:
    """Mimics a Word Footnote/Endnote: a body Range, a Reference range, Delete."""

    def __init__(self, registry: _FakeNotes, ref_start: int, text: str) -> None:
        self._registry = registry
        body = MagicMock(name="NoteBody")
        body.Text = text
        ish = MagicMock(name="NoteInlineShapes")
        ish.Count = 0
        body.InlineShapes = ish
        self.Range = body
        ref = MagicMock(name="NoteReference")
        ref.Start = int(ref_start)
        ref.Delete = MagicMock(name="ReferenceDelete")
        self.Reference = ref

    def Delete(self) -> None:
        self._registry._remove(self)


class _FakeNotes:
    """Mimics doc.Footnotes / doc.Endnotes: Count, Add(Range, Reference, Text), lookup.

    `Add` is a MagicMock (side-effect appends a `_FakeNote`) so tests can assert
    the positional call args — notably the empty `Reference` that auto-numbers.
    """

    def __init__(self) -> None:
        self._items: list[_FakeNote] = []
        self.Add = MagicMock(name="NotesAdd", side_effect=self._add)

    def _add(self, Range: Any = None, Reference: str = "", Text: str = "") -> _FakeNote:
        ref_start = int(getattr(Range, "Start", 0)) if Range is not None else 0
        note = _FakeNote(self, ref_start, str(Text))
        self._items.append(note)
        # Footnotes/Endnotes are ordered by document position, like real Word.
        self._items.sort(key=lambda n: int(n.Reference.Start))
        return note

    @property
    def Count(self) -> int:
        return len(self._items)

    def __call__(self, index: int) -> _FakeNote:
        return self._items[index - 1]

    def __iter__(self) -> Iterable[Any]:
        return iter(list(self._items))

    def _remove(self, note: _FakeNote) -> None:
        self._items.remove(note)


class _FakeTOC:
    """Mimics a Word TableOfContents: a Range plus Update / UpdatePageNumbers."""

    def __init__(self) -> None:
        rng = MagicMock(name="TOCRange")
        rng.Text = "Table of Contents"
        ish = MagicMock(name="TOCInlineShapes")
        ish.Count = 0
        rng.InlineShapes = ish
        self.Range = rng
        self.Update = MagicMock(name="TOCUpdate")
        self.UpdatePageNumbers = MagicMock(name="TOCUpdatePageNumbers")


class _FakeTOCs:
    """Mimics doc.TablesOfContents: Add(...) records its (positional) args, Count, lookup."""

    def __init__(self) -> None:
        self._items: list[_FakeTOC] = []
        # A MagicMock wrapper so tests can assert the positional Add arguments.
        self.Add = MagicMock(name="TOCAdd", side_effect=self._add)

    def _add(self, *args: Any, **kwargs: Any) -> _FakeTOC:
        toc = _FakeTOC()
        self._items.append(toc)
        return toc

    @property
    def Count(self) -> int:
        return len(self._items)

    def __call__(self, index: int) -> _FakeTOC:
        return self._items[index - 1]


class _FakeFieldBlock:
    """A field block (Index / TableOfFigures) with a Range and refresh methods."""

    def __init__(self) -> None:
        self.Range = _make_range(0, 0)
        self.Range.Text = ""
        self.Update = MagicMock(name="FieldBlockUpdate")
        self.UpdatePageNumbers = MagicMock(name="FieldBlockUpdatePageNumbers")


class _FakeIndexes:
    """Mimics doc.Indexes: MarkEntry records its args; Add(...) records + vends an Index."""

    def __init__(self) -> None:
        self._items: list[_FakeFieldBlock] = []
        self.MarkEntry = MagicMock(name="MarkEntry")
        self.Add = MagicMock(name="IndexAdd", side_effect=self._add)

    def _add(self, *args: Any, **kwargs: Any) -> _FakeFieldBlock:
        idx = _FakeFieldBlock()
        self._items.append(idx)
        return idx

    @property
    def Count(self) -> int:
        return len(self._items)

    def __call__(self, index: int) -> _FakeFieldBlock:
        return self._items[index - 1]


class _FakeTablesOfFigures:
    """Mimics doc.TablesOfFigures: Add(...) records its args (positional + keyword)."""

    def __init__(self) -> None:
        self._items: list[_FakeFieldBlock] = []
        self.Add = MagicMock(name="TOFAdd", side_effect=self._add)

    def _add(self, *args: Any, **kwargs: Any) -> _FakeFieldBlock:
        tof = _FakeFieldBlock()
        self._items.append(tof)
        return tof

    @property
    def Count(self) -> int:
        return len(self._items)

    def __call__(self, index: int) -> _FakeFieldBlock:
        return self._items[index - 1]


class _FakeTablesOfAuthorities:
    """Mimics doc.TablesOfAuthorities: Add(...) records its args, vends a field block."""

    def __init__(self) -> None:
        self._items: list[_FakeFieldBlock] = []
        self.Add = MagicMock(name="TOAAdd", side_effect=self._add)

    def _add(self, *args: Any, **kwargs: Any) -> _FakeFieldBlock:
        toa = _FakeFieldBlock()
        self._items.append(toa)
        return toa

    @property
    def Count(self) -> int:
        return len(self._items)

    def __call__(self, index: int) -> _FakeFieldBlock:
        return self._items[index - 1]


class _FakeSource:
    """A bibliography source: Tag/Cited/XML plus a recording Delete()."""

    def __init__(self, tag: str, xml: str, *, cited: bool = True) -> None:
        self.Tag = tag
        self.XML = xml
        self.Cited = cited
        self.Delete = MagicMock(name="SourceDelete")


class _FakeSources:
    """Mimics doc.Bibliography.Sources: Add(xml) records + parses the tag, lookup."""

    def __init__(self) -> None:
        self._items: list[_FakeSource] = []
        self.Add = MagicMock(name="SourcesAdd", side_effect=self._add)

    def _add(self, xml: str) -> _FakeSource:
        m = re.search(r"<b:Tag>(.*?)</b:Tag>", xml)
        src = _FakeSource(m.group(1) if m else "", xml)
        self._items.append(src)
        return src

    @property
    def Count(self) -> int:
        return len(self._items)

    def __call__(self, index: int) -> _FakeSource:
        return self._items[index - 1]

    def __iter__(self) -> Iterable[Any]:
        return iter(list(self._items))


class _FakeBibliography:
    """Mimics doc.Bibliography: a settable BibliographyStyle and a Sources store."""

    def __init__(self) -> None:
        self.Sources = _FakeSources()
        self.BibliographyStyle = "APA"


class _FakeThemeColor:
    """One theme colour: a settable BGR .RGB plus its 1-based scheme index."""

    def __init__(self, index: int, rgb: int) -> None:
        self.ThemeColorSchemeIndex = index
        self.RGB = rgb


class _FakeColorScheme:
    """Mimics ThemeColorScheme: 12 Colors(i), a Count, and recording Load/Save."""

    def __init__(self) -> None:
        # Distinct, in-range BGR defaults so colours read back as unique hexes.
        self._colors = [_FakeThemeColor(i, i * 0x010101) for i in range(1, 13)]
        self.Load = MagicMock(name="ColorSchemeLoad")
        self.Save = MagicMock(name="ColorSchemeSave")

    @property
    def Count(self) -> int:
        return len(self._colors)

    def Colors(self, index: int) -> _FakeThemeColor:
        return self._colors[index - 1]


class _FakeThemeFont:
    """One theme font with a settable .Name (one per script slot)."""

    def __init__(self, name: str) -> None:
        self.Name = name


class _FakeThemeFontCollection:
    """Mimics MajorFont/MinorFont: Item(i) over three script slots."""

    def __init__(self, name: str) -> None:
        self._fonts = [_FakeThemeFont(name) for _ in range(3)]

    @property
    def Count(self) -> int:
        return len(self._fonts)

    def Item(self, index: int) -> _FakeThemeFont:
        return self._fonts[index - 1]


class _FakeThemeFontScheme:
    """Mimics ThemeFontScheme: MajorFont/MinorFont collections + recording Load."""

    def __init__(self) -> None:
        self.MajorFont = _FakeThemeFontCollection("Aptos Display")
        self.MinorFont = _FakeThemeFontCollection("Aptos")
        self.Load = MagicMock(name="FontSchemeLoad")
        self.Save = MagicMock(name="FontSchemeSave")


class _FakeOfficeTheme:
    """Mimics doc.DocumentTheme: the colour, font, and effect schemes."""

    def __init__(self) -> None:
        self.ThemeColorScheme = _FakeColorScheme()
        self.ThemeFontScheme = _FakeThemeFontScheme()
        self.ThemeEffectScheme = MagicMock(name="ThemeEffectScheme")


class _FakeWordList:
    """Mimics a Word List COM object: a Range plus a ListParagraphs count."""

    def __init__(self, rng: Any, count: int) -> None:
        self.Range = rng
        lp = MagicMock(name="ListParagraphs")
        lp.Count = count
        self.ListParagraphs = lp


class _FakeLists:
    """Mimics doc.Lists: Count, 1-based call lookup, iteration."""

    def __init__(self, lists: list[_FakeWordList]) -> None:
        self._lists = lists

    @property
    def Count(self) -> int:
        return len(self._lists)

    def __call__(self, index: int) -> _FakeWordList:
        return self._lists[index - 1]

    def __iter__(self) -> Iterable[Any]:
        return iter(self._lists)


_WHICH_INDEX = {"primary": 1, "first": 2, "even": 3}


class _FakeTextFrame:
    """Mimics Shape.TextFrame — a TextRange whose `.Text` drives `HasText`,
    plus the internal margins / word-wrap `set_text_frame` writes."""

    def __init__(self, text: str = "") -> None:
        self.TextRange = MagicMock(name="TextRange")
        self.TextRange.Text = text
        self.TextRange.Font = MagicMock(name="TextRange.Font")
        self.TextRange.ParagraphFormat = MagicMock(name="TextRange.ParagraphFormat")
        self.MarginLeft = 7.2
        self.MarginRight = 7.2
        self.MarginTop = 3.6
        self.MarginBottom = 3.6
        self.WordWrap = -1

    @property
    def HasText(self) -> bool:
        return bool(self.TextRange.Text)


class _FakeShapeRange:
    """Mimics a ShapeRange (Document.Shapes.Range / Shape.Ungroup result).

    Carries the selected shapes; `Group()` collapses them into one group shape
    (removing the members from the owner and appending the group), mirroring live
    Word's grouping.
    """

    def __init__(self, owner: _FakeFloatingShapes, shapes: list[_FakeFloatingShape]) -> None:
        self._owner = owner
        self._shapes = list(shapes)

    @property
    def Count(self) -> int:
        return len(self._shapes)

    def Item(self, index: int) -> _FakeFloatingShape:
        return self._shapes[int(index) - 1]

    def __iter__(self) -> Iterable[Any]:
        return iter(list(self._shapes))

    def Group(self) -> _FakeFloatingShape:
        children = list(self._shapes)
        for s in children:
            self._owner._items.remove(s)
        group = _FakeFloatingShape(
            self._owner, shape_type=6, name=f"Group {len(self._owner._items) + 1}"
        )  # msoGroup
        group._children = children
        self._owner._items.append(group)
        return group


class _FakeFloatingShape:
    """A floating shape (text box / picture / WordArt) — the `shape:N` restyle surface.

    Carries the layout/appearance attributes the floating-shape mutators read and
    write (Type, Width/Height, Left/Top, Relative*Position, LockAspectRatio,
    AlternativeText, WrapFormat, TextFrame) plus an Anchor range whose `StoryType`
    keeps body shapes apart from header-story watermarks.
    """

    def __init__(
        self,
        owner: _FakeFloatingShapes,
        *,
        name: str = "",
        text: str = "",
        shape_type: int = 17,  # msoTextBox
        anchor_start: int = 0,
        story_type: int = 1,  # wdMainTextStory
        width: float = 200.0,
        height: float = 100.0,
        alt_text: str = "",
    ) -> None:
        self._owner = owner
        self.Name = name
        self.Type = shape_type
        self.Text = text
        self.Width = width
        self.Height = height
        self.LockAspectRatio = -1
        self.AlternativeText = alt_text
        self.Left = 0.0
        self.Top = 0.0
        self.Rotation = 0.0
        self.RelativeHorizontalPosition = 0
        self.RelativeVerticalPosition = 0
        self.WrapFormat = _FakeWrapFormat()
        self.TextFrame = _FakeTextFrame(text)
        self.PictureFormat = _FakePictureFormat()
        self._children: list[_FakeFloatingShape] = []
        anchor = _make_range(anchor_start, anchor_start)
        anchor.StoryType = story_type
        self.Anchor = anchor
        for child in ("TextEffect", "Line", "Fill"):
            setattr(self, child, MagicMock(name=f"Shape.{child}"))

    @property
    def ZOrderPosition(self) -> int:
        # 1-based index in the owner's stack (back-to-front); higher = nearer the front.
        try:
            return self._owner._items.index(self) + 1
        except ValueError:
            return 0

    def ZOrder(self, cmd: int) -> None:
        items = self._owner._items
        i = items.index(self)
        items.remove(self)
        if int(cmd) == 0:  # msoBringToFront
            items.append(self)
        elif int(cmd) == 1:  # msoSendToBack
            items.insert(0, self)
        elif int(cmd) == 2:  # msoBringForward
            items.insert(min(i + 1, len(items)), self)
        else:  # msoSendBackward
            items.insert(max(i - 1, 0), self)

    @property
    def GroupItems(self) -> _FakeShapeRange:
        return _FakeShapeRange(self._owner, self._children)

    def Ungroup(self) -> _FakeShapeRange:
        children = list(self._children)
        self._owner._items.remove(self)
        for c in children:
            self._owner._items.append(c)
        return _FakeShapeRange(self._owner, children)

    def Delete(self) -> None:
        self._owner._remove(self)


class _FakeFloatingShapes:
    """Mimics Range.Shapes / Document.Shapes: AddTextEffect/AddTextbox, Count, lookup."""

    def __init__(self) -> None:
        self._items: list[_FakeFloatingShape] = []

    @property
    def Count(self) -> int:
        return len(self._items)

    def Item(self, index: int) -> _FakeFloatingShape:
        return self._items[int(index) - 1]

    def __call__(self, index: int) -> _FakeFloatingShape:
        return self._items[int(index) - 1]

    def __iter__(self) -> Iterable[Any]:
        return iter(list(self._items))

    def Range(self, names: Any) -> _FakeShapeRange:
        wanted = [names] if isinstance(names, str) else list(names)
        selected = [s for s in self._items if str(s.Name or "") in wanted]
        return _FakeShapeRange(self, selected)

    def AddTextEffect(
        self, PresetTextEffect: int = 0, Text: str = "", **kwargs: Any
    ) -> _FakeFloatingShape:
        # msoTextEffect (WordArt); auto-named like live Word so it can be grouped.
        shape = _FakeFloatingShape(
            self, text=str(Text), shape_type=15, name=f"WordArt {len(self._items) + 1}"
        )
        self._items.append(shape)
        return shape

    def AddTextbox(
        self, Anchor: Any = None, Width: float = 200.0, Height: float = 100.0, **kwargs: Any
    ) -> _FakeFloatingShape:
        start = int(getattr(Anchor, "Start", 0)) if Anchor is not None else 0
        shape = _FakeFloatingShape(
            self,
            shape_type=17,
            anchor_start=start,
            width=float(Width),
            height=float(Height),
            name=f"Text Box {len(self._items) + 1}",  # auto-named like live Word
        )
        self._items.append(shape)
        return shape

    def _remove(self, shape: _FakeFloatingShape) -> None:
        self._items.remove(shape)


class _FakeHeaderFooter:
    """Mimics a Word HeaderFooter: a settable Range, Exists, LinkToPrevious."""

    def __init__(self, text: str = "", *, exists: bool = True, linked: bool = False) -> None:
        rng = _make_range(0, len(text))
        rng.Text = text
        self.Range = rng
        self.Shapes = _FakeFloatingShapes()  # watermark target (header story)
        self.Exists = exists
        self.LinkToPrevious = linked


class _FakeHeadersFooters:
    """Mimics Section.Headers / Section.Footers: 1/2/3-indexed call lookup."""

    def __init__(self, mapping: dict[int, _FakeHeaderFooter]) -> None:
        self._m = mapping

    def __call__(self, index: int) -> _FakeHeaderFooter:
        return self._m[int(index)]


class _FakeTextColumns:
    """Mimics PageSetup.TextColumns: SetCount records its arg; Count/Spacing settable."""

    def __init__(self) -> None:
        self.Count = 1
        self.Spacing = 36.0
        self.SetCount = MagicMock(name="SetCount")


class _FakePageSetup:
    def __init__(self, **kw: Any) -> None:
        self.Orientation = kw.get("orientation", 0)
        self.TopMargin = kw.get("top", 72.0)
        self.BottomMargin = kw.get("bottom", 72.0)
        self.LeftMargin = kw.get("left", 72.0)
        self.RightMargin = kw.get("right", 72.0)
        self.PageWidth = kw.get("width", 612.0)
        self.PageHeight = kw.get("height", 792.0)
        # Item-4 write surface (set_page_setup): paper size, gutter, columns.
        self.PaperSize = kw.get("paper_size", 2)  # wdPaperLetter
        self.Gutter = kw.get("gutter", 0.0)
        self.TextColumns = _FakeTextColumns()


class _FakeSection:
    def __init__(
        self,
        index: int,
        headers: dict[int, _FakeHeaderFooter],
        footers: dict[int, _FakeHeaderFooter],
        page_setup: _FakePageSetup,
    ) -> None:
        self.Index = index
        self.Headers = _FakeHeadersFooters(headers)
        self.Footers = _FakeHeadersFooters(footers)
        self.PageSetup = page_setup


class _FakeSections:
    """Mimics doc.Sections: Count, 1-based call lookup, iteration."""

    def __init__(self, sections: list[_FakeSection]) -> None:
        self._s = sections

    @property
    def Count(self) -> int:
        return len(self._s)

    def __call__(self, index: int) -> _FakeSection:
        return self._s[index - 1]

    def __iter__(self) -> Iterable[Any]:
        return iter(self._s)


def _build_hf_map(texts: dict[str, str]) -> dict[int, _FakeHeaderFooter]:
    """All three header/footer indices, seeded from a {which: text} dict."""
    return {idx: _FakeHeaderFooter(texts.get(name, "")) for name, idx in _WHICH_INDEX.items()}


def _make_document(
    *,
    name: str = "Test.docx",
    full_name: str = r"C:\test\Test.docx",
    bookmarks: dict[str, tuple[int, int]] | None = None,
    content_controls: list[dict[str, Any]] | None = None,
    paragraphs: list[dict[str, Any]] | None = None,
    content: str = "",
    styles: tuple[dict[str, Any], ...] | list[dict[str, Any]] | None = None,
    tables: list[dict[str, Any]] | None = None,
    lists: list[dict[str, Any]] | None = None,
    sections: list[dict[str, Any]] | None = None,
    footnotes: list[dict[str, Any]] | None = None,
    endnotes: list[dict[str, Any]] | None = None,
    revisions: list[dict[str, Any]] | None = None,
    images: list[dict[str, Any]] | None = None,
    equations: list[dict[str, Any]] | None = None,
    variables: dict[str, str] | None = None,
    hyperlinks: list[dict[str, Any]] | None = None,
    fields: list[dict[str, Any]] | None = None,
    builtin_properties: dict[str, Any] | None = None,
    custom_properties: dict[str, Any] | None = None,
    readonly_properties: set[str] | None = None,
    unreadable_properties: set[str] | None = None,
    spelling: list[dict[str, Any]] | None = None,
    grammar: list[dict[str, Any]] | None = None,
    readability: dict[str, float] | None = None,
) -> MagicMock:
    doc = MagicMock(name=f"Document[{name}]")
    doc.Name = name
    doc.FullName = full_name
    doc.InlineShapes = _FakeDocInlineShapes(images)
    doc.Shapes = _FakeFloatingShapes()  # floating-shape target (text boxes)
    # Document-info collections (always present so .Count is a real int, never a
    # MagicMock): variables, hyperlinks, fields, properties, and proofing.
    doc.Variables = _FakeVariables(variables)
    doc.Hyperlinks = _FakeHyperlinks(hyperlinks)
    doc.Fields = _FakeFields(fields)
    doc.BuiltInDocumentProperties = _FakeDocProperties(
        builtin_properties,
        readonly=readonly_properties,
        unreadable=unreadable_properties,
        known=_BUILTIN_PROP_NAMES,
    )
    doc.CustomDocumentProperties = _FakeDocProperties(custom_properties)
    doc.SpellingErrors = _FakeProofErrors(spelling)
    doc.GrammaticalErrors = _FakeProofErrors(grammar)
    doc.ReadabilityStatistics = _FakeReadability(readability)

    bm_registry = _FakeBookmarkRegistry()
    for bm_name, (s, e) in (bookmarks or {}).items():
        bm_registry.add(bm_name, s, e)
    doc.Bookmarks = bm_registry

    # Cross-reference items: Word returns ordered strings per reference type —
    # bookmark *names* (type 2) and heading *texts* (type 1) — which
    # `insert_cross_reference` indexes into. Footnote/endnote types use the
    # 1-based index directly, so they're not needed here.
    def _xref_items(ref_type: Any) -> tuple[str, ...]:
        rt = int(ref_type)
        if rt == 2:  # wdRefTypeBookmark
            return tuple((bookmarks or {}).keys())
        if rt == 1:  # wdRefTypeHeading
            return tuple(p["text"] for p in (paragraphs or []) if p.get("level", 10) < 10)
        return ()

    doc.GetCrossReferenceItems = MagicMock(side_effect=_xref_items)

    doc.ContentControls = _FakeContentControls(content_controls or [])
    doc.Paragraphs = _FakeParagraphs(paragraphs or [])
    doc.Styles = _FakeStyles(styles if styles is not None else _DEFAULT_STYLES)
    doc.Tables = _FakeTablesCollection(
        [_FakeTable(t["grid"], t.get("title", ""), style=t.get("style")) for t in (tables or [])]
    )
    doc.Comments = _FakeComments()
    doc.Revisions = _FakeRevisions(revisions or [])

    # Track Changes flag: a plain settable bool so doc.track_changes /
    # tracked_changes() round-trip through the same handle.
    doc.TrackRevisions = False

    # `Range(start, end)` returns a range whose `.Text` can be set/read by the
    # caller. Cache by (start, end) so a write through one handle is visible on
    # the next lookup of the same span — needed for RangeAnchor round-trips.
    _range_cache: dict[tuple[int, int], MagicMock] = {}

    def _range_factory(start: int, end: int) -> MagicMock:
        key = (int(start), int(end))
        rng = _range_cache.get(key)
        if rng is None:
            rng = _make_range(start, end)
            # Wire the range's InlineShapes so a ConvertToShape() (floating
            # insert_image) lands the picture in this document's body Shapes.
            rng.InlineShapes._doc_shapes = doc.Shapes
            rng.InlineShapes._anchor_start = int(start)
            _range_cache[key] = rng
        return rng

    doc.Range.side_effect = _range_factory

    # `Content` is the full document range; find/replace reads `.Text` from it.
    content_range = _make_range(0, len(content))
    content_range.Text = content
    doc.Content = content_range

    # Lists: build each list's Range through the cached factory so list_info /
    # restart_numbering see the same ListFormat the RangeAnchor resolves to.
    list_objs: list[_FakeWordList] = []
    for spec in lists or []:
        rng = _range_factory(spec["start"], spec["end"])
        lf = rng.ListFormat
        lf.ListType = spec.get("type", 3)
        lf.ListValue = 1 if lf.ListType in (3, 4) else 0
        lf.ListString = {2: "•", 3: "1.", 4: "1."}.get(lf.ListType, "")
        lf._gallery = {2: 1, 3: 2, 4: 3}.get(lf.ListType)
        list_objs.append(_FakeWordList(rng, spec.get("count", 0)))
    doc.Lists = _FakeLists(list_objs)
    # Custom list-template authoring (apply_list_format) mints templates here.
    doc.ListTemplates = _FakeListTemplates()

    # Sections: at least one (real documents always have one).
    section_specs = sections if sections is not None else [{}]
    section_objs: list[_FakeSection] = []
    for i, spec in enumerate(section_specs, start=1):
        section_objs.append(
            _FakeSection(
                i,
                _build_hf_map(spec.get("headers", {})),
                _build_hf_map(spec.get("footers", {})),
                _FakePageSetup(**spec.get("page_setup", {})),
            )
        )
    doc.Sections = _FakeSections(section_objs)

    # Footnotes / endnotes: seed each note's reference at the given offset (so
    # `list()` can map it back to a paragraph). Empty by default.
    def _seed_notes(specs: list[dict[str, Any]] | None) -> _FakeNotes:
        notes = _FakeNotes()
        for spec in specs or []:
            ref_start = int(spec.get("ref_start", 0))
            notes.Add(_range_factory(ref_start, ref_start), "", spec.get("text", ""))
        return notes

    doc.Footnotes = _seed_notes(footnotes)
    doc.Endnotes = _seed_notes(endnotes)
    doc.TablesOfContents = _FakeTOCs()
    doc.Indexes = _FakeIndexes()
    doc.TablesOfFigures = _FakeTablesOfFigures()
    doc.TablesOfAuthorities = _FakeTablesOfAuthorities()
    doc.Bibliography = _FakeBibliography()
    doc.DocumentTheme = _FakeOfficeTheme()
    doc.ApplyDocumentTheme = MagicMock(name="ApplyDocumentTheme")
    # The theme code derives the built-in library dir from app.Path/Version;
    # give them real strings (the derived dir won't exist — built-in name
    # resolution tests monkeypatch wordlive._themes._themes_dir).
    doc.Application.Path = r"C:\Office\root\Office16"
    doc.Application.Version = "16.0"

    # Equations: document-level OMaths, each range built through the cached
    # factory so its Start maps back to a paragraph for list()'s `para`.
    doc.OMaths = _FakeDocOMaths(equations, _range_factory)

    # Charts: AddChart2 (off the Selection) inserts into doc.InlineShapes.
    _wire_chart_insertion(doc)

    return doc


def _make_application(documents: list[MagicMock]) -> MagicMock:
    app = MagicMock(name="Word.Application")
    app.Visible = True

    # Documents collection: indexable + iterable + Count + ActiveDocument
    docs_collection = MagicMock(name="Documents")
    docs_collection.Count = len(documents)
    docs_collection.__iter__ = MagicMock(side_effect=lambda: iter(documents))
    app.Documents = docs_collection
    app.ActiveDocument = documents[0] if documents else MagicMock(name="NoActiveDoc")

    # Selection / view
    app.Selection.Start = 0
    app.Selection.End = 0
    app.Selection.Text = ""
    app.ActiveWindow.VerticalPercentScrolled = 0

    # UndoRecord
    app.UndoRecord.StartCustomRecord = MagicMock(name="StartCustomRecord")
    app.UndoRecord.EndCustomRecord = MagicMock(name="EndCustomRecord")
    return app


@pytest.fixture
def fake_word(monkeypatch: pytest.MonkeyPatch) -> MagicMock:
    """A MagicMock Application with one document, one bookmark, one heading, one CC."""
    doc = _make_document(
        bookmarks={"Address": (13, 24)},
        content_controls=[
            {"title": "Signatory", "tag": "sig", "start": 29, "end": 35, "text": "Jane Doe"}
        ],
        paragraphs=[
            {"level": 1, "text": "Introduction", "start": 0, "end": 13},
            {"level": 10, "text": "Body text here.", "start": 13, "end": 29},
            {"level": 2, "text": "Risks", "start": 29, "end": 35},
        ],
        # Self-consistent text — paragraph offsets above index into this string.
        content="Introduction\rBody text here.\rRisks\r",
        tables=[{"grid": [["A1", "B1"], ["A2", "B2"]], "title": "Grid"}],
        # A 2-item numbered list over the body region (offsets 13–29).
        lists=[{"start": 13, "end": 29, "count": 2, "type": 3}],
        # One section with a primary header and footer seeded for read tests.
        sections=[{"headers": {"primary": "Confidential Draft"}, "footers": {"primary": "Page 1"}}],
        # One footnote whose reference mark sits in the body paragraph (13–29),
        # so footnotes.list() maps it back to para:2.
        footnotes=[{"ref_start": 20, "text": "A seeded footnote."}],
        # One tracked insertion so revisions.list() / the reader have a target.
        revisions=[{"type": 1, "author": "Reviewer", "start": 13, "end": 18, "text": "Body"}],
        # One embedded PNG anchored in the body paragraph (13–29), so
        # images.list() maps it back to para:2 and read_image() round-trips.
        images=[
            {
                "mime": "image/png",
                "data": b"\x89PNG\r\n\x1a\nSEEDED",
                "alt_text": "logo",
                "start": 20,
            }
        ],
        # One display equation whose range sits in the body paragraph (13–29),
        # so equations.list() maps it back to para:2. Text is the built-up form
        # (with the internal CR markers the linear preview strips).
        equations=[{"start": 21, "end": 27, "type": 1, "text": "𝐸\r=\r𝑚\r𝑐\r2\r"}],
        # One document variable for the variables read/write tests.
        variables={"ClientName": "Acme"},
        # One hyperlink in the body paragraph (13–29) -> external URL.
        hyperlinks=[{"text": "Acme", "address": "https://acme.example", "start": 15, "end": 19}],
        # One PAGE field whose code range sits in the body paragraph (13–29).
        fields=[{"code": "PAGE", "result": "1", "type": 33, "start": 16, "end": 17}],
        # Built-in + custom document properties; "Word count" is a read-only stat
        # and "Last print date" is unset (its value access raises and is skipped).
        builtin_properties={
            "Title": "Quarterly Report",
            "Author": "Jane Doe",
            "Word count": 1234,
            "Last print date": None,
        },
        readonly_properties={"Word count"},
        unreadable_properties={"Last print date"},
        custom_properties={"Project": "Apollo"},
        # Proofing: one spelling error, one grammar error, and a readability stat.
        spelling=[{"text": "teh", "start": 14, "end": 17}],
        grammar=[{"text": "is are", "start": 20, "end": 26}],
        readability={
            "Flesch Reading Ease": 65.5,
            "Flesch-Kincaid Grade Level": 7.2,
            "Passive Sentences": 12.0,
            "Words per Sentence": 15.3,
        },
    )
    app = _make_application([doc])

    from wordlive import _com

    monkeypatch.setattr(_com, "get_active_word", lambda: app)
    monkeypatch.setattr(_com, "launch_word", lambda visible=True: app)
    return app


@pytest.fixture
def no_word(monkeypatch: pytest.MonkeyPatch) -> None:
    """Simulate Word not running."""
    from wordlive import _com
    from wordlive.exceptions import WordNotRunningError

    def _raise(*args: Any, **kwargs: Any) -> None:
        raise WordNotRunningError("no running Microsoft Word instance found")

    monkeypatch.setattr(_com, "get_active_word", _raise)
    monkeypatch.setattr(_com, "launch_word", _raise)


@pytest.fixture
def real_word():
    """Smoke fixture: yields a wordlive.Word, or skips if Word isn't reachable."""
    import wordlive
    from wordlive.exceptions import WordNotRunningError

    try:
        ctx = wordlive.attach()
        word = ctx.__enter__()
    except WordNotRunningError as e:
        pytest.skip(f"Word not running: {e}")
        return
    try:
        yield word
    finally:
        ctx.__exit__(None, None, None)
