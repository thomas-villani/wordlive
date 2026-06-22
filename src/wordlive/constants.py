"""Typed enums for the Word magic constants wordlive uses.

Values mirror the official `Wd*` enumerations exactly. Resist the urge to
pre-populate — add entries only as v0 features need them.
"""

from __future__ import annotations

from enum import IntEnum

# Word's `wdUndefined` sentinel — returned by `Font.Size`/`Font.Bold`/etc. when a
# property varies across the runs of a range (i.e. mixed formatting). The format
# read mirror (`Anchor.format_info`) special-cases it into a `mixed` list rather
# than reporting a bogus 9999999.
WD_UNDEFINED = 9999999


class WdParagraphAlignment(IntEnum):
    LEFT = 0
    CENTER = 1
    RIGHT = 2
    JUSTIFY = 3


class WdGoToItem(IntEnum):
    LINE = 3
    PAGE = 1
    BOOKMARK = -1


class WdGoToDirection(IntEnum):
    FIRST = 1
    LAST = -1
    NEXT = 2
    PREVIOUS = 3
    ABSOLUTE = 1


class WdUnits(IntEnum):
    CHARACTER = 1
    LINE = 5
    STORY = 6


class WdCollapseDirection(IntEnum):
    START = 1
    END = 0


class WdStyleType(IntEnum):
    """Word's `Style.Type` values.

    The CLI / `Style.to_dict()` emits these as lowercase strings —
    `WdStyleType.PARAGRAPH` round-trips to `"paragraph"`, etc. — so JSON
    consumers can match on the human-readable form without needing to import
    the enum. If you're filtering in Python, compare against the IntEnum
    member directly.
    """

    PARAGRAPH = 1
    CHARACTER = 2
    TABLE = 3
    LIST = 4


class WdListGalleryType(IntEnum):
    """Which `Application.ListGalleries` collection a list template comes from.

    `apply_list("bulleted"|"numbered"|"outline")` maps onto these.
    """

    BULLET = 1
    NUMBER = 2
    OUTLINE_NUMBER = 3


class WdListType(IntEnum):
    """Word's `ListFormat.ListType` values — what kind of list a range is in.

    `list_info()["type"]` emits these as lowercase strings (`BULLET` ->
    `"bulleted"`, `SIMPLE_NUMBERING` -> `"numbered"`, etc.); `NO_NUMBERING`
    is reported as `"none"`.
    """

    NO_NUMBERING = 0
    LIST_NUM_ONLY = 1
    BULLET = 2
    SIMPLE_NUMBERING = 3
    OUTLINE_NUMBERING = 4
    MIXED_NUMBERING = 5


class WdListApplyTo(IntEnum):
    """Scope argument for `ListFormat.ApplyListTemplate`."""

    WHOLE_LIST = 0
    THIS_POINT_FORWARD = 1
    SELECTION = 2


class WdNumberType(IntEnum):
    """`NumberType` argument for `ListFormat.RemoveNumbers`."""

    PARAGRAPH = 1
    LIST_NUM = 2
    ALL_NUMBERS = 3


class WdDefaultListBehavior(IntEnum):
    """`DefaultListBehavior` argument for `ListFormat.ApplyListTemplate`.

    `WORD10` enables the modern multi-level numbering behaviour and is what
    `apply_list` passes.
    """

    WORD8 = 0
    WORD9 = 1
    WORD10 = 2


class WdListNumberStyle(IntEnum):
    """`ListLevel.NumberStyle` values — the numbering scheme for one list level.

    The common subset `apply_list_format` exposes by name. `BULLET` (23) is a
    trap: setting it on a multi-level template level raises (`0x800a1200`,
    live-probed 2026-06-21), so a bullet level is authored by setting
    `NumberFormat` to the glyph plus a `Symbol`/`Wingdings` font instead.
    """

    ARABIC = 0
    UPPERCASE_ROMAN = 1
    LOWERCASE_ROMAN = 2
    UPPERCASE_LETTER = 3
    LOWERCASE_LETTER = 4
    ORDINAL = 5
    CARDINAL_TEXT = 6
    ORDINAL_TEXT = 7
    BULLET = 23


class WdTrailingCharacter(IntEnum):
    """`ListLevel.TrailingCharacter` — what follows a list level's number/bullet."""

    TAB = 0
    SPACE = 1
    NONE = 2


class WdListLevelAlignment(IntEnum):
    """`ListLevel.Alignment` — how a list level's number/bullet aligns."""

    LEFT = 0
    CENTER = 1
    RIGHT = 2


class WdHeaderFooterIndex(IntEnum):
    """Which header/footer of a section — `Section.Headers(index)`.

    `header()` / `footer()` accept the string aliases (`"primary"`,
    `"first"`, `"even"`) that map onto these.
    """

    PRIMARY = 1
    FIRST_PAGE = 2
    EVEN_PAGES = 3


class WdOrientation(IntEnum):
    """Page orientation — `PageSetup.Orientation`."""

    PORTRAIT = 0
    LANDSCAPE = 1


class WdWrapType(IntEnum):
    """`Shape.WrapFormat.Type` — how text wraps around a floating image.

    `insert_image(wrap=...)` maps its string values onto these. `INLINE` is
    never set through `WrapFormat`: an inline image stays an `InlineShape` and
    is never converted to a floating `Shape`.
    """

    SQUARE = 0
    TIGHT = 1
    THROUGH = 2
    TOP_BOTTOM = 3
    FRONT = 4
    BEHIND = 5
    INLINE = 7


class MsoTriState(IntEnum):
    """Office's tri-state boolean — used here for `InlineShape.LockAspectRatio`."""

    TRUE = -1
    FALSE = 0


class MsoTextOrientation(IntEnum):
    """`Shapes.AddTextbox(Orientation=...)` — only the horizontal box is exposed."""

    HORIZONTAL = 1


class MsoShapeType(IntEnum):
    """`Shape.Type` — discriminates the floating shapes `shape:N` addresses.

    A narrow slice of Office's full `MsoShapeType`: the kinds a wordlive document
    actually grows. `TEXT_BOX` is what `insert_text_box` adds; `PICTURE` is what a
    floating `insert_image` (`InlineShape.ConvertToShape`) becomes; `TEXT_EFFECT`
    is WordArt (e.g. a watermark, though those live in the header story and so are
    excluded from the body-shape collection); `GROUP` / `AUTO_SHAPE` round out the
    common cases. `_shapes.shape_kind` maps these onto the public `shape_type`
    string and falls back to ``"other"`` for anything unmapped.
    """

    AUTO_SHAPE = 1
    GROUP = 6
    PICTURE = 13
    TEXT_EFFECT = 15
    TEXT_BOX = 17


class MsoZOrderCmd(IntEnum):
    """`Shape.ZOrder(...)` commands — restack a floating shape in the float layer.

    The four canonical reorder verbs `ShapeAnchor.set_z_order` exposes
    (``"front"`` / ``"back"`` / ``"forward"`` / ``"backward"``). The
    in-front-of-text / behind-text commands are a wrap concern (`set_wrap`'s
    ``"front"`` / ``"behind"``), not a z-order one, so they're left out here.
    """

    BRING_TO_FRONT = 0
    SEND_TO_BACK = 1
    BRING_FORWARD = 2
    SEND_BACKWARD = 3


class MsoPresetTextEffect(IntEnum):
    """`Shapes.AddTextEffect(PresetTextEffect=...)` — the plain WordArt preset.

    `TEXT_EFFECT1` (`msoTextEffect1`) is the unstyled, fill-only preset Word's own
    text-watermark feature uses; the fill colour / transparency are then set on
    the returned shape.
    """

    TEXT_EFFECT1 = 0


class WdShapePosition(IntEnum):
    """Sentinel positions for `Shape.Left` / `Shape.Top` (centre on the relative-to frame)."""

    CENTER = -999995


class WdRelativeHorizontalPosition(IntEnum):
    """`Shape.RelativeHorizontalPosition` — what `Shape.Left` is measured from.

    A narrow slice: the page margin (default) and the page edge — the two frames
    `set_position` exposes for a floating shape's horizontal offset.
    """

    MARGIN = 0
    PAGE = 1


class WdRelativeVerticalPosition(IntEnum):
    """`Shape.RelativeVerticalPosition` — what `Shape.Top` is measured from.

    A narrow slice: the page margin (default) and the page edge — the two frames
    `set_position` exposes for a floating shape's vertical offset.
    """

    MARGIN = 0
    PAGE = 1


class WdWrapSideType(IntEnum):
    """`Shape.WrapFormat.Side` — which sides of a floating shape text flows past.

    Only meaningful for the `square` / `tight` / `through` wraps; for
    `top-bottom` / `front` / `behind` Word silently coerces `Side` back to
    `BOTH` (live-probed 2026-06-19). `LARGEST` lets text flow only past the
    wider side.
    """

    BOTH = 0
    LEFT = 1
    RIGHT = 2
    LARGEST = 3


class WdStoryType(IntEnum):
    """`Range.StoryType` / `Shape.Anchor.StoryType` — which story a range lives in.

    Only the main text story is named: it's the guard `_shapes.body_shapes` uses
    to keep header/footer-anchored shapes (watermarks) out of the body-shape
    collection, even though `Document.Shapes` already separates the stories.
    """

    MAIN_TEXT = 1


class WdSaveFormat(IntEnum):
    """`Document.SaveAs2(FileFormat=...)` values — the formats `save_as` exposes.

    A deliberately narrow slice of Word's full `WdSaveFormat`: the modern Open
    XML `.docx` (`DOCUMENT_DEFAULT`, what `save_as(fmt="docx")` writes). PDF
    export goes through `ExportAsFixedFormat` (`export_pdf`), not `SaveAs2`, so
    PDF isn't listed here. The legacy `.doc`, `.rtf`, `.txt`, and `.html`
    formats are deferred until a use case needs them, mirroring how the other
    `Wd*` subsets were kept narrow.
    """

    DOCUMENT_DEFAULT = 16


class WdExportFormat(IntEnum):
    """`Document.ExportAsFixedFormat` output format. Only PDF is used (for snapshots)."""

    PDF = 17


class WdExportRange(IntEnum):
    """Which pages `ExportAsFixedFormat` writes.

    `FROM_TO` pairs with the `From`/`To` (1-based, inclusive) page numbers;
    `ALL_DOCUMENT` ignores them and exports everything.
    """

    ALL_DOCUMENT = 0
    FROM_TO = 3


class WdExportItem(IntEnum):
    """What `ExportAsFixedFormat(Item=...)` includes in the exported PDF.

    `DOCUMENT_CONTENT` (the default) renders the final text only; `WITH_MARKUP`
    renders tracked changes and comments as visible revision marks / balloons —
    the `markup="all"` snapshot mode. No view mutation is needed: the export
    parameter decides, so the user's on-screen markup setting is left untouched.
    """

    DOCUMENT_CONTENT = 0
    WITH_MARKUP = 7


class WdRevisionType(IntEnum):
    """`Revision.Type` values — the tracked-change kinds `doc.revisions` reports.

    A subset of Word's full `WdRevisionType`: the common insert / delete /
    property (formatting) revisions plus paragraph-number / property changes.
    `RevisionCollection` maps each int onto a human string (`"insert"`,
    `"delete"`, `"format"`, …); an unrecognised value reports as `"other"`.
    """

    NO_REVISION = 0
    INSERT = 1
    DELETE = 2
    PROPERTY = 3
    PARAGRAPH_NUMBER = 4
    DISPLAY_FIELD = 5
    RECONCILE = 6
    CONFLICT = 7
    STYLE = 8
    REPLACE = 9
    PARAGRAPH_PROPERTY = 10
    TABLE_PROPERTY = 11
    SECTION_PROPERTY = 12
    STYLE_DEFINITION = 13
    MOVE_SOURCE = 14
    MOVE_TARGET = 15
    CELL_INSERTION = 16
    CELL_DELETION = 17
    CELL_MERGE = 18


class WdBreakType(IntEnum):
    """`Range.InsertBreak(Type=...)` values — the break kinds `insert_break` exposes.

    A deliberate subset of Word's full `WdBreakType`: the page/column break and
    the two section breaks the v0.12 plan calls for. The line/text-wrapping
    breaks (6, 9–11) and even/odd-page section breaks (4, 5) are omitted until a
    use case needs them, mirroring how `WdStyleType` was kept narrow.
    `insert_break(kind=...)` maps its string keys onto these members.
    """

    SECTION_NEXT_PAGE = 2
    SECTION_CONTINUOUS = 3
    PAGE = 7
    COLUMN = 8


class WdUnderline(IntEnum):
    """`Font.Underline` values — the subset `format_run(underline=...)` uses.

    `underline=True` maps to `SINGLE`, `False` to `NONE`. Word's full
    `WdUnderline` has ~17 members (double, dotted, wavy, …); they're omitted
    until a use case needs them, mirroring how `WdBreakType` was kept narrow.
    """

    NONE = 0
    SINGLE = 1


class WdColorIndex(IntEnum):
    """`Range.HighlightColorIndex` values — the named text-highlight colours.

    Highlight is a fixed *palette index*, not an arbitrary RGB, so it has its own
    enum (the colour helper, which yields a BGR long, is wrong for it).
    `format_run(highlight=...)` maps its string keys onto these members.
    """

    AUTO = 0
    BLACK = 1
    BLUE = 2
    TURQUOISE = 3
    BRIGHT_GREEN = 4
    PINK = 5
    RED = 6
    YELLOW = 7
    WHITE = 8
    DARK_BLUE = 9
    TEAL = 10
    GREEN = 11
    VIOLET = 12
    DARK_RED = 13
    DARK_YELLOW = 14
    GRAY_50 = 15
    GRAY_25 = 16


class WdLineStyle(IntEnum):
    """`Border.LineStyle` values — the subset `set_borders(style=...)` exposes.

    A deliberate slice of Word's full `WdLineStyle`; the decorative/art styles
    (waves, 3-D, multi-colour) are omitted until needed. `NONE` removes a border.
    """

    NONE = 0
    SINGLE = 1
    DOT = 2
    DASH_SMALL_GAP = 3
    DASH_LARGE_GAP = 4
    DASH_DOT = 5
    DASH_DOT_DOT = 6
    DOUBLE = 7


class WdLineSpacing(IntEnum):
    """`ParagraphFormat.LineSpacingRule` values — `format_paragraph(line_spacing=…)`.

    `SINGLE`/`ONE_POINT_FIVE`/`DOUBLE` are the named multiples (no companion
    `LineSpacing` value needed). `MULTIPLE` carries an arbitrary multiple of
    single spacing (its `LineSpacing` is the multiple × 12pt). `AT_LEAST` and
    `EXACTLY` set a fixed minimum / exact line height in points.
    """

    SINGLE = 0
    ONE_POINT_FIVE = 1
    DOUBLE = 2
    AT_LEAST = 3
    EXACTLY = 4
    MULTIPLE = 5


class WdDropPosition(IntEnum):
    """`DropCap.Position` — where `drop_cap(position=...)` puts the dropped letter.

    `NONE` removes an existing drop cap; `DROPPED` sets it into the body text;
    `MARGIN` hangs it out in the left margin.
    """

    NONE = 0
    DROPPED = 1
    MARGIN = 2


class WdTabAlignment(IntEnum):
    """`TabStop.Alignment` — `add_tab_stop(align=...)` maps its keys onto these."""

    LEFT = 0
    CENTER = 1
    RIGHT = 2
    DECIMAL = 3
    BAR = 4


class WdTabLeader(IntEnum):
    """`TabStop.Leader` — the dot/dash leader for `add_tab_stop(leader=...)`.

    `SPACES` (the default) means no visible leader.
    """

    SPACES = 0
    DOTS = 1
    DASHES = 2
    LINES = 3
    HEAVY = 4
    MIDDLE_DOT = 5


class WdBorderType(IntEnum):
    """`Borders(index)` selectors — which edge of a range/cell a border is on.

    `set_borders(sides=...)` maps `"top"`/`"bottom"`/`"left"`/`"right"` onto the
    four outer edges; `"all"`/`"box"` applies to all four. The interior
    horizontal/vertical gridlines (only meaningful for a multi-cell/table range)
    are included for completeness.
    """

    TOP = -1
    LEFT = -2
    BOTTOM = -3
    RIGHT = -4
    HORIZONTAL = -5
    VERTICAL = -6


class WdInformation(IntEnum):
    """`Range.Information(...)` selectors.

    `ACTIVE_END_PAGE_NUMBER` returns the 1-based page a (collapsed) range falls
    on, which is how an anchor is mapped to the page(s) a snapshot should render.
    `WITH_IN_TABLE` is True when the range sits inside a table — used when
    inserting a new table to detect (and separate) an adjacent table that Word
    would otherwise silently merge into.

    The line/column selectors report the laid-out position of the range's *first*
    character (`anchor.location()`); `ACTIVE_END_ADJUSTED_PAGE_NUMBER` is the
    page number as the user would read it (honouring per-section page-number
    restarts), kept alongside the raw `ACTIVE_END_PAGE_NUMBER`. All of these are
    print-layout reads, so `location()` repaginates first.
    """

    ACTIVE_END_ADJUSTED_PAGE_NUMBER = 1
    ACTIVE_END_PAGE_NUMBER = 3
    FIRST_CHARACTER_COLUMN_NUMBER = 9
    FIRST_CHARACTER_LINE_NUMBER = 10
    WITH_IN_TABLE = 12


class WdStatistic(IntEnum):
    """`Document.ComputeStatistics(...)` selectors — Word's own counters.

    The narrow subset `Document.stats()` surfaces: the counts Word computes
    itself (pages/words/characters/lines/paragraphs). The structural counts
    (tables, images, comments, …) come from wordlive's own collections instead,
    so they're not mirrored here. `wdStatisticCharactersWithSpaces` (5) and
    `wdStatisticFarEastCharacters` (6) exist but aren't exposed yet. Page/line
    counts are print-layout reads, so `stats()` repaginates first.
    """

    WORDS = 0
    LINES = 1
    PAGES = 2
    CHARACTERS = 3
    PARAGRAPHS = 4


class WdPaperSize(IntEnum):
    """`PageSetup.PaperSize` values — the subset `set_page_setup(paper_size=...)` exposes.

    A deliberate slice of Word's full `WdPaperSize` (the common US and ISO sizes);
    the dozens of envelope/legacy sizes are omitted until needed, mirroring how
    `WdBreakType` was kept narrow. Setting `PaperSize` adjusts the page's width
    and height to match.
    """

    LETTER = 2
    TABLOID = 3
    LEGAL = 4
    A3 = 6
    A4 = 7
    A5 = 8


class WdReferenceType(IntEnum):
    """`InsertCrossReference(ReferenceType=...)` / `GetCrossReferenceItems(...)` values.

    The reference *category* a cross-reference points at. `insert_cross_reference`
    maps an anchor-id target onto these: `bookmark:NAME` → `BOOKMARK`,
    `heading:N` → `HEADING`, `footnote:N` → `FOOTNOTE`, `endnote:N` → `ENDNOTE`.
    `NUMBERED_ITEM` (cross-refs to a numbered-list item) is included for
    completeness but not yet exposed.
    """

    NUMBERED_ITEM = 0
    HEADING = 1
    BOOKMARK = 2
    FOOTNOTE = 3
    ENDNOTE = 4


class WdReferenceKind(IntEnum):
    """`InsertCrossReference(ReferenceKind=...)` values — *what* a cross-ref inserts.

    A deliberately narrow slice of Word's full `WdReferenceKind`.
    `insert_cross_reference(kind=...)` maps its string keys onto these:
    `"text"` → `CONTENT_TEXT` (the heading/bookmark text), `"page"` →
    `PAGE_NUMBER`, `"above_below"` → `POSITION` ("above"/"below"), and
    `"number"` → `NUMBER_NO_CONTEXT` for headings/bookmarks or
    `FOOTNOTE_NUMBER`/`ENDNOTE_NUMBER` for notes.
    """

    CONTENT_TEXT = -1
    NUMBER_NO_CONTEXT = -3
    FOOTNOTE_NUMBER = 5
    ENDNOTE_NUMBER = 6
    PAGE_NUMBER = 7
    POSITION = 15


class WdCaptionPosition(IntEnum):
    """`InsertCaption(Position=...)` — whether a caption goes above or below.

    `insert_caption` picks by convention when no position is given: `ABOVE` for
    a ``"Table"`` label, `BELOW` for figures and everything else.
    """

    ABOVE = 0
    BELOW = 1


class WdAutoFitBehavior(IntEnum):
    """`Table.AutoFitBehavior(Behavior=...)` values — how a table sizes itself.

    `Table.autofit(mode=...)` maps its string keys onto these: ``"fixed"`` pins
    the current column widths (no auto-resize), ``"content"`` shrinks/grows each
    column to fit its cell contents, and ``"window"`` stretches the table to the
    page/container width. `FIXED` also requires `Table.AllowAutoFit = False`,
    which `autofit` sets for that mode.
    """

    FIXED = 0
    CONTENT = 1
    WINDOW = 2


class WdRowAlignment(IntEnum):
    """`Table.Rows.Alignment` values — how a table sits across the page width.

    `Table.set_alignment(...)` maps ``"left"``/``"center"``/``"right"`` onto
    these. Distinct from `WdParagraphAlignment` (which aligns text *within* a
    paragraph) — this aligns the *whole table* between the page margins, so it
    has no `JUSTIFY`.
    """

    LEFT = 0
    CENTER = 1
    RIGHT = 2


class WdCellVerticalAlignment(IntEnum):
    """`Cell.VerticalAlignment` values — where a cell's text sits vertically.

    `Cell.set_vertical_alignment(...)` maps ``"top"``/``"center"``/``"bottom"``
    onto these. Note the **gap at 2**: that slot is `wdAlignVerticalJustify` in
    Word's shared vertical-alignment value space, which a *cell* rejects (live
    Word raises ``0x80020009`` for `VerticalAlignment = 2`), so only 0/1/3 are
    valid here.
    """

    TOP = 0
    CENTER = 1
    BOTTOM = 3


class MsoDocProperty(IntEnum):
    """`Office.DocumentProperty.Type` values — used when adding a custom property.

    `PropertyCollection.set(..., custom=True)` infers the type from the Python
    value: `bool` -> `BOOLEAN`, `int` -> `NUMBER`, `float` -> `FLOAT`, and
    everything else -> `STRING` (the safe default Word accepts for arbitrary
    text). `DATE` exists for completeness but isn't inferred (wordlive has no
    first-class date value to map onto it).
    """

    NUMBER = 1
    BOOLEAN = 2
    DATE = 3
    STRING = 4
    FLOAT = 5


class WdFieldType(IntEnum):
    """`Fields.Add(Type=...)` values — the field kinds `insert_field(kind=...)` exposes.

    A narrow, agent-useful slice of Word's ~100 field types. `EMPTY` is the
    raw-code escape hatch: `insert_field("field", text="REF foo \\\\h")` inserts an
    arbitrary field whose code is carried in the text, so REF/TOC/etc. are
    reachable without a member each. The named members cover the publishing
    staples — page numbers, counts, date/time, document metadata — plus the
    citation/authority field kinds (`CITATION`, `BIBLIOGRAPHY`, `TOA`,
    `TOA_ENTRY`) that wordlive inserts via the `EMPTY` raw-code path and reads
    back to identify (these four numerics were confirmed against live Word — the
    "obvious" guesses 119/34/73-for-the-entry were all wrong).
    """

    EMPTY = -1
    AUTHOR = 17
    NUM_PAGES = 26
    FILE_NAME = 29
    DATE = 31
    TIME = 32
    PAGE = 33
    TITLE = 49
    TOA = 73  # table of authorities (the field the table itself renders as)
    TOA_ENTRY = 74  # a TA mark — one entry feeding the table of authorities
    CITATION = 96
    BIBLIOGRAPHY = 97


class WdContentControlType(IntEnum):
    """`ContentControls.Add(Type=...)` values — the control kinds `insert_content_control` exposes.

    `insert_content_control(kind=...)` maps its string keys onto these:
    ``"rich_text"`` → `RICH_TEXT` (the default — formatted text), ``"text"`` →
    `TEXT` (plain text), ``"picture"`` → `PICTURE`, ``"combo_box"`` →
    `COMBO_BOX` (a dropdown that also accepts typed text), ``"dropdown"`` →
    `DROPDOWN_LIST`, ``"date"`` → `DATE`, ``"building_block"`` →
    `BUILDING_BLOCK_GALLERY`, ``"group"`` → `GROUP`, ``"checkbox"`` →
    `CHECKBOX` (Word 2013+), and ``"repeating_section"`` → `REPEATING_SECTION`
    (Word 2013+). The form-building primitive — pair `combo_box`/`dropdown`
    with a list of items.
    """

    RICH_TEXT = 0
    TEXT = 1
    PICTURE = 2
    COMBO_BOX = 3
    DROPDOWN_LIST = 4
    BUILDING_BLOCK_GALLERY = 5
    DATE = 6
    GROUP = 7
    CHECKBOX = 8
    REPEATING_SECTION = 9


class WdIndexType(IntEnum):
    """`Indexes.Add(Type=...)` values — how index subentries are laid out.

    `insert_index(run_in=...)` maps onto these: ``False`` → `INDENT`
    (subentries indented under the main entry, one per line — the common book
    style) and ``True`` → `RUNIN` (subentries run together in a single
    paragraph).
    """

    INDENT = 0
    RUNIN = 1


class XlChartType(IntEnum):
    """`Range.InlineShapes.AddChart2(Type=...)` values — the chart kinds `insert_chart` exposes.

    A deliberately narrow slice of Excel's ~80-member `XlChartType` (charts are
    Excel-backed: `AddChart2` embeds a chart whose data lives in a hidden Excel
    workbook). `insert_chart(kind=...)` maps its string keys onto these:
    ``"bar"`` → `COLUMN_CLUSTERED` (vertical clustered columns — Word's own
    "bar"-button default), ``"pie"`` → `PIE`, ``"line"`` → `LINE`, and
    ``"scatter"`` → `XY_SCATTER_MARKERS` (markers only, no connecting line — the
    scientific default; both axes numeric). Multi-series, secondary axes, and
    axis/series formatting are deferred — keep this narrow.
    """

    COLUMN_CLUSTERED = 51
    LINE = 4
    PIE = 5
    XY_SCATTER_MARKERS = 65


class XlLegendPosition(IntEnum):
    """`Chart.Legend.Position` values — where the legend sits relative to the plot.

    Used by `ChartAnchor.format(legend_position=...)`. `CORNER` is the top-right
    overlay; the rest are outside the plot on that edge.
    """

    BOTTOM = -4107
    CORNER = 2
    LEFT = -4131
    RIGHT = -4152
    TOP = -4160


class XlAxisType(IntEnum):
    """`Chart.Axes(Type, ...)` selector. `CATEGORY` is the x/horizontal axis,
    `VALUE` the y/value axis. `ChartAnchor.set_axis(which=...)` maps
    ``"category"``/``"x"`` → `CATEGORY` and ``"value"``/``"y"`` → `VALUE`."""

    CATEGORY = 1
    VALUE = 2


class XlAxisGroup(IntEnum):
    """`Chart.Axes(Type, AxisGroup)` group selector. wordlive targets the
    `PRIMARY` axis group only (secondary axes are deferred)."""

    PRIMARY = 1


class XlScaleType(IntEnum):
    """`Axis.ScaleType` values. `LOGARITHMIC` is the order-of-magnitude scale;
    `ChartAnchor.set_axis(scale=...)` maps ``"linear"`` / ``"log"`` onto these."""

    LINEAR = -4132
    LOGARITHMIC = -4133


class XlTrendlineType(IntEnum):
    """`Trendlines.Add(Type=...)` values — the curve fitted to a series.

    `ChartAnchor.add_trendline(kind=...)` maps ``"linear"``, ``"exponential"``,
    ``"logarithmic"``, ``"moving_average"``, ``"polynomial"``, ``"power"`` onto
    these. A power/exponential fit with `DisplayEquation` draws the law of best
    fit on a scatter — first-class for scientific publishing.
    """

    LINEAR = -4132
    EXPONENTIAL = 5
    LOGARITHMIC = -4133
    MOVING_AVERAGE = 6
    POLYNOMIAL = 3
    POWER = 4


class XlMarkerStyle(IntEnum):
    """`Series.MarkerStyle` / `Point.MarkerStyle` values — the data-point glyph.

    Used by `ChartAnchor.format_series(marker=...)` for line/scatter series.
    `AUTOMATIC` lets Word pick; `NONE` hides markers. The rest name the glyph.
    """

    NONE = -4142
    AUTOMATIC = -4105
    SQUARE = 1
    DIAMOND = 2
    TRIANGLE = 3
    X = -4168
    STAR = 5
    DOT = -4118
    DASH = -4115
    CIRCLE = 8
    PLUS = 9


class XlErrorBarDirection(IntEnum):
    """`Series.ErrorBar(Direction=...)` — which axis the error bars extend along.

    wordlive draws value-axis (`Y`) error bars; `X` is exposed for scatter where
    the independent variable also carries uncertainty.
    """

    X = -4168
    Y = 1


class XlErrorBarInclude(IntEnum):
    """`Series.ErrorBar(Include=...)` — which side(s) of each point to draw.

    `ChartAnchor.add_error_bars(include=...)` maps ``"both"`` / ``"plus"`` /
    ``"minus"`` onto `BOTH` / `PLUS_VALUES` / `MINUS_VALUES` (the axis selector is
    the separate `axis=` parameter).
    """

    BOTH = 1
    PLUS_VALUES = 2
    MINUS_VALUES = 3
    NONE = 4


class XlErrorBarType(IntEnum):
    """`Series.ErrorBar(Type=...)` — how the error amount is computed.

    `ChartAnchor.add_error_bars(kind=...)` maps ``"fixed"`` / ``"percent"`` /
    ``"stdev"`` / ``"sterror"`` onto these. `STANDARD_ERROR` ignores the supplied
    amount (Word computes it from the series); the others use it.
    """

    FIXED_VALUE = 1
    PERCENT = 2
    STANDARD_DEVIATION = 3
    STANDARD_ERROR = 4
    CUSTOM = -4114
