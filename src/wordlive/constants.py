"""Typed enums for the Word magic constants wordlive uses.

Values mirror the official `Wd*` enumerations exactly. Resist the urge to
pre-populate — add entries only as v0 features need them.
"""

from __future__ import annotations

from enum import IntEnum


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
    """

    ACTIVE_END_PAGE_NUMBER = 3
    WITH_IN_TABLE = 12


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

    `insert_caption` defaults to `BELOW` (the figure convention).
    """

    ABOVE = 0
    BELOW = 1


class WdFieldType(IntEnum):
    """`Fields.Add(Type=...)` values — the field kinds `insert_field(kind=...)` exposes.

    A narrow, agent-useful slice of Word's ~100 field types. `EMPTY` is the
    raw-code escape hatch: `insert_field("field", text="REF foo \\\\h")` inserts an
    arbitrary field whose code is carried in the text, so REF/TOC/etc. are
    reachable without a member each. The named members cover the publishing
    staples — page numbers, counts, date/time, and document metadata.
    """

    EMPTY = -1
    AUTHOR = 17
    NUM_PAGES = 26
    FILE_NAME = 29
    DATE = 31
    TIME = 32
    PAGE = 33
    TITLE = 49
