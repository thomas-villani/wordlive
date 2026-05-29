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
