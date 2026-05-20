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
