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
