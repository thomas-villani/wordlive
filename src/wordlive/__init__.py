"""wordlive — drive a running Microsoft Word instance from Python.

Quick start:

    import wordlive as wl

    with wl.attach() as word:
        doc = word.documents.active
        with doc.edit("Update address"):
            doc.bookmarks["Address"].set_text("123 Main St")
"""

from __future__ import annotations

from . import constants
from ._anchors import Anchor, Bookmark, ContentControl, Heading, HeadingCollection
from ._app import Word, attach, connect
from ._document import Document, DocumentCollection
from ._edit import EditScope
from ._selection import Selection, SelectionSnapshot
from ._styles import Style, StyleCollection
from .exceptions import (
    AmbiguousMatchError,
    AnchorNotFoundError,
    ComError,
    DocumentNotFoundError,
    StyleNotFoundError,
    WordBusyError,
    WordNotRunningError,
    WordliveError,
)

__all__ = [
    "AmbiguousMatchError",
    "Anchor",
    "AnchorNotFoundError",
    "Bookmark",
    "ComError",
    "ContentControl",
    "Document",
    "DocumentCollection",
    "DocumentNotFoundError",
    "EditScope",
    "Heading",
    "HeadingCollection",
    "Selection",
    "SelectionSnapshot",
    "Style",
    "StyleCollection",
    "StyleNotFoundError",
    "Word",
    "WordBusyError",
    "WordNotRunningError",
    "WordliveError",
    "attach",
    "connect",
    "constants",
]
