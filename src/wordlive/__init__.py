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
from ._anchors import (
    Anchor,
    Bookmark,
    ContentControl,
    Heading,
    HeadingCollection,
    Paragraph,
    ParagraphCollection,
    RangeAnchor,
)
from ._app import Word, attach, connect
from ._comments import Comment, CommentCollection
from ._document import Document, DocumentCollection
from ._edit import EditScope
from ._lists import ListCollection
from ._sections import HeaderFooter, Section, SectionCollection
from ._selection import Selection, SelectionSnapshot
from ._styles import Style, StyleCollection
from ._tables import Cell, Table, TableCollection
from .exceptions import (
    AmbiguousMatchError,
    AnchorNotFoundError,
    ComError,
    DocumentNotFoundError,
    ImageSourceError,
    StyleNotFoundError,
    WordBusyError,
    WordliveError,
    WordNotRunningError,
)

__all__ = [
    "AmbiguousMatchError",
    "Anchor",
    "AnchorNotFoundError",
    "Bookmark",
    "Cell",
    "ComError",
    "Comment",
    "CommentCollection",
    "ContentControl",
    "Document",
    "DocumentCollection",
    "DocumentNotFoundError",
    "EditScope",
    "Heading",
    "HeadingCollection",
    "HeaderFooter",
    "ImageSourceError",
    "ListCollection",
    "Paragraph",
    "ParagraphCollection",
    "RangeAnchor",
    "Section",
    "SectionCollection",
    "Selection",
    "SelectionSnapshot",
    "Style",
    "StyleCollection",
    "StyleNotFoundError",
    "Table",
    "TableCollection",
    "Word",
    "WordBusyError",
    "WordNotRunningError",
    "WordliveError",
    "attach",
    "connect",
    "constants",
]
