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
    BookmarkCollection,
    ContentControl,
    EndAnchor,
    Heading,
    HeadingCollection,
    Paragraph,
    ParagraphCollection,
    RangeAnchor,
    StartAnchor,
)
from ._app import Word, attach, connect
from ._comments import Comment, CommentCollection
from ._document import Document, DocumentCollection
from ._edit import EditScope
from ._lists import ListCollection
from ._notes import Endnote, EndnoteCollection, Footnote, FootnoteCollection
from ._sections import HeaderFooter, Section, SectionCollection
from ._selection import Selection, SelectionSnapshot
from ._snapshot import Snapshot
from ._styles import Style, StyleCollection
from ._tables import Cell, Table, TableCollection
from ._toc import Toc
from .exceptions import (
    AmbiguousMatchError,
    AnchorNotFoundError,
    ComError,
    DocumentNotFoundError,
    ImageSourceError,
    ReplaceVerificationError,
    SnapshotError,
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
    "BookmarkCollection",
    "Cell",
    "ComError",
    "Comment",
    "CommentCollection",
    "ContentControl",
    "Document",
    "DocumentCollection",
    "DocumentNotFoundError",
    "EditScope",
    "EndAnchor",
    "Endnote",
    "EndnoteCollection",
    "Footnote",
    "FootnoteCollection",
    "Heading",
    "HeadingCollection",
    "HeaderFooter",
    "ImageSourceError",
    "ListCollection",
    "Paragraph",
    "ParagraphCollection",
    "RangeAnchor",
    "ReplaceVerificationError",
    "Section",
    "SectionCollection",
    "Selection",
    "SelectionSnapshot",
    "Snapshot",
    "SnapshotError",
    "StartAnchor",
    "Style",
    "StyleCollection",
    "StyleNotFoundError",
    "Table",
    "TableCollection",
    "Toc",
    "Word",
    "WordBusyError",
    "WordNotRunningError",
    "WordliveError",
    "attach",
    "connect",
    "constants",
]
