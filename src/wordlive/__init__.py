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
    EquationAnchor,
    EquationCollection,
    Heading,
    HeadingCollection,
    ImageAnchor,
    ImageCollection,
    Paragraph,
    ParagraphCollection,
    RangeAnchor,
    StartAnchor,
)
from ._app import Word, attach, connect
from ._citations import Bibliography, Citation
from ._comments import Comment, CommentCollection
from ._document import Document, DocumentCollection
from ._edit import EditScope
from ._fields import Field, FieldCollection
from ._hyperlinks import Hyperlink, HyperlinkCollection
from ._index import Index
from ._lists import ListCollection
from ._notes import Endnote, EndnoteCollection, Footnote, FootnoteCollection
from ._properties import PropertyCollection
from ._revisions import Revision, RevisionCollection
from ._sections import HeaderFooter, Section, SectionCollection
from ._selection import Selection, SelectionSnapshot
from ._snapshot import Snapshot
from ._sources import Source
from ._styles import Style, StyleCollection
from ._tables import Cell, Table, TableCollection
from ._toa import TableOfAuthorities
from ._toc import TableOfFigures, Toc
from ._variables import VariableCollection
from .exceptions import (
    AmbiguousMatchError,
    AnchorNotFoundError,
    ComError,
    DocumentNotFoundError,
    EquationError,
    ImageSourceError,
    OpError,
    PathNotAllowedError,
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
    "Bibliography",
    "Bookmark",
    "BookmarkCollection",
    "Cell",
    "Citation",
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
    "EquationAnchor",
    "EquationCollection",
    "EquationError",
    "Field",
    "FieldCollection",
    "Footnote",
    "FootnoteCollection",
    "Heading",
    "HeadingCollection",
    "HeaderFooter",
    "Hyperlink",
    "HyperlinkCollection",
    "ImageAnchor",
    "ImageCollection",
    "ImageSourceError",
    "Index",
    "ListCollection",
    "OpError",
    "Paragraph",
    "ParagraphCollection",
    "PathNotAllowedError",
    "PropertyCollection",
    "RangeAnchor",
    "ReplaceVerificationError",
    "Revision",
    "RevisionCollection",
    "Section",
    "SectionCollection",
    "Selection",
    "SelectionSnapshot",
    "Snapshot",
    "SnapshotError",
    "Source",
    "StartAnchor",
    "Style",
    "StyleCollection",
    "StyleNotFoundError",
    "Table",
    "TableCollection",
    "TableOfAuthorities",
    "TableOfFigures",
    "Toc",
    "VariableCollection",
    "Word",
    "WordBusyError",
    "WordNotRunningError",
    "WordliveError",
    "attach",
    "connect",
    "constants",
]
