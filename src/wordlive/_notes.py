"""Footnotes & endnotes — note structures anchored to a range.

A footnote/endnote pairs a *reference mark* in the main text story with a *note
body* that lives in its own story (Word keeps footnote and endnote text in
separate `StoryRanges`). The insertion verbs (`insert_footnote` /
`insert_endnote`) live on the base `Anchor` — a note attaches to whatever range
you target, exactly like `insert_table` / `insert_field` — and return the note
as an addressable anchor.

A note is addressed by 1-based index (`footnote:N` / `endnote:N`), matching
Word's own `Footnotes(n)` / `Endnotes(n)` ordering, and resolves to the
**note-body range**: `set_text` edits the note, `delete()` removes the mark and
its body together. `doc.footnotes` / `doc.endnotes` are read-only discovery
collections (siblings of `doc.comments` / `doc.lists`).
"""

from __future__ import annotations

from collections.abc import Iterator
from typing import TYPE_CHECKING, Any

from . import _com
from ._anchors import Anchor, range_text
from .exceptions import AnchorNotFoundError

if TYPE_CHECKING:
    from ._document import Document


def _clean_note(raw: Any) -> str:
    """Strip a note body's leading reference mark and trailing paragraph markers.

    Word stores a footnote/endnote body with the auto-number reference mark
    (a non-printing control character) at the front and the usual paragraph /
    cell terminators at the back; neither is meaningful note text.
    """
    text = str(raw or "")
    # Leading auto-number / custom-mark control chars, then any separator space.
    text = text.lstrip("\x01\x02\x03\x04\x05")
    return text.strip("\r\n\x07 \t")


def index_of_note(coll: Any, pos: int) -> int:
    """1-based index of the note in `coll` whose reference mark sits at `pos`.

    `Document.Footnotes` / `Endnotes` are ordered by document position, not
    insertion order, so a note inserted before existing ones is *not* at
    `Count`. Positions are stable ints (COM-identity comparison is not
    reliable), so we match on the reference mark's `Start`. Falls back to
    `Count` if no match is found (the just-appended common case).
    """
    count = int(coll.Count)
    for i in range(1, count + 1):
        try:
            if int(coll(i).Reference.Start) == pos:
                return i
        except Exception:
            continue
    return count


class _NoteAnchor(Anchor):
    """Shared base for footnote / endnote anchors located by 1-based index.

    Subclasses set `kind`, `_scheme` (`"footnote"` / `"endnote"`), and
    `_collection_attr` (the COM document collection name).
    """

    _scheme: str = "note"
    _collection_attr: str = ""

    def __init__(self, doc: Document, index: int) -> None:
        super().__init__(doc, name=f"{self._scheme}:{index}")
        self._index = index

    @property
    def index(self) -> int:
        return self._index

    @property
    def anchor_id(self) -> str:
        return f"{self._scheme}:{self._index}"

    def _collection(self) -> Any:
        return getattr(self._doc.com, self._collection_attr)

    def _note(self) -> Any:
        coll = self._collection()
        n = int(coll.Count)
        if not (1 <= self._index <= n):
            raise AnchorNotFoundError(self._scheme, f"{self._scheme}:{self._index}")
        return coll(self._index)

    def _range(self) -> Any:
        return self._note().Range

    @property
    def text(self) -> str:
        """The note body text (its reference mark and terminators stripped)."""
        with _com.translate_com_errors():
            return _clean_note(range_text(self._note().Range))

    def set_text(self, text: str) -> None:
        """Replace the note body text in place."""
        with _com.translate_com_errors():
            self._note().Range.Text = text

    def delete(self) -> None:
        """Remove the note — both its reference mark and its body.

        The base `Anchor.delete()` would clear only the body range; a note is
        removed by deleting its reference mark in the main text, which takes the
        body with it.
        """
        with _com.translate_com_errors():
            self._note().Reference.Delete()


class Footnote(_NoteAnchor):
    """A footnote, addressed `footnote:N` and resolving to its note-body range."""

    kind = "footnote"
    _scheme = "footnote"
    _collection_attr = "Footnotes"


class Endnote(_NoteAnchor):
    """An endnote, addressed `endnote:N` and resolving to its note-body range."""

    kind = "endnote"
    _scheme = "endnote"
    _collection_attr = "Endnotes"


class _NoteCollection:
    """Read-only, iterable view over a document's footnotes or endnotes.

    Index a note by 1-based position (`doc.footnotes[2]`) to get its anchor —
    so `set_text` / `delete` are immediately available. `list()` returns a
    summary per note; positions match Word's own `Footnotes(n)` / `Endnotes(n)`
    ordering.
    """

    _scheme: str = "note"
    _collection_attr: str = ""
    _anchor_cls: type[_NoteAnchor] = _NoteAnchor

    def __init__(self, doc: Document) -> None:
        self._doc = doc

    def _collection(self) -> Any:
        return getattr(self._doc.com, self._collection_attr)

    def __len__(self) -> int:
        with _com.translate_com_errors():
            return int(self._collection().Count)

    def __getitem__(self, index: int) -> _NoteAnchor:
        if isinstance(index, bool) or not isinstance(index, int):
            raise TypeError(f"{self._scheme} index must be int, got {type(index).__name__}")
        n = len(self)
        if not (1 <= index <= n):
            raise AnchorNotFoundError(self._scheme, str(index))
        return self._anchor_cls(self._doc, index)

    def __iter__(self) -> Iterator[_NoteAnchor]:
        with _com.translate_com_errors():
            count = int(self._collection().Count)
        for i in range(1, count + 1):
            yield self._anchor_cls(self._doc, i)

    def list(self) -> list[dict[str, Any]]:
        """All notes as `{index, anchor_id, marker, text, para}` dicts.

        `marker` is the note's 1-based number (the rendered mark for default
        auto-numbering). `para` is the `para:N` anchor of the paragraph holding
        the reference mark (or `None` if it can't be located) — so an agent can
        see *where* each note is anchored without reading the whole document.
        """
        out: list[dict[str, Any]] = []
        with _com.translate_com_errors():
            coll = self._collection()
            count = int(coll.Count)
            for i in range(1, count + 1):
                note = coll(i)
                try:
                    ref_start = int(note.Reference.Start)
                except Exception:
                    ref_start = None
                para_id: str | None = None
                if ref_start is not None:
                    para = self._doc.paragraphs.at(ref_start)
                    para_id = para.anchor_id if para is not None else None
                out.append(
                    {
                        "index": i,
                        "anchor_id": f"{self._scheme}:{i}",
                        "marker": str(i),
                        "text": _clean_note(range_text(note.Range)),
                        "para": para_id,
                    }
                )
        return out


class FootnoteCollection(_NoteCollection):
    """`doc.footnotes` — read-only discovery over the document's footnotes."""

    _scheme = "footnote"
    _collection_attr = "Footnotes"
    _anchor_cls = Footnote


class EndnoteCollection(_NoteCollection):
    """`doc.endnotes` — read-only discovery over the document's endnotes."""

    _scheme = "endnote"
    _collection_attr = "Endnotes"
    _anchor_cls = Endnote
