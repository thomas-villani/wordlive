"""`bookmark:NAME` anchors, the bookmark collection, and wordlive's hidden pins."""

from __future__ import annotations

from collections.abc import Iterator
from typing import TYPE_CHECKING, Any

from .. import _com
from ..exceptions import AnchorNotFoundError

if TYPE_CHECKING:
    from .._document import Document

from ._base import Anchor
from ._helpers import (
    _utf16_len,
)
from ._refs import (
    _pin_name_for,
    _validate_bookmark_name,
)

# ---------------------------------------------------------------------------
# Bookmarks
# ---------------------------------------------------------------------------


class Bookmark(Anchor):
    kind = "bookmark"

    # Set when this bookmark is a wordlive durable handle (`_wl_<code>`): the
    # anchor then reports `pin:<code>` instead of `bookmark:_wl_<code>`.
    _pin_code: str | None = None

    @classmethod
    def pin(cls, doc: Document, code: str) -> Bookmark:
        """A `Bookmark` over the hidden `_wl_<code>` pin, reporting `pin:<code>`."""
        bm = cls(doc, _pin_name_for(code))
        bm._pin_code = code
        return bm

    @property
    def anchor_id(self) -> str:
        if self._pin_code is not None:
            return f"pin:{self._pin_code}"
        return f"bookmark:{self.name}"

    def _range(self) -> Any:
        doc_com = self._doc.com
        if not doc_com.Bookmarks.Exists(self.name):
            raise AnchorNotFoundError("bookmark", self.name)
        return doc_com.Bookmarks(self.name).Range

    def set_text(self, text: str) -> None:
        with _com.translate_com_errors():
            doc_com = self._doc.com
            if not doc_com.Bookmarks.Exists(self.name):
                raise AnchorNotFoundError("bookmark", self.name)
            rng = doc_com.Bookmarks(self.name).Range
            start = int(rng.Start)
            rng.Text = text
            # Setting Range.Text deletes the bookmark; re-add covering the new content.
            # Word measures Range offsets in UTF-16 code units, not Python code points.
            new_end = start + _utf16_len(text)
            new_rng = doc_com.Range(start, new_end)
            doc_com.Bookmarks.Add(Name=self.name, Range=new_rng)


def _bookmarks_including_hidden(doc_com: Any) -> list[Any]:
    """Every bookmark, *including* Word's hidden (leading-underscore) ones.

    Word omits underscore-prefixed bookmarks — its own `_Toc`/`_Ref` anchors and
    wordlive's `_wl_` pins — from `Document.Bookmarks` iteration unless the
    collection's `ShowHidden` flag is on (a real-Word behaviour the fake COM
    fixture doesn't model). We flip it on for the read and restore it after, so
    pin enumeration / idempotency see the `_wl_` handles. The fake has no
    `ShowHidden`, so the toggle is best-effort.
    """
    bms = doc_com.Bookmarks
    previous: bool | None
    try:
        previous = bool(bms.ShowHidden)
        bms.ShowHidden = True
    except Exception:
        previous = None
    try:
        return list(bms)
    finally:
        if previous is not None:
            try:
                bms.ShowHidden = previous
            except Exception:
                pass


def _is_user_bookmark(name: str) -> bool:
    """Word auto-creates internal bookmarks for TOC entries, cross-references,
    and form-field anchors — all of them named with a leading underscore. Those
    are noise for the user-facing `list()` / iteration paths; agents addressing
    them by exact name (via `bookmarks[name]`) still work.
    """
    return not name.startswith("_")


class BookmarkCollection:
    """Indexable view over a document's bookmarks.

    `list()` and iteration return only user-visible bookmarks. Word's hidden
    bookmarks (`_Toc...`, `_Ref...`, etc.) are filtered out by default; address
    them by their exact name through `bookmarks[name]` if you need them.
    """

    def __init__(self, doc: Document) -> None:
        self._doc = doc

    def __getitem__(self, name: str) -> Bookmark:
        with _com.translate_com_errors():
            if not self._doc.com.Bookmarks.Exists(name):
                raise AnchorNotFoundError("bookmark", name)
        return Bookmark(self._doc, name)

    def __contains__(self, name: object) -> bool:
        if not isinstance(name, str):
            return False
        with _com.translate_com_errors():
            return bool(self._doc.com.Bookmarks.Exists(name))

    def add(self, name: str, anchor: Anchor | str) -> Bookmark:
        """Create a bookmark named `name` over `anchor`'s range and return it.

        `anchor` is an [`Anchor`][wordlive.Anchor] or an anchor id string
        (resolved via `doc.anchor_by_id`). `name` is validated against Word's
        rules — it must start with a letter and contain only letters, digits, and
        underscores (no spaces), max 40 characters — and an invalid name raises
        `OpError` *before* anything is created. Adding a bookmark with an existing
        name moves it to the new range (Word's own behaviour). This is the
        prerequisite for internal hyperlinks
        ([`Anchor.link_to`][wordlive.Anchor.link_to]) and cross-references
        ([`Anchor.insert_cross_reference`][wordlive.Anchor.insert_cross_reference]).
        Wrap in `doc.edit(...)` for atomic undo.
        """
        _validate_bookmark_name(name)
        resolved = self._doc.anchor_by_id(anchor) if isinstance(anchor, str) else anchor
        with _com.translate_com_errors():
            rng = resolved.com
            self._doc.com.Bookmarks.Add(Name=name, Range=rng)
        return Bookmark(self._doc, name)

    def list(self, *, include_hidden: bool = False) -> list[str]:
        """Names of every user-visible bookmark in document order.

        Set `include_hidden=True` to also return Word's internal bookmarks
        (TOC entries, cross-references, etc.) whose names start with `_`.
        """
        with _com.translate_com_errors():
            if include_hidden:
                return [str(bm.Name) for bm in _bookmarks_including_hidden(self._doc.com)]
            names = [str(bm.Name) for bm in self._doc.com.Bookmarks]
        return [n for n in names if _is_user_bookmark(n)]

    def __iter__(self) -> Iterator[Bookmark]:
        for name in self.list():
            yield Bookmark(self._doc, name)
