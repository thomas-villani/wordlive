"""`cc:NAME` content-control anchors and their collection."""

from __future__ import annotations

from collections.abc import Iterator
from typing import TYPE_CHECKING, Any

from .. import _charts, _com
from ..exceptions import AnchorNotFoundError, OpError

if TYPE_CHECKING:
    from .._document import Document

from ._base import Anchor
from ._helpers import (
    _CC_LIST_TYPES,
    _normalize_cc_items,
    range_text,
)

# ---------------------------------------------------------------------------
# Content controls
# ---------------------------------------------------------------------------


def _cc_by_name(doc_com: Any, name: str) -> Any | None:
    """Find a content control by its Title (Tag falls back). Returns None if missing.

    Reject empty `name` explicitly — many content controls have neither a
    Title nor a Tag, and the naive `cc.Title or "" == ""` test would match
    the first such control. Callers asking for `""` get `None` instead.
    """
    if not name:
        return None
    for cc in doc_com.ContentControls:
        if str(cc.Title or "") == name or str(cc.Tag or "") == name:
            return cc
    return None


class ContentControl(Anchor):
    kind = "content_control"

    def __init__(self, doc: Document, name: str, *, com: Any | None = None) -> None:
        super().__init__(doc, name)
        # A freshly created control caches its live COM object, so the returned
        # wrapper resolves even when unnamed (and survives edits — Word maintains
        # the control's identity). Named lookups still go through `_cc_by_name`.
        self._cc_com = com

    @property
    def anchor_id(self) -> str:
        return f"cc:{self.name}"

    def _cc(self) -> Any:
        if self._cc_com is not None:
            return self._cc_com
        cc = _cc_by_name(self._doc.com, self.name)
        if cc is None:
            raise AnchorNotFoundError("content_control", self.name)
        return cc

    def _range(self) -> Any:
        return self._cc().Range

    @property
    def text(self) -> str:
        with _com.translate_com_errors():
            cc = self._cc()
            return range_text(cc.Range)

    def set_text(self, text: str) -> None:
        with _com.translate_com_errors():
            cc = self._cc()
            cc.Range.Text = text

    def set_properties(
        self,
        *,
        title: Any = _charts._UNSET,
        tag: Any = _charts._UNSET,
        lock_contents: bool | None = None,
        lock_control: bool | None = None,
    ) -> ContentControl:
        """Re-set this control's metadata in place — no delete + reinsert.

        Tri-state: omit a field to leave it untouched, pass a string to set it,
        or `None` (equivalently ``""``) to clear `title` / `tag`. `lock_contents`
        stops the user editing the value; `lock_control` stops them deleting the
        control — pass a bool to set either, omit to leave. Renaming the title
        (or the tag, when untitled) changes the control's `cc:NAME` anchor id;
        the returned handle keeps working regardless. Returns `self` (chainable);
        wrap in `doc.edit(...)` for atomic undo. Bad input raises `OpError`.
        """
        try:
            with _com.translate_com_errors():
                cc = self._cc()
                if title is not _charts._UNSET:
                    cc.Title = "" if title is None else str(title)
                if tag is not _charts._UNSET:
                    cc.Tag = "" if tag is None else str(tag)
                if lock_contents is not None:
                    cc.LockContents = bool(lock_contents)
                if lock_control is not None:
                    cc.LockContentControl = bool(lock_control)
                if title is not _charts._UNSET or tag is not _charts._UNSET:
                    # Keep anchor_id honest after a rename (title beats tag).
                    self.name = str(cc.Title or cc.Tag or "")
        except (ValueError, TypeError) as e:
            raise OpError(str(e)) from e
        return self

    def set_items(self, items: list[Any]) -> ContentControl:
        """Replace a combo-box / dropdown's choice list — no delete + reinsert.

        `items` is the full new list (it replaces the existing entries, not
        appends); each is a string or a `{"text": ..., "value": ...}` dict.
        Only valid on a ``combo_box`` / ``dropdown`` control — any other kind
        raises `OpError`. Returns `self` (chainable); wrap in `doc.edit(...)`
        for atomic undo. Bad input raises `OpError`.
        """
        try:
            with _com.translate_com_errors():
                cc = self._cc()
                if cc.Type not in _CC_LIST_TYPES:
                    raise ValueError(
                        "set_items is only valid for a 'combo_box' or 'dropdown' control"
                    )
                pairs = _normalize_cc_items(items)
                entries = cc.DropdownListEntries
                entries.Clear()
                for entry_text, value in pairs:
                    entries.Add(entry_text, value)  # positional (Text, Value)
        except (ValueError, TypeError) as e:
            raise OpError(str(e)) from e
        return self


class ContentControlCollection:
    def __init__(self, doc: Document) -> None:
        self._doc = doc

    def __getitem__(self, name: str) -> ContentControl:
        with _com.translate_com_errors():
            if _cc_by_name(self._doc.com, name) is None:
                raise AnchorNotFoundError("content_control", name)
        return ContentControl(self._doc, name)

    def __contains__(self, name: object) -> bool:
        if not isinstance(name, str):
            return False
        with _com.translate_com_errors():
            return _cc_by_name(self._doc.com, name) is not None

    def list(self) -> list[str]:
        with _com.translate_com_errors():
            names: list[str] = []
            for cc in self._doc.com.ContentControls:
                names.append(str(cc.Title or cc.Tag or ""))
            return names

    def __iter__(self) -> Iterator[ContentControl]:
        for name in self.list():
            if name:
                yield ContentControl(self._doc, name)

    def add(self, anchor: Anchor | str, kind: str = "rich_text", **kwargs: Any) -> ContentControl:
        """Create a content control over an anchor and return it.

        Symmetric with [`bookmarks.add`][wordlive.BookmarkCollection.add]: a
        document-level entry point for the per-anchor
        [`insert_content_control`][wordlive.Anchor.insert_content_control].
        `anchor` is an [`Anchor`][wordlive.Anchor] or an anchor-id string
        (`range:START-END`, `cc:NAME`, `heading:N`, …); `kind` and the keyword
        options (`title`, `tag`, `items`, `where`, `lock_contents`,
        `lock_control`) pass straight through. Wrap in `doc.edit(...)` for atomic
        undo.
        """
        target = self._doc.anchor_by_id(anchor) if isinstance(anchor, str) else anchor
        return target.insert_content_control(kind, **kwargs)
