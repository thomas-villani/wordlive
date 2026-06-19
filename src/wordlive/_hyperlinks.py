"""Hyperlinks — read and edit the document's links as structured data.

wordlive *creates* links via `anchor.link_to(...)` / the `add_hyperlink` op;
`doc.hyperlinks` is the discovery + edit mirror. It reports every link's visible
text, its destination (an external `address` or an internal `sub_address`
bookmark), and a `range:START-END` id over the link so a hit can be fed straight
back into `read` / `replace` / `comments.add` — and a `Hyperlink.update(...)`
retargets / relabels it in place (the `set_hyperlink` op / CLI / MCP).

Hyperlinks are addressed by 1-based index (`doc.hyperlinks[2]`), matching Word's
own `Hyperlinks(n)` ordering (document order). Listing is non-mutating.
"""

from __future__ import annotations

from collections.abc import Iterator
from typing import TYPE_CHECKING, Any

from . import _com
from .exceptions import AnchorNotFoundError, OpError

if TYPE_CHECKING:
    from ._document import Document


def _safe_str(obj: Any, attr: str) -> str:
    """`str(obj.attr)` or `""` — some link attributes raise on access."""
    try:
        return str(getattr(obj, attr) or "")
    except Exception:
        return ""


class Hyperlink:
    """A single hyperlink, located by its 1-based document index."""

    def __init__(self, doc: Document, com: Any, index: int) -> None:
        self._doc = doc
        self._com = com
        self._index = index

    @property
    def com(self) -> Any:
        """Raw COM Hyperlink object — escape hatch (Follow, Delete, sub-ranges, …)."""
        return self._com

    @property
    def index(self) -> int:
        return self._index

    @property
    def text(self) -> str:
        """The link's visible (clickable) text."""
        return _safe_str(self._com, "TextToDisplay")

    @property
    def address(self) -> str:
        """The external destination (URL / mailto / file path), or ``""`` for an internal link."""
        return _safe_str(self._com, "Address")

    @property
    def sub_address(self) -> str:
        """The in-document target — a bookmark name for an internal jump, else ``""``."""
        return _safe_str(self._com, "SubAddress")

    def to_dict(self) -> dict[str, Any]:
        """`{index, text, address, sub_address, screen_tip, anchor_id, para}` — the `list()` shape.

        `anchor_id` is a `range:START-END` over the link's range; `para` is the
        `para:N` the link sits in (or ``None``). For an internal link `address`
        is empty and `sub_address` holds the bookmark name it points at.
        """
        with _com.translate_com_errors():
            rng = self._com.Range
            try:
                start, end = int(rng.Start), int(rng.End)
            except Exception:
                start, end = None, None
        para_id: str | None = None
        if start is not None:
            para = self._doc.paragraphs.at(start)
            para_id = para.anchor_id if para is not None else None
        return {
            "index": self._index,
            "text": self.text,
            "address": self.address,
            "sub_address": self.sub_address,
            "screen_tip": _safe_str(self._com, "ScreenTip"),
            "anchor_id": f"range:{start}-{end}" if start is not None else None,
            "para": para_id,
        }

    def update(
        self,
        *,
        address: str | None = None,
        sub_address: str | None = None,
        text: str | None = None,
        screen_tip: str | None = None,
    ) -> Hyperlink:
        """Retarget / relabel this link in place — no delete + reinsert.

        Pass a string to set a field; omit it (or pass ``None``) to leave it.
        `address` is the external destination (URL / mailto / file path);
        `sub_address` is the in-document target (a bookmark name); `text` is the
        visible clickable text; `screen_tip` is the hover tooltip. `address` and
        `sub_address` stay orthogonal — setting one does not clear the other.

        These setters *retarget*, they don't unlink. `sub_address` and
        `screen_tip` can be emptied with ``""``, but Word keeps every link
        pointing somewhere with visible text, so `address` and `text` **cannot**
        be cleared (passing ``""`` raises `OpError` — delete the link via `.com`
        to remove it). Returns `self` (chainable); wrap in `doc.edit(...)` for
        atomic undo. Bad input raises `OpError`.
        """
        try:
            with _com.translate_com_errors():
                if address is not None:
                    if address == "":
                        raise ValueError(
                            "a hyperlink's address cannot be cleared; delete the link to remove it"
                        )
                    self._com.Address = str(address)
                if sub_address is not None:
                    self._com.SubAddress = str(sub_address)
                if text is not None:
                    if text == "":
                        raise ValueError(
                            "a hyperlink's visible text cannot be cleared; "
                            "delete the link to remove it"
                        )
                    self._com.TextToDisplay = str(text)
                if screen_tip is not None:
                    self._com.ScreenTip = str(screen_tip)
        except (ValueError, TypeError) as e:
            raise OpError(str(e)) from e
        return self

    def set_address(self, address: str) -> Hyperlink:
        """Set the external destination (URL / mailto / file path)."""
        return self.update(address=address)

    def set_sub_address(self, sub_address: str) -> Hyperlink:
        """Set the in-document target (a bookmark name); ``""`` clears it."""
        return self.update(sub_address=sub_address)

    def set_text(self, text: str) -> Hyperlink:
        """Set the visible (clickable) text."""
        return self.update(text=text)

    def set_screen_tip(self, screen_tip: str) -> Hyperlink:
        """Set the hover tooltip; ``""`` clears it."""
        return self.update(screen_tip=screen_tip)

    def __repr__(self) -> str:
        dest = self.address or self.sub_address or "?"
        return f"<Hyperlink {self._index} -> {dest!r}>"


class HyperlinkCollection:
    """Indexable, iterable view over a document's hyperlinks (`doc.hyperlinks`).

    Listing is read-only; index a [`Hyperlink`][wordlive.Hyperlink] to edit it
    in place (`doc.hyperlinks[2].update(...)`).
    """

    def __init__(self, doc: Document) -> None:
        self._doc = doc

    def __len__(self) -> int:
        with _com.translate_com_errors():
            return int(self._doc.com.Hyperlinks.Count)

    def __getitem__(self, index: int) -> Hyperlink:
        if isinstance(index, bool) or not isinstance(index, int):
            raise TypeError(f"hyperlink index must be int, got {type(index).__name__}")
        n = len(self)
        if not (1 <= index <= n):
            raise AnchorNotFoundError("hyperlink", str(index))
        with _com.translate_com_errors():
            return Hyperlink(self._doc, self._doc.com.Hyperlinks(index), index)

    def __iter__(self) -> Iterator[Hyperlink]:
        with _com.translate_com_errors():
            count = int(self._doc.com.Hyperlinks.Count)
        for i in range(1, count + 1):
            with _com.translate_com_errors():
                com = self._doc.com.Hyperlinks(i)
            yield Hyperlink(self._doc, com, i)

    def list(self) -> list[dict[str, Any]]:
        """Every hyperlink as `{index, text, address, sub_address, screen_tip, anchor_id, para}`."""
        return [h.to_dict() for h in self]
