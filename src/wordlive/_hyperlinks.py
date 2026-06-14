"""Hyperlinks — read the document's links as structured data.

wordlive can already *create* links (`anchor.link_to(...)` / the `add_hyperlink`
op); `doc.hyperlinks` is the read mirror — the discovery half. It reports every
link's visible text, its destination (an external `address` or an internal
`sub_address` bookmark), and a `range:START-END` id over the link so a hit can be
fed straight back into `read` / `replace` / `comments.add`.

Hyperlinks are addressed by 1-based index (`doc.hyperlinks[2]`), matching Word's
own `Hyperlinks(n)` ordering (document order). Reading is non-mutating.
"""

from __future__ import annotations

from collections.abc import Iterator
from typing import TYPE_CHECKING, Any

from . import _com
from .exceptions import AnchorNotFoundError

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

    def __repr__(self) -> str:
        dest = self.address or self.sub_address or "?"
        return f"<Hyperlink {self._index} -> {dest!r}>"


class HyperlinkCollection:
    """Indexable, iterable, read-only view over a document's hyperlinks (`doc.hyperlinks`)."""

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
