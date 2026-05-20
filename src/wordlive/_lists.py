"""Lists & numbering — apply / read / restart bullets and numbered lists.

List operations act on a *range's paragraphs*, so the verbs themselves live on
the base `Anchor` (`apply_list`, `remove_list`, `list_info`, `restart_numbering`,
`indent_list`, `outdent_list`) and delegate to the small COM helpers here. This
module also holds the document-scoped `doc.lists` discovery collection.

Word's list model is notoriously fiddly: numbering comes from a `ListTemplate`
pulled out of one of three `Application.ListGalleries` (bullet / number / outline),
and "restart vs. continue the previous list" is a boolean on
`ListFormat.ApplyListTemplate`. We expose the 80%-useful shape — apply a fresh
bulleted/numbered/outline list, read what list a range is in, restart numbering,
and step a list item in/out a level — and leave custom list-template authoring to
the `.com` escape hatch.
"""

from __future__ import annotations

from typing import Any, Iterator, TYPE_CHECKING

from . import _com
from .constants import (
    WdDefaultListBehavior,
    WdListApplyTo,
    WdListGalleryType,
    WdListType,
)
from .exceptions import AnchorNotFoundError

if TYPE_CHECKING:
    from ._anchors import RangeAnchor
    from ._document import Document


# Accepted `list_type` strings -> the gallery they apply from. The three
# canonical names are bulleted / numbered / outline; common variants alias on.
_GALLERY_FOR: dict[str, WdListGalleryType] = {
    "bulleted": WdListGalleryType.BULLET,
    "bullet": WdListGalleryType.BULLET,
    "bullets": WdListGalleryType.BULLET,
    "numbered": WdListGalleryType.NUMBER,
    "number": WdListGalleryType.NUMBER,
    "numbers": WdListGalleryType.NUMBER,
    "outline": WdListGalleryType.OUTLINE_NUMBER,
    "outline-number": WdListGalleryType.OUTLINE_NUMBER,
    "outline_number": WdListGalleryType.OUTLINE_NUMBER,
}

_CANONICAL_TYPES = ("bulleted", "numbered", "outline")

# ListFormat.ListType int -> the human string list_info() reports.
_LIST_TYPE_NAMES: dict[int, str] = {
    int(WdListType.NO_NUMBERING): "none",
    int(WdListType.LIST_NUM_ONLY): "number-only",
    int(WdListType.BULLET): "bulleted",
    int(WdListType.SIMPLE_NUMBERING): "numbered",
    int(WdListType.OUTLINE_NUMBERING): "outline",
    int(WdListType.MIXED_NUMBERING): "mixed",
}


def gallery_for(list_type: str) -> WdListGalleryType:
    """Resolve a `list_type` string to its `WdListGalleryType`.

    Raises `ValueError` for an unknown name — symmetric with how
    `_coerce_alignment` rejects bad alignment strings.
    """
    try:
        return _GALLERY_FOR[str(list_type).lower()]
    except KeyError:
        raise ValueError(
            f"unknown list type {list_type!r}; expected one of {list(_CANONICAL_TYPES)}"
        )


def _read(lf: Any, attr: str, default: Any) -> Any:
    """Read a `ListFormat` attribute, tolerating non-list ranges that raise."""
    try:
        value = getattr(lf, attr)
    except Exception:
        return default
    return default if value is None else value


def apply_list_template(rng: Any, gallery_type: WdListGalleryType, *, continue_previous: bool) -> None:
    """Apply gallery `gallery_type`'s first template to `rng`'s paragraphs."""
    app = rng.Application
    gallery = app.ListGalleries(int(gallery_type))
    template = gallery.ListTemplates(1)
    rng.ListFormat.ApplyListTemplate(
        ListTemplate=template,
        ContinuePreviousList=bool(continue_previous),
        ApplyTo=int(WdListApplyTo.WHOLE_LIST),
        DefaultListBehavior=int(WdDefaultListBehavior.WORD10),
    )


def restart_numbering(rng: Any) -> None:
    """Re-apply the range's current list template starting at 1.

    Raises `ValueError` if the range isn't part of a list (no template to
    restart).
    """
    lf = rng.ListFormat
    template = _read(lf, "ListTemplate", None)
    if template is None:
        raise ValueError("range is not part of a list; nothing to restart")
    lf.ApplyListTemplate(
        ListTemplate=template,
        ContinuePreviousList=False,
        ApplyTo=int(WdListApplyTo.WHOLE_LIST),
        DefaultListBehavior=int(WdDefaultListBehavior.WORD10),
    )


def read_list_info(rng: Any) -> dict[str, Any]:
    """Describe the list a range sits in: `{type, level, number, string}`.

    `type` is `"none"` when the range carries no list formatting. `number` is
    the value of the first paragraph's number (0 for bullets), and `string` is
    its rendered marker (`"1."`, `"a)"`, `"•"`, …).
    """
    lf = rng.ListFormat
    try:
        list_type = int(_read(lf, "ListType", 0))
    except (TypeError, ValueError):
        list_type = 0
    try:
        level = int(_read(lf, "ListLevelNumber", 0))
    except (TypeError, ValueError):
        level = 0
    try:
        number = int(_read(lf, "ListValue", 0))
    except (TypeError, ValueError):
        number = 0
    return {
        "type": _LIST_TYPE_NAMES.get(list_type, "unknown"),
        "level": level,
        "number": number,
        "string": str(_read(lf, "ListString", "") or ""),
    }


class ListCollection:
    """Read-only, iterable view over the document's lists (`doc.lists`).

    Index a list by 1-based position (`doc.lists[2]`) to get a
    [`RangeAnchor`][wordlive.RangeAnchor] over its whole range — so every list
    verb (`apply_list`, `restart_numbering`, …) is immediately available on it.
    `list()` returns a summary per list; positions match Word's own
    `Document.Lists(n)` ordering.
    """

    def __init__(self, doc: "Document") -> None:
        self._doc = doc

    def __len__(self) -> int:
        with _com.translate_com_errors():
            return int(self._doc.com.Lists.Count)

    def _spans(self) -> list[tuple[int, int]]:
        with _com.translate_com_errors():
            count = int(self._doc.com.Lists.Count)
            spans: list[tuple[int, int]] = []
            for i in range(1, count + 1):
                rng = self._doc.com.Lists(i).Range
                spans.append((int(rng.Start), int(rng.End)))
        return spans

    def __getitem__(self, index: int) -> "RangeAnchor":
        if isinstance(index, bool) or not isinstance(index, int):
            raise TypeError(f"list index must be int, got {type(index).__name__}")
        n = len(self)
        if not (1 <= index <= n):
            raise AnchorNotFoundError("list", str(index))
        with _com.translate_com_errors():
            rng = self._doc.com.Lists(index).Range
            start, end = int(rng.Start), int(rng.End)
        return self._doc.range(start, end)

    def __iter__(self) -> Iterator["RangeAnchor"]:
        for start, end in self._spans():
            yield self._doc.range(start, end)

    def list(self) -> list[dict[str, Any]]:
        """All lists as `{index, type, count, anchor_id}` dicts."""
        out: list[dict[str, Any]] = []
        with _com.translate_com_errors():
            count = int(self._doc.com.Lists.Count)
            for i in range(1, count + 1):
                lst = self._doc.com.Lists(i)
                rng = lst.Range
                start, end = int(rng.Start), int(rng.End)
                info = read_list_info(rng)
                try:
                    n_items = int(lst.ListParagraphs.Count)
                except Exception:
                    n_items = 0
                out.append(
                    {
                        "index": i,
                        "type": info["type"],
                        "count": n_items,
                        "anchor_id": f"range:{start}-{end}",
                    }
                )
        return out
