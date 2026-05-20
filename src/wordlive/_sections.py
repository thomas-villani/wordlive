"""Sections, headers, and footers.

A header/footer is, structurally, just a `Range` — so `HeaderFooter` *is* an
`Anchor` (like `Cell`): it inherits `text` / `set_text` / `apply_style` /
`format_paragraph` and only overrides the bits that differ (the COM range and
the anchor id). That makes "put the client name in the first-page header" the
same `set_text` call as any other write, and it round-trips through
`Document.anchor_by_id` as `header:S:WHICH` / `footer:S:WHICH`.

`WHICH` is one of `primary` / `first` / `even` (Word's primary, first-page, and
even-pages header/footer). `S` is the 1-based section index. The document-scoped
`doc.sections` collection reaches each section's headers, footers, and a
read-only `page_setup()` summary.
"""

from __future__ import annotations

from typing import Any, Iterator, TYPE_CHECKING

from . import _com
from ._anchors import Anchor
from .constants import WdHeaderFooterIndex, WdOrientation
from .exceptions import AnchorNotFoundError

if TYPE_CHECKING:
    from ._document import Document


# Accepted `which` strings -> WdHeaderFooterIndex. Canonical names are
# primary / first / even; the verbose variants alias on.
_WHICH_TO_INDEX: dict[str, WdHeaderFooterIndex] = {
    "primary": WdHeaderFooterIndex.PRIMARY,
    "first": WdHeaderFooterIndex.FIRST_PAGE,
    "firstpage": WdHeaderFooterIndex.FIRST_PAGE,
    "first_page": WdHeaderFooterIndex.FIRST_PAGE,
    "first-page": WdHeaderFooterIndex.FIRST_PAGE,
    "even": WdHeaderFooterIndex.EVEN_PAGES,
    "evenpages": WdHeaderFooterIndex.EVEN_PAGES,
    "even_pages": WdHeaderFooterIndex.EVEN_PAGES,
    "even-pages": WdHeaderFooterIndex.EVEN_PAGES,
}

_CANONICAL_WHICH: dict[int, str] = {
    int(WdHeaderFooterIndex.PRIMARY): "primary",
    int(WdHeaderFooterIndex.FIRST_PAGE): "first",
    int(WdHeaderFooterIndex.EVEN_PAGES): "even",
}


def which_index(which: str) -> WdHeaderFooterIndex:
    """Resolve a `which` string to its `WdHeaderFooterIndex`.

    Raises `ValueError` for an unknown name.
    """
    try:
        return _WHICH_TO_INDEX[str(which).lower()]
    except KeyError:
        raise ValueError(
            f"unknown header/footer {which!r}; expected one of ['primary', 'first', 'even']"
        )


def _safe(obj: Any, attr: str, default: Any) -> Any:
    try:
        value = getattr(obj, attr)
    except Exception:
        return default
    return default if value is None else value


class HeaderFooter(Anchor):
    """A section's header or footer, addressed as `header:S:WHICH` / `footer:S:WHICH`.

    Subclasses `Anchor`, so `text`, `set_text`, `insert_before/after`,
    `apply_style`, and `format_paragraph` all work unchanged — only the COM
    range and anchor id are overridden here. `WHICH` is `primary`, `first`, or
    `even`.
    """

    def __init__(self, doc: "Document", section_index: int, which: str, *, is_footer: bool) -> None:
        self._section_index = int(section_index)
        self._which = _CANONICAL_WHICH[int(which_index(which))]
        self._is_footer = bool(is_footer)
        self.kind = "footer" if is_footer else "header"
        super().__init__(doc, name=f"{self.kind}:{self._section_index}:{self._which}")

    @property
    def anchor_id(self) -> str:
        return f"{self.kind}:{self._section_index}:{self._which}"

    @property
    def section_index(self) -> int:
        return self._section_index

    @property
    def which(self) -> str:
        return self._which

    def _hf(self) -> Any:
        section = self._doc.com.Sections(self._section_index)
        collection = section.Footers if self._is_footer else section.Headers
        return collection(int(which_index(self._which)))

    def _range(self) -> Any:
        return self._hf().Range

    @property
    def text(self) -> str:
        # Strip the trailing paragraph mark so a one-line header reads cleanly,
        # mirroring Cell.text / paragraph_text.
        with _com.translate_com_errors():
            return str(self._hf().Range.Text or "").rstrip("\r\n\x07")

    def set_text(self, text: str) -> None:
        with _com.translate_com_errors():
            self._hf().Range.Text = text

    @property
    def exists(self) -> bool:
        """Whether this header/footer actually has content defined for the section."""
        with _com.translate_com_errors():
            return bool(self._hf().Exists)

    @property
    def linked_to_previous(self) -> bool:
        """Whether this header/footer inherits from the previous section's."""
        with _com.translate_com_errors():
            return bool(self._hf().LinkToPrevious)


class Section:
    """Wraps a Word `Section`, located by its 1-based document position."""

    def __init__(self, doc: "Document", com: Any, index: int) -> None:
        self._doc = doc
        self._com = com
        self._index = index

    @property
    def com(self) -> Any:
        return self._com

    @property
    def index(self) -> int:
        return self._index

    def header(self, which: str = "primary") -> HeaderFooter:
        """The section's header for `which` (`primary` / `first` / `even`)."""
        return HeaderFooter(self._doc, self._index, which, is_footer=False)

    def footer(self, which: str = "primary") -> HeaderFooter:
        """The section's footer for `which` (`primary` / `first` / `even`)."""
        return HeaderFooter(self._doc, self._index, which, is_footer=True)

    def page_setup(self) -> dict[str, Any]:
        """Read-only `{orientation, *_margin, page_width, page_height}` in points."""
        with _com.translate_com_errors():
            ps = self._com.PageSetup
            try:
                orientation = int(_safe(ps, "Orientation", 0))
            except (TypeError, ValueError):
                orientation = 0
            return {
                "orientation": "landscape"
                if orientation == int(WdOrientation.LANDSCAPE)
                else "portrait",
                "top_margin": float(_safe(ps, "TopMargin", 0.0)),
                "bottom_margin": float(_safe(ps, "BottomMargin", 0.0)),
                "left_margin": float(_safe(ps, "LeftMargin", 0.0)),
                "right_margin": float(_safe(ps, "RightMargin", 0.0)),
                "page_width": float(_safe(ps, "PageWidth", 0.0)),
                "page_height": float(_safe(ps, "PageHeight", 0.0)),
            }

    def to_dict(self) -> dict[str, Any]:
        """`{index, page_setup}` — the JSON shape `sections.list()` emits."""
        return {"index": self._index, "page_setup": self.page_setup()}

    def __repr__(self) -> str:
        return f"<Section {self._index}>"


class SectionCollection:
    """Indexable, iterable view over a document's sections (`doc.sections`).

    Index by 1-based position (`doc.sections[1]`). Every document has at least
    one section; `doc.sections[1].header()` is the common entry point.
    """

    def __init__(self, doc: "Document") -> None:
        self._doc = doc

    def __len__(self) -> int:
        with _com.translate_com_errors():
            return int(self._doc.com.Sections.Count)

    def __getitem__(self, index: int) -> Section:
        if isinstance(index, bool) or not isinstance(index, int):
            raise TypeError(f"section index must be int, got {type(index).__name__}")
        n = len(self)
        if not (1 <= index <= n):
            raise AnchorNotFoundError("section", str(index))
        with _com.translate_com_errors():
            com = self._doc.com.Sections(index)
        return Section(self._doc, com, index)

    def __iter__(self) -> Iterator[Section]:
        with _com.translate_com_errors():
            count = int(self._doc.com.Sections.Count)
        for i in range(1, count + 1):
            with _com.translate_com_errors():
                com = self._doc.com.Sections(i)
            yield Section(self._doc, com, i)

    def list(self) -> list[dict[str, Any]]:
        """All sections as `{index, page_setup}` dicts."""
        return [s.to_dict() for s in self]
