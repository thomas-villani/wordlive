"""Anchor types — semantic handles for ranges inside a Word document.

Anchors target a `Range`, never the live `Selection`. Each public mutation
goes through the COM error translator. Operations are intentionally small;
they compose with `Document.edit()` for atomic-undo behaviour.
"""

from __future__ import annotations

from typing import Any, Iterator, TYPE_CHECKING

from . import _com
from .exceptions import AnchorNotFoundError

if TYPE_CHECKING:
    from ._document import Document


# ---------------------------------------------------------------------------
# Base
# ---------------------------------------------------------------------------


class Anchor:
    """Abstract base — subclasses know how to materialise their COM Range."""

    kind: str = "anchor"
    name: str = ""

    def __init__(self, doc: "Document", name: str) -> None:
        self._doc = doc
        self.name = name

    @property
    def com(self) -> Any:
        """Raw COM range. Subclasses override."""
        return self._range()

    def _range(self) -> Any:
        raise NotImplementedError

    @property
    def text(self) -> str:
        with _com.translate_com_errors():
            return str(self._range().Text or "")

    @property
    def anchor_id(self) -> str:
        """Stable string identifier for this anchor (e.g. `bookmark:Address`).

        Subclasses override; the base falls back to `{kind}:{name}` which is
        only useful for built-in kinds whose `kind` matches the anchor-id
        scheme.
        """
        return f"{self.kind}:{self.name}"

    def set_text(self, text: str) -> None:
        raise NotImplementedError

    def insert_before(self, text: str) -> None:
        with _com.translate_com_errors():
            rng = self._range()
            insert_rng = self._doc.com.Range(rng.Start, rng.Start)
            insert_rng.Text = text

    def insert_after(self, text: str) -> None:
        with _com.translate_com_errors():
            rng = self._range()
            insert_rng = self._doc.com.Range(rng.End, rng.End)
            insert_rng.Text = text

    def delete(self) -> None:
        with _com.translate_com_errors():
            self._range().Delete()

    def __repr__(self) -> str:
        return f"<{type(self).__name__} {self.name!r}>"


# ---------------------------------------------------------------------------
# Bookmarks
# ---------------------------------------------------------------------------


class Bookmark(Anchor):
    kind = "bookmark"

    @property
    def anchor_id(self) -> str:
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
            new_end = start + len(text)
            new_rng = doc_com.Range(start, new_end)
            doc_com.Bookmarks.Add(Name=self.name, Range=new_rng)


class BookmarkCollection:
    """Indexable view over a document's bookmarks."""

    def __init__(self, doc: "Document") -> None:
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

    def list(self) -> list[str]:
        with _com.translate_com_errors():
            return [str(bm.Name) for bm in self._doc.com.Bookmarks]

    def __iter__(self) -> Iterator[Bookmark]:
        for name in self.list():
            yield Bookmark(self._doc, name)


# ---------------------------------------------------------------------------
# Content controls
# ---------------------------------------------------------------------------


def _cc_by_name(doc_com: Any, name: str) -> Any | None:
    """Find a content control by its Title (Tag falls back). Returns None if missing."""
    for cc in doc_com.ContentControls:
        if str(cc.Title or "") == name or str(cc.Tag or "") == name:
            return cc
    return None


class ContentControl(Anchor):
    kind = "content control"

    @property
    def anchor_id(self) -> str:
        return f"cc:{self.name}"

    def _cc(self) -> Any:
        cc = _cc_by_name(self._doc.com, self.name)
        if cc is None:
            raise AnchorNotFoundError("content control", self.name)
        return cc

    def _range(self) -> Any:
        return self._cc().Range

    @property
    def text(self) -> str:
        with _com.translate_com_errors():
            cc = self._cc()
            return str(cc.Range.Text or "")

    def set_text(self, text: str) -> None:
        with _com.translate_com_errors():
            cc = self._cc()
            cc.Range.Text = text


class ContentControlCollection:
    def __init__(self, doc: "Document") -> None:
        self._doc = doc

    def __getitem__(self, name: str) -> ContentControl:
        with _com.translate_com_errors():
            if _cc_by_name(self._doc.com, name) is None:
                raise AnchorNotFoundError("content control", name)
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


# ---------------------------------------------------------------------------
# Headings
# ---------------------------------------------------------------------------


def _paragraph_text(para: Any) -> str:
    """Heading text minus the trailing paragraph mark."""
    raw = str(para.Range.Text or "")
    return raw.rstrip("\r\n\x07")


def _find_heading_paragraph(doc_com: Any, name: str) -> tuple[Any, int] | None:
    """Locate a heading paragraph by visible text. Returns (Paragraph, 1-based index)."""
    for idx, para in enumerate(doc_com.Paragraphs, start=1):
        try:
            level = int(para.OutlineLevel)
        except Exception:
            continue
        if level >= 10:  # WdOutlineLevel: 1-9 are headings; 10 is body text
            continue
        if _paragraph_text(para) == name:
            return para, idx
    return None


def _section_range(doc_com: Any, target_para: Any, target_level: int) -> Any:
    """COM Range from the end of `target_para` to the next paragraph whose
    OutlineLevel is a heading and `<= target_level` — or to the end of the
    document's last paragraph if no such boundary exists.
    """
    paragraphs = list(doc_com.Paragraphs)
    target_start = int(target_para.Range.Start)

    idx: int | None = None
    for i, p in enumerate(paragraphs):
        try:
            if int(p.Range.Start) == target_start:
                idx = i
                break
        except Exception:
            continue
    if idx is None:
        end = int(target_para.Range.End)
        return doc_com.Range(end, end)

    section_start = int(target_para.Range.End)
    section_end: int | None = None
    for p in paragraphs[idx + 1:]:
        try:
            lvl = int(p.OutlineLevel)
        except Exception:
            continue
        if lvl < 10 and lvl <= target_level:
            section_end = int(p.Range.Start)
            break
    if section_end is None:
        try:
            section_end = int(paragraphs[-1].Range.End)
        except Exception:
            section_end = section_start
    return doc_com.Range(section_start, section_end)


class Heading(Anchor):
    kind = "heading"

    def _paragraph(self) -> Any:
        found = _find_heading_paragraph(self._doc.com, self.name)
        if found is None:
            raise AnchorNotFoundError("heading", self.name)
        return found[0]

    def _paragraph_and_index(self) -> tuple[Any, int]:
        """Default lookup goes by visible text; subclasses can override."""
        found = _find_heading_paragraph(self._doc.com, self.name)
        if found is None:
            raise AnchorNotFoundError("heading", self.name)
        return found

    @property
    def anchor_id(self) -> str:
        with _com.translate_com_errors():
            _, idx = self._paragraph_and_index()
        return f"heading:{idx}"

    @property
    def level(self) -> int:
        with _com.translate_com_errors():
            return int(self._paragraph().OutlineLevel)

    def section_range(self) -> Any:
        """COM Range covering the body under this heading.

        Spans from the end of the heading paragraph to the start of the next
        heading whose level is `<=` this one's (or to the end of the document
        if no such heading exists). Excludes the heading paragraph itself.
        """
        with _com.translate_com_errors():
            para = self._paragraph()
            level = int(para.OutlineLevel)
            return _section_range(self._doc.com, para, level)

    def section_text(self) -> str:
        """Plain text of the body under this heading."""
        with _com.translate_com_errors():
            return str(self.section_range().Text or "")

    def _range(self) -> Any:
        return self._paragraph().Range

    @property
    def text(self) -> str:
        with _com.translate_com_errors():
            return _paragraph_text(self._paragraph())

    def set_text(self, text: str) -> None:
        with _com.translate_com_errors():
            para_range = self._paragraph().Range
            start = int(para_range.Start)
            end = int(para_range.End)
            # Preserve the trailing paragraph mark.
            inner = self._doc.com.Range(start, max(start, end - 1))
            inner.Text = text

    def insert_paragraph_after(self, text: str, style: str | None = None) -> None:
        """Insert a new paragraph immediately after this heading."""
        with _com.translate_com_errors():
            doc_com = self._doc.com
            para_range = self._paragraph().Range
            end = int(para_range.End)
            insert_rng = doc_com.Range(end, end)
            insert_rng.Text = text + "\r"
            if style:
                styled = doc_com.Range(end, end + len(text))
                styled.Style = doc_com.Styles(style)


class _IndexedHeading(Heading):
    """A Heading located by 1-based paragraph index — used by anchor_by_id('heading:N').

    Disambiguates duplicate heading text. The display name is set to the resolved
    heading text the first time `_paragraph()` succeeds so error messages and
    `.name` reads stay informative.
    """

    def __init__(self, doc: "Document", paragraph_index: int) -> None:
        super().__init__(doc, name=f"heading:{paragraph_index}")
        self._paragraph_index = paragraph_index

    @property
    def anchor_id(self) -> str:
        return f"heading:{self._paragraph_index}"

    def _paragraph(self) -> Any:
        for idx, para in enumerate(self._doc.com.Paragraphs, start=1):
            if idx != self._paragraph_index:
                continue
            try:
                level = int(para.OutlineLevel)
            except Exception:
                break
            if level >= 10:
                break
            self.name = _paragraph_text(para) or self.name
            return para
        raise AnchorNotFoundError("heading", f"heading:{self._paragraph_index}")

    def _paragraph_and_index(self) -> tuple[Any, int]:
        return self._paragraph(), self._paragraph_index
