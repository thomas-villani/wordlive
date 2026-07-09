"""`heading:N` anchors and the heading collection (the outline addressing scheme)."""

from __future__ import annotations

from collections.abc import Iterator
from typing import TYPE_CHECKING, Any

from .. import _com
from ..exceptions import AnchorNotFoundError, OpError

if TYPE_CHECKING:
    from .._document import Document

from ._base import Anchor
from ._helpers import (
    paragraph_text,
)
from ._range import RangeAnchor
from ._refs import (
    _stale_anchor_hint,
)


def _find_heading_paragraph(doc_com: Any, name: str) -> tuple[Any, int] | None:
    """Locate a heading paragraph by visible text. Returns (Paragraph, 1-based index)."""
    for idx, para in enumerate(doc_com.Paragraphs, start=1):
        try:
            level = int(para.OutlineLevel)
        except Exception:
            continue
        if level >= 10:  # WdOutlineLevel: 1-9 are headings; 10 is body text
            continue
        if paragraph_text(para) == name:
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
    for p in paragraphs[idx + 1 :]:
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

    def replace_section_body(self, body: Any, *, markdown: bool = False) -> RangeAnchor:
        """Rewrite this heading's body, leaving the heading paragraph intact.

        The "rewrite section X" workflow: clears the span under this heading
        (`section_range`, up to the next same-or-higher heading) and inserts
        `body` after the heading. With ``markdown=False`` (default) `body` is the
        `insert_block` items shape (or a bare string); with ``markdown=True``
        `body` is a constrained-Markdown string routed through `insert_markdown`.
        Returns the new body's spanning [`RangeAnchor`][wordlive.RangeAnchor].
        Wrap in `doc.edit(...)` for atomic undo.
        """
        with _com.translate_com_errors():
            span = self.section_range()
            doc_com = self._doc.com
            doc_com.Range(int(span.Start), int(span.End)).Delete()
        if markdown:
            if not isinstance(body, str):
                raise OpError("replace_section_body with markdown=True requires a string body")
            return self.insert_markdown(body, where="after")
        if isinstance(body, str):
            body = [body]
        if not isinstance(body, list):
            raise OpError(
                f"replace_section_body body must be a string or list; got {type(body).__name__}"
            )
        return self.insert_block(body, where="after")

    def _range(self) -> Any:
        return self._paragraph().Range

    @property
    def text(self) -> str:
        with _com.translate_com_errors():
            return paragraph_text(self._paragraph())

    def set_text(self, text: str) -> None:
        with _com.translate_com_errors():
            para_range = self._paragraph().Range
            start = int(para_range.Start)
            end = int(para_range.End)
            # Preserve the trailing paragraph mark.
            inner = self._doc.com.Range(start, max(start, end - 1))
            inner.Text = text

    # insert_paragraph_after / insert_paragraph_before are inherited from Anchor;
    # for a Heading, _range() is the heading paragraph, so "after" lands a new
    # paragraph just below the heading (the original v0 behaviour).


class HeadingCollection:
    """Iterable, indexable view over a document's headings.

    Symmetric with `BookmarkCollection` and `ContentControlCollection`:

        for h in doc.headings:           # iteration → Heading per heading paragraph
            ...
        doc.headings["Risks"]            # by visible text
        doc.headings[3]                  # by 1-based paragraph index
        "Risks" in doc.headings          # membership
        doc.headings.list()              # same shape as doc.outline()
    """

    def __init__(self, doc: Document) -> None:
        self._doc = doc

    def __getitem__(self, key: str | int) -> Heading:
        if isinstance(key, bool):
            # bool is a subclass of int; reject before the int branch matches.
            raise TypeError(f"heading key must be str or int, got {type(key).__name__}")
        if isinstance(key, int):
            return _IndexedHeading(self._doc, key)
        if isinstance(key, str):
            with _com.translate_com_errors():
                if _find_heading_paragraph(self._doc.com, key) is None:
                    raise AnchorNotFoundError("heading", key)
            return Heading(self._doc, key)
        raise TypeError(f"heading key must be str or int, got {type(key).__name__}")

    def __contains__(self, key: object) -> bool:
        if isinstance(key, bool):
            return False
        if isinstance(key, int):
            # 1-based paragraph index must reference an actual heading paragraph.
            with _com.translate_com_errors():
                for idx, para in enumerate(self._doc.com.Paragraphs, start=1):
                    if idx != key:
                        continue
                    try:
                        lvl = int(para.OutlineLevel)
                    except Exception:
                        return False
                    return lvl < 10
            return False
        if not isinstance(key, str):
            return False
        with _com.translate_com_errors():
            return _find_heading_paragraph(self._doc.com, key) is not None

    def list(self) -> list[dict[str, Any]]:
        """Same shape as `Document.outline()` — `[{level, text, anchor_id}, ...]`."""
        out: list[dict[str, Any]] = []
        with _com.translate_com_errors():
            for idx, para in enumerate(self._doc.com.Paragraphs, start=1):
                try:
                    level = int(para.OutlineLevel)
                except Exception:
                    continue
                if level >= 10:
                    continue
                out.append(
                    {
                        "level": level,
                        "text": paragraph_text(para),
                        "anchor_id": f"heading:{idx}",
                    }
                )
        return out

    def __iter__(self) -> Iterator[Heading]:
        for entry in self.list():
            # Each entry's anchor_id is `heading:N`; index-based heading
            # disambiguates duplicate visible text.
            idx = int(entry["anchor_id"].split(":", 1)[1])
            yield _IndexedHeading(self._doc, idx)


class _IndexedHeading(Heading):
    """A Heading located by 1-based paragraph index — used by anchor_by_id('heading:N').

    Disambiguates duplicate heading text. The display name is set to the resolved
    heading text the first time `_paragraph()` succeeds so error messages and
    `.name` reads stay informative.
    """

    def __init__(self, doc: Document, paragraph_index: int) -> None:
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
            self.name = paragraph_text(para) or self.name
            return para
        raise AnchorNotFoundError(
            "heading",
            f"heading:{self._paragraph_index}",
            hint=_stale_anchor_hint(self._doc.com, "heading", self._paragraph_index),
        )

    def _paragraph_and_index(self) -> tuple[Any, int]:
        return self._paragraph(), self._paragraph_index
