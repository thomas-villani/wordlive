"""Anchor types — semantic handles for ranges inside a Word document.

Anchors target a `Range`, never the live `Selection`. Each public mutation
goes through the COM error translator. Operations are intentionally small;
they compose with `Document.edit()` for atomic-undo behaviour.
"""

from __future__ import annotations

from abc import ABC, abstractmethod
from typing import Any, Iterator, TYPE_CHECKING

from . import _com
from .constants import WdParagraphAlignment
from .exceptions import AnchorNotFoundError

if TYPE_CHECKING:
    from ._document import Document


_ALIGNMENT_NAMES = {
    "left": WdParagraphAlignment.LEFT,
    "center": WdParagraphAlignment.CENTER,
    "centre": WdParagraphAlignment.CENTER,
    "right": WdParagraphAlignment.RIGHT,
    "justify": WdParagraphAlignment.JUSTIFY,
}


def _utf16_len(s: str) -> int:
    """Length of `s` in UTF-16 code units — Word's native character count.

    Python's `len()` counts code points, so astral-plane characters (emoji,
    historic scripts) count as 1. Word counts UTF-16 code units, so the same
    character counts as 2. Use this whenever the result is fed back into a
    Word `Range(start, end)` after a `Range.Text = ...` assignment.
    """
    return len(s.encode("utf-16-le")) // 2


def _coerce_alignment(value: Any) -> int:
    if isinstance(value, WdParagraphAlignment):
        return int(value)
    if isinstance(value, int) and not isinstance(value, bool):
        return value
    if isinstance(value, str):
        try:
            return int(_ALIGNMENT_NAMES[value.lower()])
        except KeyError:
            raise ValueError(
                f"unknown alignment {value!r}; expected one of "
                f"{sorted(set(_ALIGNMENT_NAMES))}"
            )
    raise TypeError(f"alignment must be WdParagraphAlignment, int, or str; got {type(value).__name__}")


# ---------------------------------------------------------------------------
# Base
# ---------------------------------------------------------------------------


class Anchor(ABC):
    """Abstract base — subclasses know how to materialise their COM Range.

    Concrete subclasses must implement `_range()` and `set_text()`. Other
    operations (`text`, `insert_before`, `insert_after`, `delete`,
    `apply_style`, `format_paragraph`) are derived and inherited as-is.
    """

    kind: str = "anchor"
    name: str = ""

    def __init__(self, doc: "Document", name: str) -> None:
        self._doc = doc
        self.name = name

    @property
    def com(self) -> Any:
        """Raw COM range. Subclasses override."""
        return self._range()

    @abstractmethod
    def _range(self) -> Any:
        """Return the COM Range that this anchor refers to. Must be overridden."""

    @property
    def text(self) -> str:
        with _com.translate_com_errors():
            return str(self._range().Text or "")

    @property
    @abstractmethod
    def anchor_id(self) -> str:
        """Stable string identifier for this anchor (e.g. `bookmark:Address`).

        Each anchor kind has its own scheme (`bookmark:`, `cc:`, `heading:`),
        so subclasses must declare theirs explicitly — no useful default
        exists at this level.
        """

    @abstractmethod
    def set_text(self, text: str) -> None:
        """Replace the anchor's text in place. Must be overridden."""

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

    def apply_style(self, name: str) -> None:
        """Apply the named paragraph or character style to this anchor's range.

        Word selects paragraph- vs. character-style behaviour from the style's
        own `Type`; we don't model that distinction. Raises `StyleNotFoundError`
        if the style isn't defined in the document.
        """
        style = self._doc.styles[name]  # raises StyleNotFoundError if missing
        with _com.translate_com_errors():
            self._range().Style = style.com

    def format_paragraph(
        self,
        *,
        alignment: Any = None,
        left_indent: float | None = None,
        right_indent: float | None = None,
        first_line_indent: float | None = None,
        space_before: float | None = None,
        space_after: float | None = None,
    ) -> None:
        """Set paragraph-formatting properties on this anchor's range.

        All kwargs are optional; only the ones explicitly passed are written.
        Indent and spacing values are in points (Word's native unit for
        `ParagraphFormat.LeftIndent` etc.). `alignment` accepts a
        `WdParagraphAlignment` enum, its int value, or a string
        (`"left"`/`"center"`/`"right"`/`"justify"`).
        """
        with _com.translate_com_errors():
            pf = self._range().ParagraphFormat
            if alignment is not None:
                pf.Alignment = _coerce_alignment(alignment)
            if left_indent is not None:
                pf.LeftIndent = float(left_indent)
            if right_indent is not None:
                pf.RightIndent = float(right_indent)
            if first_line_indent is not None:
                pf.FirstLineIndent = float(first_line_indent)
            if space_before is not None:
                pf.SpaceBefore = float(space_before)
            if space_after is not None:
                pf.SpaceAfter = float(space_after)

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
            # Word measures Range offsets in UTF-16 code units, not Python code points.
            new_end = start + _utf16_len(text)
            new_rng = doc_com.Range(start, new_end)
            doc_com.Bookmarks.Add(Name=self.name, Range=new_rng)


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

    def list(self, *, include_hidden: bool = False) -> list[str]:
        """Names of every user-visible bookmark in document order.

        Set `include_hidden=True` to also return Word's internal bookmarks
        (TOC entries, cross-references, etc.) whose names start with `_`.
        """
        with _com.translate_com_errors():
            names = [str(bm.Name) for bm in self._doc.com.Bookmarks]
        if include_hidden:
            return names
        return [n for n in names if _is_user_bookmark(n)]

    def __iter__(self) -> Iterator[Bookmark]:
        for name in self.list():
            yield Bookmark(self._doc, name)


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

    @property
    def anchor_id(self) -> str:
        return f"cc:{self.name}"

    def _cc(self) -> Any:
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


# ---------------------------------------------------------------------------
# Headings
# ---------------------------------------------------------------------------


def paragraph_text(para: Any) -> str:
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
            return paragraph_text(self._paragraph())

    def set_text(self, text: str) -> None:
        with _com.translate_com_errors():
            para_range = self._paragraph().Range
            start = int(para_range.Start)
            end = int(para_range.End)
            # Preserve the trailing paragraph mark.
            inner = self._doc.com.Range(start, max(start, end - 1))
            inner.Text = text

    def insert_paragraph_after(self, text: str, style: str | None = None) -> None:
        """Insert a new paragraph immediately after this heading.

        If `style` is given, it must name a style defined in the document;
        otherwise `StyleNotFoundError` is raised before any text is inserted.
        """
        # Validate the style up-front so a bad name raises StyleNotFoundError
        # before we mutate the document.
        if style is not None:
            style_obj = self._doc.styles[style]
        else:
            style_obj = None
        with _com.translate_com_errors():
            doc_com = self._doc.com
            para_range = self._paragraph().Range
            end = int(para_range.End)
            insert_rng = doc_com.Range(end, end)
            insert_rng.Text = text + "\r"
            if style_obj is not None:
                # Word measures Range offsets in UTF-16 code units; using
                # Python's len() under-counts surrogate pairs and leaves the
                # tail of the inserted paragraph un-styled.
                styled = doc_com.Range(end, end + _utf16_len(text))
                styled.Style = style_obj.com


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

    def __init__(self, doc: "Document") -> None:
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
            self.name = paragraph_text(para) or self.name
            return para
        raise AnchorNotFoundError("heading", f"heading:{self._paragraph_index}")

    def _paragraph_and_index(self) -> tuple[Any, int]:
        return self._paragraph(), self._paragraph_index
