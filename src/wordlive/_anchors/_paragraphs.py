"""`para:N` anchors and the paragraph collection."""

from __future__ import annotations

from collections.abc import Iterator
from typing import TYPE_CHECKING, Any

from .. import _com
from ..exceptions import AnchorNotFoundError

if TYPE_CHECKING:
    from .._document import Document

from ._base import Anchor
from ._helpers import (
    paragraph_text,
    range_text,
)
from ._refs import (
    _stale_anchor_hint,
)

# ---------------------------------------------------------------------------
# Paragraphs
# ---------------------------------------------------------------------------


class Paragraph(Anchor):
    """A paragraph located by 1-based index over `doc.Paragraphs`.

    `para:N` addresses *any* paragraph — body text, headings, list items alike.
    `heading:N` is the same index space narrowed to heading paragraphs, so
    `para:5` and `heading:5` resolve to the same paragraph when paragraph 5 is a
    heading. A `Paragraph` inherits every anchor verb (`set_text`, `apply_style`,
    `format_paragraph`, `apply_list`, `insert_paragraph_before/after`, …).
    """

    kind = "paragraph"

    def __init__(self, doc: Document, index: int) -> None:
        super().__init__(doc, name=f"para:{index}")
        self._index = index

    @property
    def anchor_id(self) -> str:
        return f"para:{self._index}"

    @property
    def index(self) -> int:
        return self._index

    def _paragraph(self) -> Any:
        for idx, para in enumerate(self._doc.com.Paragraphs, start=1):
            if idx == self._index:
                # Keep .name informative for repr / error messages.
                self.name = paragraph_text(para) or self.name
                return para
        raise AnchorNotFoundError(
            "paragraph",
            f"para:{self._index}",
            hint=_stale_anchor_hint(self._doc.com, "para", self._index),
        )

    @property
    def level(self) -> int:
        with _com.translate_com_errors():
            return int(self._paragraph().OutlineLevel)

    @property
    def is_heading(self) -> bool:
        return self.level < 10

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
            # Preserve the trailing paragraph mark so the paragraph isn't merged
            # with the next one (same approach as Heading.set_text).
            inner = self._doc.com.Range(start, max(start, end - 1))
            inner.Text = text


class ParagraphCollection:
    """Indexable, iterable view over every paragraph in the document.

    Unlike `headings`, this includes body paragraphs and list items, not just
    heading paragraphs. Index by 1-based position (`doc.paragraphs[2]`); iterate
    for a `Paragraph` per paragraph. `list()` emits each paragraph's `start` /
    `end` offsets, so a body paragraph can be turned into a `range:START-END`
    insertion point for mid-paragraph edits.
    """

    def __init__(self, doc: Document) -> None:
        self._doc = doc

    def _count(self) -> int:
        with _com.translate_com_errors():
            return sum(1 for _ in self._doc.com.Paragraphs)

    def __len__(self) -> int:
        return self._count()

    def __getitem__(self, index: int) -> Paragraph:
        if isinstance(index, bool) or not isinstance(index, int):
            raise TypeError(f"paragraph index must be int, got {type(index).__name__}")
        if index < 1 or index > self._count():
            raise AnchorNotFoundError(
                "paragraph",
                f"para:{index}",
                hint=_stale_anchor_hint(self._doc.com, "para", index),
            )
        return Paragraph(self._doc, index)

    def __iter__(self) -> Iterator[Paragraph]:
        with _com.translate_com_errors():
            count = sum(1 for _ in self._doc.com.Paragraphs)
        for idx in range(1, count + 1):
            yield Paragraph(self._doc, idx)

    def at(self, offset: int) -> Paragraph | None:
        """Return the paragraph whose range contains `offset`, or None.

        Used to map a character offset (e.g. the cursor position) back to a
        `para:N` anchor.
        """
        with _com.translate_com_errors():
            for idx, para in enumerate(self._doc.com.Paragraphs, start=1):
                rng = para.Range
                if int(rng.Start) <= offset < int(rng.End):
                    return Paragraph(self._doc, idx)
        return None

    def list(self) -> list[dict[str, Any]]:
        """Every paragraph as `[{index, anchor_id, level, style, is_heading, start, end, text}, ...]`.

        `style` is the paragraph's applied Word style name (e.g. ``"Normal"``,
        ``"List Number"``, ``"Heading 2"``) — the handle to feed back into
        `apply_style` / a write's `style=` to mirror existing formatting, since
        `level` (Word's `OutlineLevel`) is `10` for *all* non-heading paragraphs
        and so can't distinguish a list item from body text. It's `None` if the
        style can't be read.
        """
        out: list[dict[str, Any]] = []
        with _com.translate_com_errors():
            # Ask the document once whether any inline shape exists. If none does, no
            # paragraph range can hold one, and each row's text read skips fetching
            # (and wrapping) a per-paragraph `InlineShapes` collection. On failure,
            # assume they may exist and let the per-range probe decide.
            try:
                may_have_shapes = int(self._doc.com.InlineShapes.Count) > 0
            except Exception:
                may_have_shapes = True
            for idx, para in enumerate(self._doc.com.Paragraphs, start=1):
                # Fetch `para.Range` once. Each access mints a fresh COM object that
                # pywin32 must wrap (a QueryInterface + type lookup), which dominates
                # this walk — three fetches per paragraph made it ~3x its true cost.
                rng = para.Range
                try:
                    level = int(para.OutlineLevel)
                except Exception:
                    level = 10
                try:
                    style: str | None = str(rng.Style.NameLocal)
                except Exception:
                    style = None
                out.append(
                    {
                        "index": idx,
                        "anchor_id": f"para:{idx}",
                        "level": level,
                        "style": style,
                        "is_heading": level < 10,
                        "start": int(rng.Start),
                        "end": int(rng.End),
                        "text": range_text(rng, may_have_shapes=may_have_shapes).rstrip("\r\n\x07"),
                    }
                )
        return out
