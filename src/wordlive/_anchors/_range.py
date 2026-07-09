"""Offset-addressed anchors: `range:START-END`, plus the `start` / `end` sentinels."""

from __future__ import annotations

from typing import TYPE_CHECKING, Any

from .. import _com

if TYPE_CHECKING:
    from .._document import Document

from ._base import Anchor
from ._helpers import (
    _utf16_len,
)

# ---------------------------------------------------------------------------
# Arbitrary ranges
# ---------------------------------------------------------------------------


class RangeAnchor(Anchor):
    """An anchor over an arbitrary character range — `doc.range(start, end)`.

    Unlike bookmarks/headings/cells, a range anchor names nothing in the
    document: it's a pair of absolute character offsets (UTF-16 code units, the
    same coordinates Word's `Document.Range(start, end)` uses and that
    `Document.find()` emits as `range:START-END`). It's the generic target when
    no named anchor exists — feed a `find()` hit straight into a `replace`, or
    drop a comment on an offset span.

    The anchor is ephemeral: offsets resolve live against the document on each
    access, so an edit elsewhere that shifts the text can leave it pointing at
    the wrong span. Resolve, act, discard. `set_text` keeps the anchor's own
    `end` in sync with the replacement so chained ops on the same instance stay
    consistent.
    """

    kind = "range"

    def __init__(self, doc: Document, start: int, end: int) -> None:
        start = int(start)
        end = int(end)
        if start < 0 or end < start:
            raise ValueError(f"invalid range offsets: start={start}, end={end}")
        super().__init__(doc, name=f"range:{start}-{end}")
        self._start = start
        self._end = end

    @property
    def start(self) -> int:
        return self._start

    @property
    def end(self) -> int:
        return self._end

    @property
    def anchor_id(self) -> str:
        return f"range:{self._start}-{self._end}"

    def _range(self) -> Any:
        return self._doc.com.Range(self._start, self._end)

    def set_text(self, text: str) -> None:
        with _com.translate_com_errors():
            rng = self._doc.com.Range(self._start, self._end)
            rng.Text = text
        # A Range.Text assignment resizes the span; keep our end in sync so a
        # follow-up read/op on the same anchor sees the replacement rather than
        # the stale coordinates. Word counts UTF-16 code units, not code points.
        self._end = self._start + _utf16_len(text)


# ---------------------------------------------------------------------------
# Start / end of document
# ---------------------------------------------------------------------------


class StartAnchor(Anchor):
    """A zero-width anchor at the very start of the document body — `doc.start`.

    The mirror of [`EndAnchor`][wordlive.EndAnchor]: the insertion point before
    the first paragraph. `doc.start` returns it and `anchor_by_id("start")`
    resolves it, so "prepend to the document" composes with the usual verbs and
    the CLI `--anchor-id` plumbing.

    Only the *prepend* direction is meaningful at a single start-point, so every
    insert verb lands text at the start: `insert_paragraph_before` /
    `insert_paragraph_after` add a new first paragraph (delegating to
    [`Document.prepend_paragraph`][wordlive.Document.prepend_paragraph]), and
    `insert_before` / `insert_after` / `set_text` prepend inline (delegating to
    [`Document.prepend`][wordlive.Document.prepend]). `text` is always empty and
    `delete()` is a no-op. `insert_image` and `apply_style` are inherited: they
    resolve to the collapsed start position.
    """

    kind = "start"

    def __init__(self, doc: Document) -> None:
        super().__init__(doc, name="start")

    @property
    def anchor_id(self) -> str:
        return "start"

    def _range(self) -> Any:
        # Collapsed at offset 0 — the position Document.prepend* writes to.
        return self._doc.com.Range(0, 0)

    def set_text(self, text: str) -> None:
        # Nothing to replace at the start-point — prepend instead.
        self._doc.prepend(text)

    def insert_after(self, text: str) -> None:
        self._doc.prepend(text)

    def insert_before(self, text: str) -> None:
        # A single start-point has no distinct "after"; prepending is the only
        # sensible reading, and it keeps `--anchor-id start` honest either way.
        self._doc.prepend(text)

    def insert_paragraph_after(self, text: str, style: str | None = None) -> None:
        self._doc.prepend_paragraph(text, style=style)

    def insert_paragraph_before(self, text: str, style: str | None = None) -> None:
        self._doc.prepend_paragraph(text, style=style)


class EndAnchor(Anchor):
    """A zero-width anchor at the very end of the document body — `doc.end`.

    The one position no content names: the insertion point past the last
    paragraph. `doc.end` returns it and `anchor_by_id("end")` resolves it, so
    "append to the document" composes with the same verbs and the same CLI
    `--anchor-id` plumbing as every other anchor — no `.com` drop needed.

    Only the *append* direction is meaningful at a single end-point, so every
    insert verb lands text at the end: `insert_paragraph_after` /
    `insert_paragraph_before` add a new final paragraph (delegating to
    [`Document.append_paragraph`][wordlive.Document.append_paragraph]), and
    `insert_after` / `insert_before` / `set_text` append inline (delegating to
    [`Document.append`][wordlive.Document.append]). `text` is always empty and
    `delete()` is a no-op — there is no content here to read or remove.
    `insert_image` and `apply_style` are inherited: they resolve to the
    collapsed end position, so an image lands at the end and a style falls on
    the final paragraph.
    """

    kind = "end"

    def __init__(self, doc: Document) -> None:
        super().__init__(doc, name="end")

    @property
    def anchor_id(self) -> str:
        return "end"

    def _range(self) -> Any:
        # Collapsed just before the final paragraph mark — the position
        # Document.append* writes to, and a safe target for the inherited verbs
        # (a zero-width span reads "" and deletes nothing).
        with _com.translate_com_errors():
            end = int(self._doc.com.Content.End)
        pos = max(0, end - 1)
        return self._doc.com.Range(pos, pos)

    def set_text(self, text: str) -> None:
        # Nothing to replace at the end-point — append instead.
        self._doc.append(text)

    def insert_after(self, text: str) -> None:
        self._doc.append(text)

    def insert_before(self, text: str) -> None:
        # A single end-point has no distinct "before"; appending is the only
        # sensible reading, and it keeps `--anchor-id end` honest either way.
        self._doc.append(text)

    def insert_paragraph_after(self, text: str, style: str | None = None) -> None:
        self._doc.append_paragraph(text, style=style)

    def insert_paragraph_before(self, text: str, style: str | None = None) -> None:
        self._doc.append_paragraph(text, style=style)
