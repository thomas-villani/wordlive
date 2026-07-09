"""Reads over an anchor's range: revision-aware text, snapshot, location."""

from __future__ import annotations

from typing import TYPE_CHECKING, Any

from .. import _com, _images, _revisions
from ..constants import (
    WdInformation,
)

if TYPE_CHECKING:
    from pathlib import Path

    from .._snapshot import Snapshot

from ._helpers import (
    range_text,
)

if TYPE_CHECKING:
    pass

from ._anchor_core import AnchorCore


class AnchorReadMixin(AnchorCore):
    """Reads over an anchor's range: revision-aware text, snapshot, location."""

    def _revision_runs(self, start: int, end: int) -> list[dict[str, Any]]:
        """`{change, start, end, text}` for each insert/delete revision overlapping `[start, end)`."""
        runs: list[dict[str, Any]] = []
        for row in self._doc.revisions.list():
            change = row["type"]
            if change not in ("insert", "delete"):
                continue
            r_start, r_end = int(row["start"]), int(row["end"])
            if r_end <= start or r_start >= end:
                continue
            runs.append({"change": change, "start": r_start, "end": r_end, "text": row["text"]})
        return runs

    def revision_segments(self) -> list[dict[str, Any]]:
        """The anchor's text split into tracked-change segments (revision-aware read).

        Returns `[{text, change}, …]` in document order, where `change` is
        ``"insert"``, ``"delete"``, or ``None`` (unchanged). Word's `text` read
        shows the *final* view (inserted runs present, deleted runs gone); this
        also surfaces the deleted runs, so you can see both sides of a tracked
        edit. [`text_final`][wordlive.Anchor.text_final] and
        [`text_original`][wordlive.Anchor.text_original] are the two flattened
        views. The structured, whole-document counterpart is `doc.revisions`.
        """
        with _com.translate_com_errors():
            rng = self._range()
            start, end = int(rng.Start), int(rng.End)
            final_text = range_text(rng)
        return _revisions.segment_runs(final_text, start, self._revision_runs(start, end))

    @property
    def text_final(self) -> str:
        """The anchor's text **as if every tracked change in it were accepted**.

        Inserted runs stay, deleted runs drop — the after-the-edits view. Equal to
        `text` when nothing tracked touches the range. The mirror is
        [`text_original`][wordlive.Anchor.text_original]; the per-segment breakdown
        is [`revision_segments`][wordlive.Anchor.revision_segments].
        """
        return "".join(s["text"] for s in self.revision_segments() if s["change"] != "delete")

    @property
    def text_original(self) -> str:
        """The anchor's text **as if every tracked change in it were rejected**.

        Deleted runs stay, inserted runs drop — the before-the-edits view. The
        mirror of [`text_final`][wordlive.Anchor.text_final].
        """
        return "".join(s["text"] for s in self.revision_segments() if s["change"] != "insert")

    def snapshot(
        self, out: str | Path | None = None, *, dpi: int = 150, max_dim: int | None = None
    ) -> list[Snapshot]:
        """Render the page(s) this anchor sits on to PNG — let a model *see* it.

        A heading expands to its whole section; any other anchor renders the
        page(s) its range spans. Returns a list of
        [`Snapshot`][wordlive.Snapshot] (one per page); pass `out` to also write
        the image(s) to disk. `max_dim` caps each page's long edge in pixels (for
        a cheaper render). Sugar for
        [`Document.snapshot_anchor`][wordlive.Document.snapshot_anchor]; see it
        for the full semantics. Requires the `snapshot` extra (PyMuPDF).
        """
        return self._doc.snapshot_anchor(self._as_anchor, out, dpi=dpi, max_dim=max_dim)

    def read_image(self) -> tuple[bytes, str]:
        """Extract the image embedded in this anchor's range as `(bytes, mime_type)`.

        The read side of the image story — pull an embedded picture's original
        bytes back out (e.g. to hand to a vision model), the counterpart to
        [`insert_image`][wordlive.Anchor.insert_image]. The range must contain
        exactly one picture: an [`image:N`][wordlive.ImageAnchor] anchor (or any
        single-image text anchor) reads cleanly, while a range with no image — or
        more than one — raises `ImageSourceError`. `bytes` is the picture's raw
        encoded data (PNG/JPEG/…); `mime_type` is its content type
        (``"image/png"``, ``"image/jpeg"``, …). Discover what's there first with
        [`doc.images`][wordlive.Document.images]. Read-only — nothing is mutated.
        """
        with _com.translate_com_errors():
            return _images.read_image_from_range(self._range())

    def location(self) -> dict[str, Any]:
        """Where this anchor sits in the laid-out document — a pure read.

        Returns `{page, end_page, line, column, in_table}`:

        - `page` / `end_page` — the 1-based pages the anchor's **first** and
          **last** characters fall on (equal for a collapsed/single-line anchor);
          the pair is the anchor's *page span*, so a section/table/image that
          straddles a page boundary reports both. `page` is what answers "what
          page is this on"; scan `paragraphs` and watch `page` step up to find
          "which paragraph starts page 2".
        - `line` / `column` — the first character's 1-based line and column in
          the page's text grid (`Range.Information`).
        - `in_table` — whether the anchor sits inside a table.

        Page/line numbers are only meaningful in print layout, so the document
        is **repaginated first** (content-neutral — it touches neither the
        user's selection, scroll, nor view), mirroring the guarantee a
        `snapshot` gives. No politeness concern: this mutates nothing — the
        document's `Saved` state is snapshotted and restored around the
        repaginate, which would otherwise flip Word's dirty bit.
        """
        with _com.translate_com_errors(), _com.preserve_saved(self._doc.com):
            rng = self._range()
            self._doc.com.Repaginate()
            start, end = int(rng.Start), int(rng.End)
            doc_com = self._doc.com
            head = doc_com.Range(start, start)
            tail = doc_com.Range(end, end)
            return {
                "page": int(head.Information(int(WdInformation.ACTIVE_END_PAGE_NUMBER))),
                "end_page": int(tail.Information(int(WdInformation.ACTIVE_END_PAGE_NUMBER))),
                "line": int(head.Information(int(WdInformation.FIRST_CHARACTER_LINE_NUMBER))),
                "column": int(head.Information(int(WdInformation.FIRST_CHARACTER_COLUMN_NUMBER))),
                "in_table": bool(rng.Information(int(WdInformation.WITH_IN_TABLE))),
            }
