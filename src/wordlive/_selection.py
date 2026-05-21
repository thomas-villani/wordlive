"""Selection wrapper + snapshot/restore primitives for politeness preservation."""

from __future__ import annotations

from dataclasses import dataclass
from typing import TYPE_CHECKING, Any

from . import _com

if TYPE_CHECKING:
    from ._app import Word


@dataclass(frozen=True)
class SelectionSnapshot:
    """A point-in-time capture of where the user's cursor and view are."""

    start: int
    end: int
    vertical_percent: int | None = None
    """ActiveWindow.VerticalPercentScrolled at snapshot time, or None if unavailable."""


class Selection:
    """Wrapper around Application.Selection. Mostly used for reads."""

    def __init__(self, word: Word) -> None:
        self._word = word

    @property
    def com(self) -> Any:
        return self._word.com.Selection

    def info(self) -> dict[str, Any]:
        """Structured snapshot of the current selection for `wordlive` reads.

        `collapsed` is true when there's an insertion point but no selected
        text (`start == end`). The CLI's `cursor read` enriches this with the
        containing `para:N` anchor.
        """
        with _com.translate_com_errors():
            sel = self.com
            start = int(sel.Start)
            end = int(sel.End)
            return {
                "start": start,
                "end": end,
                "collapsed": start == end,
                "text": str(sel.Text or ""),
            }

    def write(self, text: str, *, replace: bool = True) -> None:
        """Insert `text` at the user's cursor — the deliberate cursor write.

        Unlike every anchor write, this targets the live `Selection`. With a
        spanning selection and `replace=True` (the default) the selected text is
        overwritten; with `replace=False`, or a collapsed cursor, the text is
        inserted at the selection start. Either way the cursor is left *after*
        the inserted text.

        This intentionally moves the cursor, so it fights `EditScope`'s
        cursor-preservation. To get atomic undo without snapping the cursor
        back, wrap it: ::

            with doc.edit("type at cursor") as scope:
                scope.allow_cursor_move()
                doc.selection.write("…")
        """
        with _com.translate_com_errors():
            sel = self.com
            start = int(sel.Start)
            end = int(sel.End)
            doc = self._word.com.ActiveDocument
            target = doc.Range(start, end if replace else start)
            target.Text = text
            # Collapse the cursor to just after the inserted text. Word counts
            # UTF-16 code units, so encode rather than using len().
            n = len(text.encode("utf-16-le")) // 2
            try:
                doc.Range(start + n, start + n).Select()
            except Exception:
                pass


def snapshot(word: Word) -> SelectionSnapshot:
    """Capture the user's current Selection and scroll position."""
    with _com.translate_com_errors():
        sel = word.com.Selection
        start = int(sel.Start)
        end = int(sel.End)

    vertical: int | None = None
    try:
        vertical = int(word.com.ActiveWindow.VerticalPercentScrolled)
    except Exception:
        vertical = None

    return SelectionSnapshot(start=start, end=end, vertical_percent=vertical)


def restore(word: Word, snap: SelectionSnapshot) -> None:
    """Best-effort restoration of a previous Selection + scroll position."""
    with _com.translate_com_errors():
        try:
            rng = word.com.ActiveDocument.Range(snap.start, snap.end)
            rng.Select()
        except Exception:
            # Range may now be out of bounds (document shrank). Fall back to collapsed start.
            try:
                rng = word.com.ActiveDocument.Range(snap.start, snap.start)
                rng.Select()
            except Exception:
                pass

    if snap.vertical_percent is not None:
        try:
            word.com.ActiveWindow.VerticalPercentScrolled = snap.vertical_percent
        except Exception:
            pass
