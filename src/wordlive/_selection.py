"""Selection wrapper + snapshot/restore primitives for politeness preservation."""

from __future__ import annotations

from dataclasses import dataclass
from typing import Any, TYPE_CHECKING

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

    def __init__(self, word: "Word") -> None:
        self._word = word

    @property
    def com(self) -> Any:
        return self._word.com.Selection

    def info(self) -> dict[str, Any]:
        """Structured snapshot of the current selection for `wordlive` reads."""
        with _com.translate_com_errors():
            sel = self.com
            return {
                "start": int(sel.Start),
                "end": int(sel.End),
                "text": str(sel.Text or ""),
            }


def snapshot(word: "Word") -> SelectionSnapshot:
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


def restore(word: "Word", snap: SelectionSnapshot) -> None:
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
