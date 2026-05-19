"""EditScope: atomic-undo + Selection preservation context manager."""

from __future__ import annotations

from typing import Any, TYPE_CHECKING

from . import _com, _selection
from ._selection import SelectionSnapshot

if TYPE_CHECKING:
    from ._app import Word


class EditScope:
    """Wraps a Word UndoRecord + a Selection snapshot.

    One `Ctrl-Z` reverts every mutation made inside the `with` block. The
    user's cursor and scroll position are restored on exit unless code inside
    the scope calls `allow_cursor_move()`.
    """

    def __init__(self, word: "Word", label: str) -> None:
        self._word = word
        self._label = label
        self._snapshot: SelectionSnapshot | None = None
        self._undo: Any | None = None
        self._move_allowed: bool = False
        self._undo_started: bool = False

    @property
    def word(self) -> "Word":
        return self._word

    @property
    def label(self) -> str:
        return self._label

    def allow_cursor_move(self) -> None:
        """Opt out of restoring the user's Selection on scope exit."""
        self._move_allowed = True

    def __enter__(self) -> "EditScope":
        self._snapshot = _selection.snapshot(self._word)
        with _com.translate_com_errors():
            self._undo = self._word.com.UndoRecord
            try:
                self._undo.StartCustomRecord(self._label)
                self._undo_started = True
            except Exception:
                # Older Word versions (<2010) lack UndoRecord. Continue without
                # atomic undo — operations still work, just not as one Ctrl-Z.
                self._undo_started = False
        return self

    def __exit__(self, exc_type: type[BaseException] | None, exc: BaseException | None, tb: Any) -> None:
        if self._undo_started and self._undo is not None:
            try:
                with _com.translate_com_errors():
                    self._undo.EndCustomRecord()
            except Exception:
                pass

        if exc_type is None and not self._move_allowed and self._snapshot is not None:
            try:
                _selection.restore(self._word, self._snapshot)
            except Exception:
                pass

        self._undo = None
