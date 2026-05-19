"""Word application wrapper + attach()/connect() context managers."""

from __future__ import annotations

from contextlib import contextmanager
from typing import Any, Iterator, TYPE_CHECKING

from . import _com
from .exceptions import WordNotRunningError

if TYPE_CHECKING:
    from ._document import DocumentCollection
    from ._selection import Selection


class Word:
    """Handle to a running Word.Application COM object."""

    def __init__(self, app: Any) -> None:
        self._app = app

    @property
    def com(self) -> Any:
        """Raw Application COM object — escape hatch when wordlive doesn't cover something."""
        return self._app

    @property
    def visible(self) -> bool:
        return bool(self._app.Visible)

    @visible.setter
    def visible(self, value: bool) -> None:
        with _com.translate_com_errors():
            self._app.Visible = bool(value)

    @property
    def documents(self) -> "DocumentCollection":
        from ._document import DocumentCollection

        return DocumentCollection(self)

    @property
    def selection(self) -> "Selection":
        from ._selection import Selection

        return Selection(self)


@contextmanager
def attach() -> Iterator[Word]:
    """Attach to an already-running Word instance.

    Raises `WordNotRunningError` if no instance is available. Does not launch
    Word and does not close it on exit.
    """
    with _com.com_apartment():
        app = _com.get_active_word()
        try:
            yield Word(app)
        finally:
            del app


@contextmanager
def connect(launch_if_missing: bool = True, visible: bool = True) -> Iterator[Word]:
    """Attach to a running Word, or launch a new one if missing.

    With `launch_if_missing=False` this behaves like `attach()`. Wordlive never
    closes Word on exit — even when it launched the instance itself, the user
    is expected to own its lifecycle.
    """
    with _com.com_apartment():
        try:
            app = _com.get_active_word()
        except WordNotRunningError:
            if not launch_if_missing:
                raise
            app = _com.launch_word(visible=visible)
        try:
            yield Word(app)
        finally:
            del app
