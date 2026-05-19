"""Thin COM helpers — the mockable seam between wordlive and pywin32.

Tests substitute fakes for `get_active_word` / `launch_word` via monkeypatch.
Everything else in wordlive only sees duck-typed dispatch objects.
"""

from __future__ import annotations

from contextlib import contextmanager
from typing import Any, Iterator

from .exceptions import WordNotRunningError, from_com_error


_WORD_PROG_ID = "Word.Application"


@contextmanager
def com_apartment() -> Iterator[None]:
    """STA apartment lifecycle. Nests safely via pythoncom's reference counting."""
    try:
        import pythoncom  # type: ignore[import-not-found]
    except ImportError:
        # Non-Windows or pywin32 missing: yield without initialising. The first
        # real COM call will fail with a clearer error.
        yield
        return

    pythoncom.CoInitialize()
    try:
        yield
    finally:
        pythoncom.CoUninitialize()


def get_active_word() -> Any:
    """Return the Application COM object for an already-running Word, or raise."""
    try:
        from win32com.client import GetActiveObject  # type: ignore[import-not-found]
    except ImportError as e:
        raise WordNotRunningError(
            "pywin32 is not installed; wordlive requires Windows + pywin32"
        ) from e

    try:
        return GetActiveObject(_WORD_PROG_ID)
    except Exception as e:  # pywintypes.com_error or similar
        raise WordNotRunningError(
            "no running Microsoft Word instance found"
        ) from e


def launch_word(visible: bool = True) -> Any:
    """Launch a new Word instance and return its Application COM object."""
    try:
        from win32com.client import Dispatch  # type: ignore[import-not-found]
    except ImportError as e:
        raise WordNotRunningError(
            "pywin32 is not installed; wordlive requires Windows + pywin32"
        ) from e

    app = Dispatch(_WORD_PROG_ID)
    try:
        app.Visible = bool(visible)
    except Exception:
        # Some COM stubs may not let us flip Visible immediately; not fatal.
        pass
    return app


@contextmanager
def translate_com_errors() -> Iterator[None]:
    """Translate pywintypes.com_error into wordlive's typed exceptions."""
    try:
        import pywintypes  # type: ignore[import-not-found]

        com_error_type: type = pywintypes.com_error
    except ImportError:
        com_error_type = ()  # type: ignore[assignment]

    try:
        yield
    except Exception as exc:  # noqa: BLE001
        if com_error_type and isinstance(exc, com_error_type):  # type: ignore[arg-type]
            raise from_com_error(exc) from exc
        raise
