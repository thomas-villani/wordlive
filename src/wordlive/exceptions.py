"""Public exception taxonomy for wordlive."""

from __future__ import annotations

from typing import Any


class WordliveError(Exception):
    """Base class for all wordlive errors."""


class WordNotRunningError(WordliveError):
    """No running Word instance is available."""


class DocumentNotFoundError(WordliveError):
    """The requested document is not open in Word."""

    def __init__(self, name: str) -> None:
        super().__init__(f"document not found: {name!r}")
        self.name = name


class AnchorNotFoundError(WordliveError):
    """The requested anchor (bookmark / content control / heading) does not exist."""

    def __init__(self, kind: str, name: str) -> None:
        super().__init__(f"{kind} not found: {name!r}")
        self.kind = kind
        self.name = name


class StyleNotFoundError(AnchorNotFoundError):
    """The requested paragraph or character style is not defined in the document.

    Subclass of `AnchorNotFoundError` so it shares the same exit code (2) and so
    `except AnchorNotFoundError` catches both bookmark-misses and style-misses.
    Retryable after re-reading `doc.styles.list()`.
    """

    def __init__(self, name: str) -> None:
        super().__init__("style", name)


class AmbiguousMatchError(WordliveError):
    """A find/replace pattern matched more than one occurrence without disambiguation.

    Carries the list of matches so callers (notably LLM drivers) can pick an
    `occurrence` index and retry.
    """

    def __init__(self, find: str, matches: list[dict[str, Any]]) -> None:
        super().__init__(
            f"{len(matches)} matches for {find!r}; pass --all or --occurrence N to disambiguate"
        )
        self.find = find
        self.matches = matches


class ImageSourceError(WordliveError):
    """An image given to `insert_image` couldn't be turned into an embeddable file.

    Raised for a missing or unreadable path, malformed base64, or bytes whose
    format isn't a recognised raster image (PNG/JPEG/GIF/BMP/TIFF). It's a
    bad-input error — not a "named thing is missing" — so it maps to the
    generic exit code (1) rather than reusing the anchor-not-found code.
    Not retryable: fix the input.
    """

    def __init__(self, message: str) -> None:
        super().__init__(message)


class SnapshotError(WordliveError):
    """A page/section snapshot couldn't be rendered.

    Raised when the optional PDF-rendering backend (PyMuPDF) isn't installed, or
    when rasterising the exported PDF fails. The PDF export itself goes through
    Word's COM, so a busy/modal Word surfaces as `WordBusyError`, not this. It's
    an environment/dependency problem rather than a "named thing is missing", so
    it maps to the generic exit code (1). Fix by installing the extra:
    `pip install "wordlive[snapshot]"` (or `uv add "wordlive[snapshot]"`).
    """

    def __init__(self, message: str) -> None:
        super().__init__(message)


class WordBusyError(WordliveError):
    """Word rejected the RPC — typically a modal dialog or a transient busy state.

    Retryable in principle; caller decides.
    """

    def __init__(
        self, message: str = "Word is busy or in a modal dialog", *, hresult: int | None = None
    ) -> None:
        super().__init__(message)
        self.hresult = hresult
        self.retryable = True


class ComError(WordliveError):
    """Generic wrapper for an unclassified pywintypes.com_error."""

    def __init__(
        self, message: str, *, hresult: int | None = None, description: str | None = None
    ) -> None:
        super().__init__(message)
        self.hresult = hresult
        self.description = description


# HRESULTs we recognise as "Word is momentarily unavailable" rather than a real error.
_BUSY_HRESULTS: frozenset[int] = frozenset(
    {
        0x80010001,  # RPC_E_CALL_REJECTED — call rejected by callee (modal dialog, busy)
        0x8001010A,  # RPC_E_SERVERCALL_RETRYLATER — server busy, retry later
        0x80010005,  # RPC_E_SERVERCALL_REJECTED — server rejected the call
        -2147418111,  # signed form of RPC_E_CALL_REJECTED
        -2147417846,  # signed form of RPC_E_SERVERCALL_RETRYLATER
    }
)


def _decode_com_error(exc: Any) -> tuple[int | None, str | None, str]:
    """Pull (hresult, description, readable_message) out of a pywintypes.com_error.

    pywintypes.com_error.args is (hresult, message, exc_info, arg_err) where exc_info,
    when present, is (wcode, source, description, helpfile, helpcontext, scode).
    """
    args: tuple[Any, ...] = getattr(exc, "args", ()) or ()
    hresult: int | None = None
    description: str | None = None
    message = str(exc)

    if len(args) >= 1 and isinstance(args[0], int):
        hresult = args[0]
    if len(args) >= 3 and args[2]:
        exc_info = args[2]
        try:
            description = exc_info[2] if len(exc_info) > 2 else None
            scode = exc_info[5] if len(exc_info) > 5 else None
        except (TypeError, IndexError):
            description, scode = None, None
        if scode is not None and hresult is None:
            hresult = scode

    parts = []
    if description:
        parts.append(description.strip())
    if hresult is not None:
        parts.append(f"HRESULT 0x{hresult & 0xFFFFFFFF:08X}")
    if parts:
        message = " — ".join(parts)
    return hresult, description, message


def from_com_error(exc: Any) -> WordliveError:
    """Classify a pywintypes.com_error into the appropriate wordlive exception."""
    hresult, description, message = _decode_com_error(exc)
    if hresult is not None and hresult in _BUSY_HRESULTS:
        return WordBusyError(message, hresult=hresult)
    return ComError(message, hresult=hresult, description=description)
