"""Image helpers — get images *into* a document (path/bytes/base64 → disk) and
*out* of one (an embedded picture → raw bytes + MIME type).

The **in** path: `insert_image` needs an actual file path for Word's
`InlineShapes.AddPicture`, which only reads from disk. `image_on_disk` normalises
the three input shapes an LLM might hold, sniffs the raster format from magic
bytes (to pick a sensible temp-file extension), and removes any temp file once
Word has embedded a copy.

The **out** path: `read_image_from_range` extracts the picture embedded in a
range as raw bytes + MIME type — the read side that feeds a vision model. It goes
through `Range.WordOpenXML`, which serialises the range as a Flat OPC package with
every referenced media part inlined as base64; we parse that, take the sole
`image/*` part, and base64-decode it. No clipboard, no save-to-temp, no fragile
position→media mapping — pure stdlib.
"""

from __future__ import annotations

import base64
import binascii
import os
import tempfile
import xml.etree.ElementTree as ET
from collections.abc import Iterator
from contextlib import contextmanager
from pathlib import Path
from typing import Any

from .exceptions import ImageSourceError

# The Flat OPC package namespace. `Range.WordOpenXML` emits a
# `<pkg:package>` whose `<pkg:part>` children carry a `pkg:contentType` and,
# for binary parts (images), a base64 `<pkg:binaryData>` body.
_PKG_NS = "http://schemas.microsoft.com/office/2006/xmlPackage"
_PKG_PART = f"{{{_PKG_NS}}}part"
_PKG_CONTENT_TYPE = f"{{{_PKG_NS}}}contentType"
_PKG_BINARY_DATA = f"{{{_PKG_NS}}}binaryData"

# (magic-number prefix, extension) in match order. Raster formats only; vector
# types (EMF/WMF) and anything unrecognised fall through to ImageSourceError.
_MAGIC: tuple[tuple[bytes, str], ...] = (
    (b"\x89PNG\r\n\x1a\n", ".png"),
    (b"\xff\xd8\xff", ".jpg"),
    (b"GIF87a", ".gif"),
    (b"GIF89a", ".gif"),
    (b"BM", ".bmp"),
    (b"II*\x00", ".tif"),
    (b"MM\x00*", ".tif"),
)


def _sniff_extension(data: bytes) -> str:
    """Return a file extension (with leading dot) for `data`.

    Word detects the real format from the file's content, but a matching
    extension avoids surprises — so we only accept bytes we recognise and
    raise `ImageSourceError` otherwise.
    """
    for prefix, ext in _MAGIC:
        if data.startswith(prefix):
            return ext
    raise ImageSourceError("unrecognised image data: expected PNG, JPEG, GIF, BMP, or TIFF")


def _decode_base64(text: str) -> bytes:
    """Decode a base64 string, tolerating a `data:` URL prefix and whitespace."""
    s = text.strip()
    if s.startswith("data:"):
        _, _, s = s.partition(",")  # drop a `data:image/png;base64,` prefix
    try:
        return base64.b64decode(s, validate=True)
    except (binascii.Error, ValueError) as e:
        raise ImageSourceError(
            "image string is neither an existing file path nor valid base64"
        ) from e


@contextmanager
def image_on_disk(image: str | Path | bytes) -> Iterator[str]:
    """Yield a filesystem path for `image`, cleaning up any temp file on exit.

    Resolution: `bytes` -> raw image bytes; an existing-file path (`str`/`Path`)
    -> used as-is; any other `str` -> base64. Bytes (raw or base64-decoded) are
    written to a temp file whose extension is sniffed from the data and removed
    on exit — safe once `AddPicture(SaveWithDocument=True)` has embedded a copy.
    """
    data: bytes | None
    path: Path | None
    if isinstance(image, bytes):
        data, path = image, None
    elif isinstance(image, Path):
        data, path = None, image
    elif isinstance(image, str):
        p = Path(image)
        if p.is_file():
            data, path = None, p
        else:
            data, path = _decode_base64(image), None
    else:
        raise ImageSourceError(
            f"image must be a path, bytes, or base64 string; got {type(image).__name__}"
        )

    if data is None:
        assert path is not None
        if not path.is_file():
            raise ImageSourceError(f"image file not found: {str(path)!r}")
        try:
            with path.open("rb"):
                pass
        except OSError as e:
            raise ImageSourceError(f"image file is unreadable: {str(path)!r}") from e
        # AddPicture resolves a relative path against *Word's* working directory,
        # not the caller's, so a relative path silently fails with COM 0x80020009
        # ("not a valid file name"). Hand COM an absolute path.
        yield str(path.resolve())
        return

    ext = _sniff_extension(data)
    fd, tmp = tempfile.mkstemp(suffix=ext, prefix="wordlive_img_")
    try:
        with os.fdopen(fd, "wb") as f:
            f.write(data)
        yield tmp
    finally:
        try:
            os.unlink(tmp)
        except OSError:
            pass


# ---------------------------------------------------------------------------
# Reading images out — Flat OPC parsing
# ---------------------------------------------------------------------------


def image_parts_in_opc(xml_text: str) -> list[tuple[str, str]]:
    """Return `[(content_type, base64_data), …]` for every image part in a Flat OPC fragment.

    `Range.WordOpenXML` always serialises the full package skeleton (so even a
    one-character range comes back wrapped in a `<pkg:package>` with the document
    parts); the media parts are the `<pkg:part>` elements whose `pkg:contentType`
    is an `image/*` MIME and which carry a base64 `<pkg:binaryData>` body. We
    return those untouched (not yet decoded) so callers can either count them or
    decode the one they want.
    """
    root = ET.fromstring(xml_text)
    out: list[tuple[str, str]] = []
    for part in root.iter(_PKG_PART):
        ctype = part.get(_PKG_CONTENT_TYPE, "")
        if not ctype.startswith("image/"):
            continue
        binary = part.find(_PKG_BINARY_DATA)
        if binary is None or binary.text is None:
            continue
        out.append((ctype, binary.text))
    return out


def read_image_from_range(rng: Any) -> tuple[bytes, str]:
    """Extract the single embedded image in a COM `rng` as `(bytes, mime_type)`.

    Serialises the range with `Range.WordOpenXML` and pulls the lone `image/*`
    part out of the resulting Flat OPC package. Raises `ImageSourceError` when
    the range carries no image, or more than one — the caller should target a
    single picture (e.g. an `image:N` anchor) in the latter case.
    """
    try:
        xml_text = str(rng.WordOpenXML)
    except Exception as e:  # noqa: BLE001 — surfaced as a clean bad-input error
        raise ImageSourceError(f"could not read the range's package XML: {e}") from e
    parts = image_parts_in_opc(xml_text)
    if not parts:
        raise ImageSourceError("no embedded image found in this range")
    if len(parts) > 1:
        raise ImageSourceError(
            f"range contains {len(parts)} images; target a single image "
            "(an image:N anchor reads exactly one)"
        )
    ctype, b64 = parts[0]
    try:
        data = base64.b64decode(b64, validate=False)
    except (binascii.Error, ValueError) as e:
        raise ImageSourceError("embedded image data is not valid base64") from e
    return data, ctype
