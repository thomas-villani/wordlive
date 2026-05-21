"""Image-source helpers — turn a path / bytes / base64 string into a file on disk.

`insert_image` needs an actual file path for Word's `InlineShapes.AddPicture`,
which only reads from disk. This module normalises the three input shapes an
LLM might hold, sniffs the raster format from magic bytes (to pick a sensible
temp-file extension), and removes any temp file once Word has embedded a copy.
"""

from __future__ import annotations

import base64
import binascii
import os
import tempfile
from contextlib import contextmanager
from pathlib import Path
from typing import Iterator

from .exceptions import ImageSourceError

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
    raise ImageSourceError(
        "unrecognised image data: expected PNG, JPEG, GIF, BMP, or TIFF"
    )


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
        yield str(path)
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
