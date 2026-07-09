"""`image:N` anchors and the image collection (see also the `_images` feature module)."""

from __future__ import annotations

from collections.abc import Iterator
from typing import TYPE_CHECKING, Any

from .. import _com, _images, _shapes
from ..exceptions import AnchorNotFoundError, OpError

if TYPE_CHECKING:
    from .._document import Document

from ._base import Anchor
from ._helpers import (
    _safe_float,
    _safe_str,
)

# ---------------------------------------------------------------------------
# Images (read side)
# ---------------------------------------------------------------------------


class ImageAnchor(Anchor):
    """An embedded picture located by 1-based index — `image:N`.

    Mirrors Word's own `InlineShapes(N)` ordering (document order). The anchor
    resolves to the picture's own one-character range, so
    [`read_image`][wordlive.Anchor.read_image] pulls exactly that image's bytes
    out. Discover the available images — with their MIME, size, and the `para:N`
    they sit in — via [`doc.images`][wordlive.Document.images]. An image carries
    no editable text, so `set_text` raises; `read_image()` is the point of it.
    """

    kind = "image"

    def __init__(self, doc: Document, index: int) -> None:
        super().__init__(doc, name=f"image:{index}")
        self._index = index

    @property
    def index(self) -> int:
        return self._index

    @property
    def anchor_id(self) -> str:
        return f"image:{self._index}"

    def _shape(self) -> Any:
        shapes = self._doc.com.InlineShapes
        n = int(shapes.Count)
        if not (1 <= self._index <= n):
            raise AnchorNotFoundError("image", f"image:{self._index}")
        return shapes.Item(self._index)

    def _range(self) -> Any:
        return self._shape().Range

    @property
    def alt_text(self) -> str:
        """The picture's accessibility (alt) text, or ``""`` if unset."""
        with _com.translate_com_errors():
            return str(self._shape().AlternativeText or "")

    def set_alt_text(self, text: str) -> ImageAnchor:
        """Set the picture's accessibility (alt) text. Returns `self` (chainable);
        wrap in `doc.edit(...)` for atomic undo."""
        with _com.translate_com_errors():
            self._shape().AlternativeText = text
        return self

    def set_size(
        self, *, width: Any = None, height: Any = None, lock_aspect: bool | None = None
    ) -> ImageAnchor:
        """Resize the inline picture. `width` / `height` are lengths (points /
        ``"3in"``); `lock_aspect` toggles proportional scaling (dropped
        automatically when both dimensions are given, so both stick). To *re-wrap*
        a picture (float it) use `insert_image(wrap=…)` — that crosses it into
        `shape:N`. Returns `self` (chainable). Bad input raises `OpError`."""
        self._apply(_shapes.apply_shape_size, width=width, height=height, lock_aspect=lock_aspect)
        return self

    def set_crop(
        self, *, left: Any = None, top: Any = None, right: Any = None, bottom: Any = None
    ) -> ImageAnchor:
        """Crop the inline picture in from its edges. `left` / `top` / `right` /
        `bottom` are the amounts trimmed off each edge (lengths in points /
        ``"0.2in"``); cropping shrinks the displayed size. At least one edge is
        required. Returns `self` (chainable). Bad input raises `OpError`."""
        self._apply(_shapes.apply_shape_crop, left=left, top=top, right=right, bottom=bottom)
        return self

    def _apply(self, fn: Any, *args: Any, **kwargs: Any) -> None:
        """Run a `_shapes` size/crop helper on this inline picture (`InlineShape`
        shares the `Width` / `Height` / `LockAspectRatio` / `PictureFormat`
        surface), translating COM and bad-input errors into the wordlive hierarchy
        (`OpError`)."""
        shape = self._shape()
        try:
            with _com.translate_com_errors():
                fn(shape, *args, **kwargs)
        except (ValueError, TypeError) as e:
            raise OpError(str(e)) from e

    def set_text(self, text: str) -> None:
        raise OpError("an image anchor has no text to set; use read_image() to extract its bytes")


class ImageCollection:
    """Read-only, iterable view over the document's embedded images (`doc.images`).

    Index an image by 1-based position (`doc.images[2]`) to get an
    [`ImageAnchor`][wordlive.ImageAnchor] (`image:N`), then `read_image()` for
    its bytes + MIME. `list()` summarises every image — id, MIME, size, alt text,
    and the `para:N` it's anchored in — so a model can see what's there before
    pulling any bytes. Positions match Word's own `InlineShapes(n)` ordering.
    """

    def __init__(self, doc: Document) -> None:
        self._doc = doc

    def __len__(self) -> int:
        with _com.translate_com_errors():
            return int(self._doc.com.InlineShapes.Count)

    def __getitem__(self, index: int) -> ImageAnchor:
        if isinstance(index, bool) or not isinstance(index, int):
            raise TypeError(f"image index must be int, got {type(index).__name__}")
        n = len(self)
        if not (1 <= index <= n):
            raise AnchorNotFoundError("image", str(index))
        return ImageAnchor(self._doc, index)

    def __iter__(self) -> Iterator[ImageAnchor]:
        with _com.translate_com_errors():
            count = int(self._doc.com.InlineShapes.Count)
        for i in range(1, count + 1):
            yield ImageAnchor(self._doc, i)

    def list(self) -> list[dict[str, Any]]:
        """Every image as `{index, anchor_id, mime, width, height, crop, alt_text, para}`.

        `mime` is the picture's content type (read from its package XML —
        ``None`` if the shape isn't a raster image, e.g. an embedded chart or OLE
        object). `width`/`height` are in points; `crop` the `{left, top, right,
        bottom}` insets in points (or ``None`` if uncropped). `para` is the
        `para:N` anchor of the paragraph the image sits in (or ``None`` if it
        can't be located). Reads each image's content type but not its
        (potentially large) bytes — call [`read_image`][wordlive.Anchor.read_image]
        for those.
        """
        out: list[dict[str, Any]] = []
        with _com.translate_com_errors():
            shapes = self._doc.com.InlineShapes
            count = int(shapes.Count)
            for i in range(1, count + 1):
                shape = shapes.Item(i)
                rng = shape.Range
                try:
                    start = int(rng.Start)
                except Exception:
                    start = None
                mime: str | None = None
                try:
                    parts = _images.image_parts_in_opc(str(rng.WordOpenXML))
                    if len(parts) == 1:
                        mime = parts[0][0]
                except Exception:
                    mime = None
                para_id: str | None = None
                if start is not None:
                    para = self._doc.paragraphs.at(start)
                    para_id = para.anchor_id if para is not None else None
                out.append(
                    {
                        "index": i,
                        "anchor_id": f"image:{i}",
                        "mime": mime,
                        "width": _safe_float(shape, "Width"),
                        "height": _safe_float(shape, "Height"),
                        "crop": _shapes.crop_values(shape),
                        "alt_text": _safe_str(shape, "AlternativeText"),
                        "para": para_id,
                    }
                )
        return out
