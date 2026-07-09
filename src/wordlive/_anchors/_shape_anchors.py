"""`shape:N` anchors, the shape collection, and text boxes (see also `_shapes`)."""

from __future__ import annotations

from collections.abc import Iterator
from typing import TYPE_CHECKING, Any

from .. import _com, _images, _shapes
from ..exceptions import AnchorNotFoundError, OpError

if TYPE_CHECKING:
    from pathlib import Path

    from .._document import Document

from ._base import Anchor
from ._helpers import (
    _safe_float,
)


class ShapeAnchor(Anchor):
    """A floating shape located by 1-based index — `shape:N`.

    Indexes the document's body-story floating shapes (text boxes, floating
    images, WordArt) in document order — the restyle handle that
    [`insert_text_box`][wordlive.Anchor.insert_text_box] and a floating
    [`insert_image`][wordlive.Anchor.insert_image] return. `shape_type` reports
    the kind (``"text_box"`` / ``"picture"`` / ``"wordart"`` / …). Restyle in
    place with `set_wrap` / `set_position` / `set_size` / `format` /
    `set_alt_text`; a text box's contents edit via `set_text`, a floating
    picture's image swaps via `replace_image`. Discover shapes via
    [`doc.shapes`][wordlive.Document.shapes] (or just the text boxes via
    [`doc.text_boxes`][wordlive.Document.text_boxes]).

    A floating shape anchors to a *paragraph*, not a character position, so
    positions renumber as shapes come and go — re-list, don't cache. Watermarks
    live in the header story and are excluded.
    """

    kind = "shape"

    def __init__(self, doc: Document, index: int) -> None:
        super().__init__(doc, name=f"shape:{index}")
        self._index = index

    @property
    def index(self) -> int:
        return self._index

    @property
    def anchor_id(self) -> str:
        return f"shape:{self._index}"

    def _shape(self) -> Any:
        shapes = _shapes.body_shapes(self._doc.com)
        if not (1 <= self._index <= len(shapes)):
            raise AnchorNotFoundError("shape", f"shape:{self._index}")
        return shapes[self._index - 1]

    def _range(self) -> Any:
        # A floating shape has no character position of its own; this is the
        # anchoring paragraph range, so inherited verbs (apply_style,
        # insert_caption) act on the paragraph the shape hangs off.
        return self._shape().Anchor

    @property
    def shape_type(self) -> str:
        """The shape kind — ``"text_box"`` / ``"picture"`` / ``"wordart"`` / ``"group"`` / …."""
        with _com.translate_com_errors():
            return _shapes.shape_kind(self._shape())

    @property
    def alt_text(self) -> str:
        """The shape's accessibility (alt) text, or ``""`` if unset."""
        with _com.translate_com_errors():
            return str(self._shape().AlternativeText or "")

    @property
    def text(self) -> str:
        """A text box's contents (``""`` for a shape with no text frame)."""
        with _com.translate_com_errors():
            shape = self._shape()
            if _shapes.shape_kind(shape) not in ("text_box", "auto_shape"):
                return ""
            try:
                return str(shape.TextFrame.TextRange.Text or "")
            except Exception:
                return ""

    def revision_segments(self) -> list[dict[str, Any]]:
        """The shape's text as a single unchanged segment (no tracked-change view).

        A floating shape's text lives in its own text-frame story, while
        `doc.revisions` enumerates the main body — the two don't share offsets, so
        tracked-change views aren't available inside shapes. This mirrors
        [`text`][wordlive.ShapeAnchor.text] (rather than reporting the *anchoring
        paragraph's* unrelated revision history). `text_final` / `text_original`
        therefore both equal `text`.
        """
        text = self.text
        return [{"text": text, "change": None}] if text else []

    @property
    def rotation(self) -> float:
        """The shape's clockwise rotation in degrees (``0.0`` if unrotated)."""
        with _com.translate_com_errors():
            return float(self._shape().Rotation)

    @property
    def z_order(self) -> int:
        """The shape's 1-based stacking position (`ZOrderPosition`; higher = nearer the front)."""
        with _com.translate_com_errors():
            return int(self._shape().ZOrderPosition)

    def set_wrap(
        self,
        wrap: str | None = None,
        *,
        side: str | None = None,
        distance_top: Any = None,
        distance_bottom: Any = None,
        distance_left: Any = None,
        distance_right: Any = None,
    ) -> ShapeAnchor:
        """Set how body text flows around the shape.

        `wrap` is the style — ``"square"`` / ``"tight"`` / ``"through"`` /
        ``"top-bottom"`` / ``"front"`` / ``"behind"``. `side` is which sides text
        flows past — ``"both"`` / ``"left"`` / ``"right"`` / ``"largest"`` (only
        ``"square"`` / ``"tight"`` / ``"through"`` honour it; Word ignores it for
        the others). `distance_*` are the standoff gaps between text and the shape
        (lengths in points / ``"0.1in"``). At least one argument is required.
        Returns `self` (chainable). Bad input raises `OpError`."""
        self._apply(
            _shapes.apply_shape_wrap,
            wrap,
            side=side,
            distance_top=distance_top,
            distance_bottom=distance_bottom,
            distance_left=distance_left,
            distance_right=distance_right,
        )
        return self

    def set_crop(
        self, *, left: Any = None, top: Any = None, right: Any = None, bottom: Any = None
    ) -> ShapeAnchor:
        """Crop a floating picture in from its edges. `left` / `top` / `right` /
        `bottom` are the amounts trimmed off each edge (lengths in points /
        ``"0.2in"``); cropping shrinks the displayed size. At least one edge is
        required. Only valid on a ``"picture"`` shape; raises `OpError` on a text
        box / WordArt / group. Returns `self` (chainable)."""
        self._apply(
            _shapes.apply_shape_crop,
            left=left,
            top=top,
            right=right,
            bottom=bottom,
            require_picture=True,
        )
        return self

    def set_position(
        self, *, left: Any = None, top: Any = None, relative_to: str | None = None
    ) -> ShapeAnchor:
        """Reposition the shape. `left` / `top` are lengths (points / ``"2in"``) or
        ``"center"``; `relative_to` is the frame they're measured from
        (``"margin"`` (default) or ``"page"``). Returns `self`. Bad input raises
        `OpError`."""
        self._apply(_shapes.apply_shape_position, left=left, top=top, relative_to=relative_to)
        return self

    def set_size(
        self, *, width: Any = None, height: Any = None, lock_aspect: bool | None = None
    ) -> ShapeAnchor:
        """Resize the shape. `width` / `height` are lengths (points / ``"3in"``);
        `lock_aspect` toggles proportional scaling (dropped automatically when both
        dimensions are given, so both stick). Returns `self`. Bad input raises
        `OpError`."""
        self._apply(_shapes.apply_shape_size, width=width, height=height, lock_aspect=lock_aspect)
        return self

    def set_rotation(self, degrees: Any) -> ShapeAnchor:
        """Rotate the shape clockwise by `degrees` (absolute angle, e.g. `30` or
        `-15`). Returns `self` (chainable). Bad input raises `OpError`."""
        self._apply(_shapes.apply_shape_rotation, degrees)
        return self

    def set_z_order(self, order: str) -> ShapeAnchor:
        """Restack the shape in the floating layer — ``"front"`` / ``"back"`` /
        ``"forward"`` / ``"backward"`` (this is the stacking order *among floats*,
        distinct from `set_wrap`'s in-front-of / behind-text).

        Note: `Document.Shapes` orders by z-order, so this **renumbers `shape:N`** —
        the returned `self` keeps its old index and may now address a different
        shape. Re-list (`doc.shapes`) before using a `shape:N` id again. Returns
        `self` for chaining the call itself; bad input raises `OpError`."""
        self._apply(_shapes.apply_shape_zorder, order)
        return self

    def set_text_frame(
        self,
        *,
        margin_left: Any = None,
        margin_right: Any = None,
        margin_top: Any = None,
        margin_bottom: Any = None,
        word_wrap: bool | None = None,
    ) -> ShapeAnchor:
        """Set a text box's internal margins and word-wrap. `margin_*` are lengths
        (points / ``"0.1in"``); `word_wrap` toggles whether text wraps to the box
        width. Only valid on a text box; raises `OpError` on a picture / WordArt /
        group. Returns `self` (chainable)."""
        self._apply(
            _shapes.apply_text_frame,
            margin_left=margin_left,
            margin_right=margin_right,
            margin_top=margin_top,
            margin_bottom=margin_bottom,
            word_wrap=word_wrap,
        )
        return self

    def format(
        self, *, fill: Any = None, border: str | bool | None = None, border_weight: Any = None
    ) -> ShapeAnchor:
        """Set the shape's fill and outline. `fill` is any colour; `border` is
        ``False`` (no outline), ``True`` (default), or a colour string;
        `border_weight` is the outline thickness (points / ``"1.5pt"``). Returns
        `self`. Bad input raises `OpError`."""
        self._apply(
            _shapes.apply_shape_format, fill=fill, border=border, border_weight=border_weight
        )
        return self

    def set_alt_text(self, text: str) -> ShapeAnchor:
        """Set the shape's accessibility (alt) text. Returns `self`."""
        with _com.translate_com_errors():
            self._shape().AlternativeText = text
        return self

    def replace_image(self, image: str | Path | bytes) -> ShapeAnchor:
        """Swap this floating picture's image in place.

        Delete + reinsert at the same anchor, preserving wrap / position / size /
        alt text — `image` is a path, raw bytes, or a base64 string (like
        `insert_image`). Only valid on a ``"picture"`` shape; raises `OpError`
        otherwise, `ImageSourceError` for a bad image. Returns `self` (chainable);
        wrap in `doc.edit(...)` for atomic undo."""
        with _images.image_on_disk(image) as disk_path:
            try:
                with _com.translate_com_errors():
                    _shapes.replace_shape_image(self._doc.com, self._shape(), disk_path)
            except (ValueError, TypeError) as e:
                raise OpError(str(e)) from e
        return self

    def set_text(self, text: str) -> None:
        """Replace a text box's contents. Raises `OpError` on a shape with no text
        frame (a picture / WordArt)."""
        with _com.translate_com_errors():
            shape = self._shape()
            if _shapes.shape_kind(shape) not in ("text_box", "auto_shape"):
                raise OpError("this shape has no text frame to set; set_text needs a text box")
            shape.TextFrame.TextRange.Text = text

    def ungroup(self) -> list[ShapeAnchor]:
        """Dissolve a group shape into its members, returning their `ShapeAnchor`s.

        The children become top-level floating shapes again (each keeps its own
        `shape:N` slot — re-list, don't cache). Only valid on a ``"group"`` shape;
        raises `OpError` otherwise. Wrap in `doc.edit(...)` for atomic undo. The
        inverse of [`Document.group_shapes`][wordlive.Document.group_shapes]."""
        with _com.translate_com_errors():
            shape = self._shape()
            if _shapes.shape_kind(shape) != "group":
                raise OpError("this shape is not a group; ungroup needs a group:N shape")
            names = _shapes.ungroup_shape(shape)
            anchors: list[ShapeAnchor] = []
            for name in names:
                if not name:
                    continue
                try:
                    idx = _shapes.index_of_named(self._doc.com, name)
                except OpError:
                    continue
                anchors.append(ShapeAnchor(self._doc, idx))
        return anchors

    def delete(self) -> None:
        """Delete the floating shape itself (not its anchoring paragraph)."""
        with _com.translate_com_errors():
            self._shape().Delete()

    def _apply(self, fn: Any, *args: Any, **kwargs: Any) -> None:
        """Run a `_shapes` mutator on this shape, translating COM and bad-input
        errors into the wordlive hierarchy (`OpError`)."""
        shape = self._shape()
        try:
            with _com.translate_com_errors():
                fn(shape, *args, **kwargs)
        except (ValueError, TypeError) as e:
            raise OpError(str(e)) from e


class ShapeCollection:
    """Iterable view over the document's floating shapes (`doc.shapes`).

    Index a shape by 1-based position (`doc.shapes[2]`) to get a
    [`ShapeAnchor`][wordlive.ShapeAnchor] (`shape:N`); `list()` summarises every
    shape — id, kind, size, wrap, and the `para:N` it's anchored in. Positions
    follow document order over the body story (header-story watermarks excluded),
    and renumber as shapes come and go — re-list, don't cache. The write mirror is
    [`insert_text_box`][wordlive.Anchor.insert_text_box] / a floating
    [`insert_image`][wordlive.Anchor.insert_image].
    """

    def __init__(self, doc: Document) -> None:
        self._doc = doc

    def __len__(self) -> int:
        with _com.translate_com_errors():
            return len(_shapes.body_shapes(self._doc.com))

    def __getitem__(self, index: int) -> ShapeAnchor:
        if isinstance(index, bool) or not isinstance(index, int):
            raise TypeError(f"shape index must be int, got {type(index).__name__}")
        n = len(self)
        if not (1 <= index <= n):
            raise AnchorNotFoundError("shape", str(index))
        return ShapeAnchor(self._doc, index)

    def __iter__(self) -> Iterator[ShapeAnchor]:
        for i in range(1, len(self) + 1):
            yield ShapeAnchor(self._doc, i)

    def list(self) -> list[dict[str, Any]]:
        """Every floating shape as `{index, anchor_id, shape_type, name, width,
        height, rotation, z_order, wrap, wrap_side, crop, alt_text, has_text,
        para}`.

        `shape_type` is the kind string; `width` / `height` are points; `rotation`
        the clockwise angle in degrees; `z_order` the 1-based stacking position;
        `wrap` the text-wrap keyword and `wrap_side` which sides text flows past;
        `crop` the picture's `{left, top, right, bottom}` insets in points (or
        `None`); `has_text` whether a text frame holds text; `para` the `para:N`
        the shape is anchored in.
        """
        out: list[dict[str, Any]] = []
        with _com.translate_com_errors():
            shapes = _shapes.body_shapes(self._doc.com)
            for i, shape in enumerate(shapes, start=1):
                kind = _shapes.shape_kind(shape)
                try:
                    wrap: str | None = _shapes.WRAP_TO_NAME.get(int(shape.WrapFormat.Type))
                except Exception:
                    wrap = None
                try:
                    wrap_side: str | None = _shapes.WRAP_SIDE_TO_NAME.get(
                        int(shape.WrapFormat.Side)
                    )
                except Exception:
                    wrap_side = None
                crop = _shapes.crop_values(shape) if kind == "picture" else None
                try:
                    z_order: int | None = int(shape.ZOrderPosition)
                except Exception:
                    z_order = None
                try:
                    has_text = kind in ("text_box", "auto_shape") and bool(shape.TextFrame.HasText)
                except Exception:
                    has_text = False
                try:
                    alt_text = str(shape.AlternativeText or "")
                except Exception:
                    alt_text = ""
                try:
                    start: int | None = int(shape.Anchor.Start)
                except Exception:
                    start = None
                para_id: str | None = None
                if start is not None:
                    para = self._doc.paragraphs.at(start)
                    para_id = para.anchor_id if para is not None else None
                out.append(
                    {
                        "index": i,
                        "anchor_id": f"shape:{i}",
                        "shape_type": kind,
                        "name": str(getattr(shape, "Name", "") or ""),
                        "width": _safe_float(shape, "Width"),
                        "height": _safe_float(shape, "Height"),
                        "rotation": _safe_float(shape, "Rotation"),
                        "z_order": z_order,
                        "wrap": wrap,
                        "wrap_side": wrap_side,
                        "crop": crop,
                        "alt_text": alt_text,
                        "has_text": has_text,
                        "para": para_id,
                    }
                )
        return out


class TextBoxCollection:
    """Iterable view over the document's text boxes (`doc.text_boxes`).

    The ``shape_type == "text_box"`` subset of [`doc.shapes`]
    [wordlive.Document.shapes] — a discovery filter, *not* a second id space: each
    text box keeps its canonical `shape:N` id (its position among *all* floating
    shapes), so `doc.text_boxes[1].anchor_id` may be e.g. `shape:3`. Index 1-based
    over the text boxes; `list()` is the text-box rows of `doc.shapes.list()`.
    """

    def __init__(self, doc: Document) -> None:
        self._doc = doc

    def _indices(self) -> list[int]:
        """1-based `shape:N` indices (over all body shapes) that are text boxes."""
        shapes = _shapes.body_shapes(self._doc.com)
        return [i for i, s in enumerate(shapes, start=1) if _shapes.shape_kind(s) == "text_box"]

    def __len__(self) -> int:
        with _com.translate_com_errors():
            return len(self._indices())

    def __getitem__(self, index: int) -> ShapeAnchor:
        if isinstance(index, bool) or not isinstance(index, int):
            raise TypeError(f"text box index must be int, got {type(index).__name__}")
        with _com.translate_com_errors():
            idxs = self._indices()
        if not (1 <= index <= len(idxs)):
            raise AnchorNotFoundError("text box", str(index))
        return ShapeAnchor(self._doc, idxs[index - 1])

    def __iter__(self) -> Iterator[ShapeAnchor]:
        with _com.translate_com_errors():
            idxs = self._indices()
        for unfiltered in idxs:
            yield ShapeAnchor(self._doc, unfiltered)

    def list(self) -> list[dict[str, Any]]:
        """The text-box rows of `doc.shapes.list()` (each keeping its `shape:N` id)."""
        return [row for row in ShapeCollection(self._doc).list() if row["shape_type"] == "text_box"]
