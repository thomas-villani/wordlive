"""Floating shapes — the body-story `Document.Shapes` addressed as `shape:N`.

`Document.Shapes` is a *separate* collection from `InlineShapes`: a floating
shape (a text box, a floating/converted image, WordArt) anchors to a paragraph,
not a character position. `shape:N` is therefore positional over document order
and renumbers when shapes are added/removed before it — **re-list, don't cache**
(the same rule as `image:N` / `chart:N`).

Live-probed facts this module encodes (2026-06-19):

- **`Shape.Type` discriminates** the kind — picture = 13, text box = 17, WordArt
  (text-effect) = 15; `shape_kind` maps these onto a public `shape_type` string.
- **Watermarks are WordArt in the *header* story** (`set_watermark`), so they
  never appear in `Document.Shapes` to begin with; `body_shapes` keeps a
  name-prefix + `Anchor.StoryType` guard anyway (belt-and-suspenders).
- **Replacing a floating picture's image is delete + reinsert at the same
  anchor**, preserving wrap / position / size. `Shape.Fill.UserPicture` was
  rejected: on a picture shape it *overlays* a second picture-fill rather than
  replacing the displayed picture (the range then holds two images).
- **Body-shape position constants and `Anchor.StoryType` read cleanly** under
  late binding (the same `Shape.Left = wdShapeCenter` / `RelativeHorizontal
  Position` calls `set_watermark` makes on a header shape).
"""

from __future__ import annotations

from typing import Any

from ._format import to_bgr, to_points
from .constants import (
    MsoShapeType,
    MsoTriState,
    MsoZOrderCmd,
    WdRelativeHorizontalPosition,
    WdRelativeVerticalPosition,
    WdShapePosition,
    WdStoryType,
    WdWrapType,
)
from .exceptions import OpError

# Watermark shapes are named like Word's own feature (mirror Document's prefix).
_WATERMARK_NAME_PREFIX = "PowerPlusWaterMarkObject"

# Shape.Type (MsoShapeType int) -> public `shape_type` string.
TYPE_TO_KIND: dict[int, str] = {
    int(MsoShapeType.TEXT_BOX): "text_box",
    int(MsoShapeType.PICTURE): "picture",
    int(MsoShapeType.TEXT_EFFECT): "wordart",
    int(MsoShapeType.GROUP): "group",
    int(MsoShapeType.AUTO_SHAPE): "auto_shape",
}

# Floating wrap keyword -> WdWrapType (same vocabulary as `insert_image`, minus
# "inline" / "auto" which only make sense at insert time). `_anchors._WRAP_NAMES`
# aliases this so the two surfaces can't drift.
WRAP_NAMES: dict[str, WdWrapType] = {
    "square": WdWrapType.SQUARE,
    "tight": WdWrapType.TIGHT,
    "through": WdWrapType.THROUGH,
    "top-bottom": WdWrapType.TOP_BOTTOM,
    "front": WdWrapType.FRONT,
    "behind": WdWrapType.BEHIND,
}
WRAP_TO_NAME: dict[int, str] = {int(v): k for k, v in WRAP_NAMES.items()}

# Z-order keyword -> MsoZOrderCmd (the four reorder verbs `set_z_order` exposes).
ZORDER_NAMES: dict[str, int] = {
    "front": int(MsoZOrderCmd.BRING_TO_FRONT),
    "back": int(MsoZOrderCmd.SEND_TO_BACK),
    "forward": int(MsoZOrderCmd.BRING_FORWARD),
    "backward": int(MsoZOrderCmd.SEND_BACKWARD),
}

# Shape kinds that carry a text frame (so `set_text` / `set_text_frame` apply).
_TEXT_FRAME_KINDS = ("text_box", "auto_shape")

# Position-frame keyword -> (horizontal enum, vertical enum) for `set_position`.
RELATIVE_TO: dict[str, tuple[int, int]] = {
    "margin": (
        int(WdRelativeHorizontalPosition.MARGIN),
        int(WdRelativeVerticalPosition.MARGIN),
    ),
    "page": (
        int(WdRelativeHorizontalPosition.PAGE),
        int(WdRelativeVerticalPosition.PAGE),
    ),
}


def _is_watermark(shape: Any) -> bool:
    try:
        return str(shape.Name or "").startswith(_WATERMARK_NAME_PREFIX)
    except Exception:
        return False


def _in_body_story(shape: Any) -> bool:
    """True if the shape is anchored in the main text story (not a header/footer)."""
    try:
        return int(shape.Anchor.StoryType) == int(WdStoryType.MAIN_TEXT)
    except Exception:
        # No readable anchor → treat as body. `Document.Shapes` already excludes
        # header-story shapes, so this only loosens an unreadable edge case.
        return True


def body_shapes(doc_com: Any) -> list[Any]:
    """The document's body-story floating shapes, in document order.

    Re-lists on every call (positions renumber as shapes come and go). Excludes
    header-story watermark shapes by both name prefix and anchor story type.
    """
    out: list[Any] = []
    shapes = doc_com.Shapes
    for i in range(1, int(shapes.Count) + 1):
        shape = shapes.Item(i)
        if _is_watermark(shape) or not _in_body_story(shape):
            continue
        out.append(shape)
    return out


def shape_kind(shape: Any) -> str:
    """Map `Shape.Type` onto the public `shape_type` string (``"other"`` if unmapped)."""
    try:
        return TYPE_TO_KIND.get(int(shape.Type), "other")
    except Exception:
        return "other"


def index_of_named(doc_com: Any, name: str) -> int:
    """1-based index, among `body_shapes`, of the shape named `name`.

    A just-inserted text box / floating image is addressed by setting a unique
    name on it and locating that name here — don't assume "last", since other
    floats can reorder. Raises `OpError` if absent (shouldn't happen right after
    an insert).
    """
    for i, shape in enumerate(body_shapes(doc_com), start=1):
        try:
            if str(shape.Name or "") == name:
                return i
        except Exception:
            continue
    raise OpError(f"could not locate inserted shape named {name!r}")


def _position_value(value: Any) -> float:
    """A length (points or unit string) or ``"center"`` → a `Shape.Left/Top` value."""
    if isinstance(value, str) and value.strip().lower() in ("center", "centre"):
        return float(WdShapePosition.CENTER)
    return to_points(value)


# --- post-insert mutators (operate on the floating Shape COM object) -----------
#
# Each helper mutates one facet of a floating shape's layout / appearance. Bad
# input raises `OpError`; COM failures surface via the caller's
# `translate_com_errors`. None of these select or activate anything, so they're
# inherently polite (the user's cursor/scroll never moves).


def apply_shape_wrap(shape: Any, wrap: str) -> None:
    """Set how body text flows around the shape (the floating wrap keywords)."""
    if wrap not in WRAP_NAMES:
        raise OpError(f"unknown wrap {wrap!r}; expected one of {sorted(WRAP_NAMES)}")
    shape.WrapFormat.Type = int(WRAP_NAMES[wrap])


def apply_shape_position(
    shape: Any,
    *,
    left: Any = None,
    top: Any = None,
    relative_to: str | None = None,
) -> None:
    """Reposition the shape. `left` / `top` are lengths (points / ``"2in"``) or
    ``"center"``; `relative_to` is the frame they're measured from
    (``"margin"`` (default) or ``"page"``)."""
    if relative_to is not None:
        if relative_to not in RELATIVE_TO:
            raise OpError(
                f"unknown relative_to {relative_to!r}; expected one of {sorted(RELATIVE_TO)}"
            )
        h, v = RELATIVE_TO[relative_to]
        shape.RelativeHorizontalPosition = h
        shape.RelativeVerticalPosition = v
    if left is not None:
        shape.Left = _position_value(left)
    if top is not None:
        shape.Top = _position_value(top)


def apply_shape_size(
    shape: Any,
    *,
    width: Any = None,
    height: Any = None,
    lock_aspect: bool | None = None,
) -> None:
    """Resize the shape. `width` / `height` are lengths (points / ``"3in"``);
    `lock_aspect` toggles proportional scaling. When both dimensions are given,
    aspect-lock is dropped for the set so both are honoured exactly (then restored
    if `lock_aspect=True` was asked)."""
    both = width is not None and height is not None
    if both and lock_aspect is not False:
        shape.LockAspectRatio = int(MsoTriState.FALSE)
    if width is not None:
        shape.Width = to_points(width)
    if height is not None:
        shape.Height = to_points(height)
    if lock_aspect is not None:
        shape.LockAspectRatio = int(MsoTriState.TRUE if lock_aspect else MsoTriState.FALSE)


def apply_shape_format(
    shape: Any,
    *,
    fill: Any = None,
    border: str | bool | None = None,
    border_weight: Any = None,
) -> None:
    """Set the shape's fill colour and outline. `fill` is any colour;
    `border` is ``False`` (no outline), ``True`` (default outline), or a colour
    string; `border_weight` is the outline thickness (points / ``"1.5pt"``)."""
    if fill is not None:
        shape.Fill.Visible = int(MsoTriState.TRUE)
        shape.Fill.Solid()
        shape.Fill.ForeColor.RGB = to_bgr(fill)
    if border is False:
        shape.Line.Visible = int(MsoTriState.FALSE)
    elif border is True:
        shape.Line.Visible = int(MsoTriState.TRUE)
    elif isinstance(border, str):
        shape.Line.Visible = int(MsoTriState.TRUE)
        shape.Line.ForeColor.RGB = to_bgr(border)
    if border_weight is not None:
        shape.Line.Weight = to_points(border_weight)


def apply_shape_rotation(shape: Any, degrees: Any) -> None:
    """Rotate the shape clockwise by `degrees` (absolute, not relative)."""
    try:
        deg = float(degrees)
    except (ValueError, TypeError) as e:
        raise OpError(f"rotation must be a number of degrees; got {degrees!r}") from e
    shape.Rotation = deg


def apply_shape_zorder(shape: Any, order: str) -> None:
    """Restack the shape in the floating layer — ``"front"`` / ``"back"`` /
    ``"forward"`` / ``"backward"``."""
    if order not in ZORDER_NAMES:
        raise OpError(f"unknown z-order {order!r}; expected one of {sorted(ZORDER_NAMES)}")
    shape.ZOrder(ZORDER_NAMES[order])


def apply_text_frame(
    shape: Any,
    *,
    margin_left: Any = None,
    margin_right: Any = None,
    margin_top: Any = None,
    margin_bottom: Any = None,
    word_wrap: bool | None = None,
) -> None:
    """Set a text box's internal margins and word-wrap.

    `margin_*` are lengths (points / ``"0.1in"``); `word_wrap` toggles whether
    text wraps to the box width (off lets a line run past the box edge). Raises
    `OpError` on a shape with no text frame (a picture / WordArt / group).
    """
    if shape_kind(shape) not in _TEXT_FRAME_KINDS:
        raise OpError("set_text_frame needs a text box (a shape with a text frame)")
    frame = shape.TextFrame
    if margin_left is not None:
        frame.MarginLeft = to_points(margin_left)
    if margin_right is not None:
        frame.MarginRight = to_points(margin_right)
    if margin_top is not None:
        frame.MarginTop = to_points(margin_top)
    if margin_bottom is not None:
        frame.MarginBottom = to_points(margin_bottom)
    if word_wrap is not None:
        frame.WordWrap = int(MsoTriState.TRUE if word_wrap else MsoTriState.FALSE)


def group_shapes(doc_com: Any, shapes: list[Any]) -> Any:
    """Group `shapes` (two or more floating shapes) into one group shape.

    Word disables grouping for shapes that can't overlap, so each member's
    `WrapFormat.AllowOverlap` is enabled first (the live-probed prerequisite).
    Returns the new group `Shape`. Raises `OpError` for fewer than two shapes or
    an unnamed member (the group is built via `Shapes.Range([names])`).
    """
    if len(shapes) < 2:
        raise OpError("grouping needs at least two shapes")
    names: list[str] = []
    for s in shapes:
        try:
            s.WrapFormat.AllowOverlap = True
        except Exception:
            pass  # best-effort; Group() will report if the shape still can't group
        name = str(s.Name or "")
        if not name:
            raise OpError("cannot group a shape with no name")
        names.append(name)
    return doc_com.Shapes.Range(names).Group()


def ungroup_shape(shape: Any) -> list[str]:
    """Dissolve a group shape into its members; return the members' names.

    The children become top-level floating shapes again (keeping their names), so
    the caller re-locates each `shape:N` by name. Raises `OpError` if `shape`
    isn't a group.
    """
    if shape_kind(shape) != "group":
        raise OpError("ungroup needs a group shape")
    items = shape.GroupItems
    names = [str(items.Item(i).Name or "") for i in range(1, int(items.Count) + 1)]
    shape.Ungroup()
    return names


def replace_shape_image(doc_com: Any, shape: Any, disk_path: str) -> Any:
    """Swap a floating picture's image by delete + reinsert at the same anchor.

    Preserves wrap / position / size / lock-aspect / alt-text / name. Returns the
    new `Shape`. `Shape.Fill.UserPicture` was rejected by a live probe: on a
    picture shape it overlays a second picture-fill instead of replacing the
    displayed picture. Raises `OpError` if the shape isn't a floating picture.
    """
    if shape_kind(shape) != "picture":
        raise OpError("replace_image needs a picture shape (a floating image)")
    pos = int(shape.Anchor.Start)
    wrap = int(shape.WrapFormat.Type)
    left, top = float(shape.Left), float(shape.Top)
    width, height = float(shape.Width), float(shape.Height)
    rel_h = int(shape.RelativeHorizontalPosition)
    rel_v = int(shape.RelativeVerticalPosition)
    lock = int(shape.LockAspectRatio)
    try:
        alt = str(shape.AlternativeText or "")
    except Exception:
        alt = ""
    name = str(shape.Name or "")
    shape.Delete()
    rng = doc_com.Range(pos, pos)
    ish = rng.InlineShapes.AddPicture(
        FileName=disk_path, LinkToFile=False, SaveWithDocument=True, Range=rng
    )
    new_shape = ish.ConvertToShape()
    new_shape.WrapFormat.Type = wrap
    new_shape.LockAspectRatio = int(MsoTriState.FALSE)  # honour both dimensions exactly
    new_shape.Width, new_shape.Height = width, height
    new_shape.RelativeHorizontalPosition = rel_h
    new_shape.RelativeVerticalPosition = rel_v
    new_shape.Left, new_shape.Top = left, top
    new_shape.LockAspectRatio = lock
    if alt:
        new_shape.AlternativeText = alt
    if name:
        new_shape.Name = name
    return new_shape
