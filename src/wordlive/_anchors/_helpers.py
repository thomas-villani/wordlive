"""Free helpers shared across the anchor classes: text/range readers, font and
paragraph-format coercion + read-back, border/tab tables, and small validators."""

from __future__ import annotations

from typing import TYPE_CHECKING, Any

from .. import _charts, _shapes
from .._format import bgr_to_hex, to_bgr, to_points
from ..constants import (
    WD_UNDEFINED,
    WdBorderType,
    WdBreakType,
    WdColorIndex,
    WdContentControlType,
    WdDropPosition,
    WdFieldType,
    WdInformation,
    WdLineSpacing,
    WdLineStyle,
    WdParagraphAlignment,
    WdTabAlignment,
    WdTabLeader,
    WdUnderline,
    WdWrapType,
)
from ..exceptions import OpError

if TYPE_CHECKING:
    pass


_ALIGNMENT_NAMES = {
    "left": WdParagraphAlignment.LEFT,
    "center": WdParagraphAlignment.CENTER,
    "centre": WdParagraphAlignment.CENTER,
    "right": WdParagraphAlignment.RIGHT,
    "justify": WdParagraphAlignment.JUSTIFY,
}


def _utf16_len(s: str) -> int:
    """Length of `s` in UTF-16 code units — Word's native character count.

    Python's `len()` counts code points, so astral-plane characters (emoji,
    historic scripts) count as 1. Word counts UTF-16 code units, so the same
    character counts as 2. Use this whenever the result is fed back into a
    Word `Range(start, end)` after a `Range.Text = ...` assignment.
    """
    return len(s.encode("utf-16-le")) // 2


def range_text(rng: Any, *, may_have_shapes: bool = True) -> str:
    """Read a COM range's text with inline shapes surfaced as ``[image]`` tokens.

    Word represents each inline shape (embedded picture / OLE object) as a single
    placeholder character in the text stream. That character is *not* a reserved
    control code — it varies by build and is indistinguishable by value from real
    text (a forward slash, on some Word versions) — so a naive string replace
    would clobber genuine characters. Instead we locate the shapes via the
    ``InlineShapes`` collection and swap only the character at each shape's own
    position, leaving real text untouched. A range with no inline shapes returns
    its raw text unchanged.

    Pass ``may_have_shapes=False`` when the caller already knows the document holds
    no inline shapes. Reaching for ``rng.InlineShapes`` mints a COM object pywin32
    must wrap, and a bulk paragraph walk pays that per paragraph just to learn the
    count is zero.
    """
    raw = str(rng.Text or "")
    if not may_have_shapes:
        return raw
    try:
        shapes = rng.InlineShapes
        count = int(shapes.Count)
        if count <= 0:
            return raw
        base = int(rng.Start)
        offsets = sorted({int(shapes.Item(i).Range.Start) - base for i in range(1, count + 1)})
    except Exception:
        # If the shape geometry can't be read, fall back to the raw text rather
        # than risk mangling it — a phantom char is better than a crash.
        return raw
    chars = list(raw)
    for off in reversed(offsets):
        if 0 <= off < len(chars):
            chars[off] = "[image]"
    return "".join(chars)


def _coerce_alignment(value: Any) -> int:
    if isinstance(value, WdParagraphAlignment):
        return int(value)
    if isinstance(value, int) and not isinstance(value, bool):
        return value
    if isinstance(value, str):
        try:
            return int(_ALIGNMENT_NAMES[value.lower()])
        except KeyError:
            raise ValueError(
                f"unknown alignment {value!r}; expected one of {sorted(set(_ALIGNMENT_NAMES))}"
            ) from None
    raise TypeError(
        f"alignment must be WdParagraphAlignment, int, or str; got {type(value).__name__}"
    )


# Highlight keywords -> WdColorIndex. Highlight is a palette index, not an RGB,
# so it bypasses the colour helper. `auto`/`none` clears the highlight.
_HIGHLIGHT_NAMES: dict[str, WdColorIndex] = {
    "none": WdColorIndex.AUTO,
    "auto": WdColorIndex.AUTO,
    "black": WdColorIndex.BLACK,
    "blue": WdColorIndex.BLUE,
    "turquoise": WdColorIndex.TURQUOISE,
    "bright-green": WdColorIndex.BRIGHT_GREEN,
    "green": WdColorIndex.GREEN,
    "pink": WdColorIndex.PINK,
    "red": WdColorIndex.RED,
    "yellow": WdColorIndex.YELLOW,
    "white": WdColorIndex.WHITE,
    "dark-blue": WdColorIndex.DARK_BLUE,
    "teal": WdColorIndex.TEAL,
    "violet": WdColorIndex.VIOLET,
    "dark-red": WdColorIndex.DARK_RED,
    "dark-yellow": WdColorIndex.DARK_YELLOW,
    "gray-50": WdColorIndex.GRAY_50,
    "gray-25": WdColorIndex.GRAY_25,
}


# Reverse of `_HIGHLIGHT_NAMES` for reading a highlight index back to a keyword.
# Skip the `auto` alias so index 0 renders as the canonical `"none"`.
_HIGHLIGHT_BY_INDEX: dict[int, str] = {
    int(v): k for k, v in _HIGHLIGHT_NAMES.items() if k != "auto"
}


def _read_highlight(raw: Any) -> str | None:
    """A `Range.HighlightColorIndex` read as a keyword (`"yellow"`, `"none"`, …),
    or `None` when Word reports `WD_UNDEFINED` (highlight varies across the runs)."""
    idx = int(raw)
    if idx == WD_UNDEFINED:
        return None
    return _HIGHLIGHT_BY_INDEX.get(idx, f"index:{idx}")


# Content-control kind keyword -> WdContentControlType. Canonical keys plus a
# few forgiving aliases (the names an agent is likely to reach for).
_CC_TYPE_NAMES: dict[str, WdContentControlType] = {
    "rich_text": WdContentControlType.RICH_TEXT,
    "richtext": WdContentControlType.RICH_TEXT,
    "text": WdContentControlType.TEXT,
    "plain_text": WdContentControlType.TEXT,
    "picture": WdContentControlType.PICTURE,
    "combo_box": WdContentControlType.COMBO_BOX,
    "combobox": WdContentControlType.COMBO_BOX,
    "dropdown": WdContentControlType.DROPDOWN_LIST,
    "dropdown_list": WdContentControlType.DROPDOWN_LIST,
    "date": WdContentControlType.DATE,
    "building_block": WdContentControlType.BUILDING_BLOCK_GALLERY,
    "group": WdContentControlType.GROUP,
    "checkbox": WdContentControlType.CHECKBOX,
    "check_box": WdContentControlType.CHECKBOX,
    "repeating_section": WdContentControlType.REPEATING_SECTION,
}


_CC_LIST_TYPES = (WdContentControlType.COMBO_BOX, WdContentControlType.DROPDOWN_LIST)


def _normalize_cc_items(items: list[Any]) -> list[tuple[str, str]]:
    """Normalize dropdown / combo-box `items` to ``(text, value)`` pairs.

    Each item is a plain string (text == value) or a ``{"text", "value"}`` dict
    (value defaults to text). Shared by `insert_content_control` and
    `ContentControl.set_items` so the two never drift.
    """
    pairs: list[tuple[str, str]] = []
    for item in items:
        if isinstance(item, dict):
            entry_text = str(item.get("text", ""))
            raw_value = item.get("value")
            value = str(raw_value) if raw_value is not None else entry_text
        else:
            entry_text = value = str(item)
        pairs.append((entry_text, value))
    return pairs


def _coerce_highlight(value: Any) -> int:
    if isinstance(value, WdColorIndex):
        return int(value)
    if isinstance(value, int) and not isinstance(value, bool):
        return value
    if isinstance(value, str):
        try:
            return int(_HIGHLIGHT_NAMES[value.lower()])
        except KeyError:
            raise ValueError(
                f"unknown highlight {value!r}; expected one of {sorted(_HIGHLIGHT_NAMES)}"
            ) from None
    raise TypeError(f"highlight must be WdColorIndex, int, or str; got {type(value).__name__}")


def _apply_font(
    font: Any,
    *,
    bold: bool | None = None,
    italic: bool | None = None,
    underline: bool | None = None,
    strikethrough: bool | None = None,
    font_name: str | None = None,
    size: Any = None,
    color: Any = None,
    subscript: bool | None = None,
    superscript: bool | None = None,
    small_caps: bool | None = None,
    all_caps: bool | None = None,
    spacing: Any = None,
) -> None:
    """Write character-formatting properties onto a COM `Font`.

    Tri-state: only kwargs that aren't `None` are written. Shared by
    `Anchor.format_run` (on a range's `Font`) and `Style.format_run` (on a
    style's `Font`) — the two diverge only in `highlight`, which is a `Range`
    property and is handled by the caller, not here. Raises `ValueError`/
    `TypeError` from the colour/length coercers on bad input.
    """
    if bold is not None:
        font.Bold = bool(bold)
    if italic is not None:
        font.Italic = bool(italic)
    if underline is not None:
        font.Underline = int(WdUnderline.SINGLE if underline else WdUnderline.NONE)
    if strikethrough is not None:
        font.StrikeThrough = bool(strikethrough)
    if font_name is not None:
        font.Name = str(font_name)
    if size is not None:
        font.Size = to_points(size)
    if color is not None:
        font.Color = to_bgr(color)
    if subscript is not None:
        font.Subscript = bool(subscript)
    if superscript is not None:
        font.Superscript = bool(superscript)
    if small_caps is not None:
        font.SmallCaps = bool(small_caps)
    if all_caps is not None:
        font.AllCaps = bool(all_caps)
    if spacing is not None:
        font.Spacing = to_points(spacing)


def _apply_paragraph_format(
    pf: Any,
    *,
    alignment: Any = None,
    left_indent: Any = None,
    right_indent: Any = None,
    first_line_indent: Any = None,
    space_before: Any = None,
    space_after: Any = None,
    line_spacing: Any = None,
    page_break_before: bool | None = None,
    keep_together: bool | None = None,
    keep_with_next: bool | None = None,
    widow_control: bool | None = None,
) -> None:
    """Write paragraph-formatting properties onto a COM `ParagraphFormat`.

    Tri-state: only kwargs that aren't `None` are written. Shared by
    `Anchor.format_paragraph` (on a range's `ParagraphFormat`) and
    `Style.format_paragraph` (on a style's `ParagraphFormat`). Indents/spacing
    accept a number (points) or a unit string via the length helper. Raises
    `ValueError`/`TypeError` on bad input.
    """
    if alignment is not None:
        pf.Alignment = _coerce_alignment(alignment)
    if left_indent is not None:
        pf.LeftIndent = to_points(left_indent)
    if right_indent is not None:
        pf.RightIndent = to_points(right_indent)
    if first_line_indent is not None:
        pf.FirstLineIndent = to_points(first_line_indent)
    if space_before is not None:
        pf.SpaceBefore = to_points(space_before)
    if space_after is not None:
        pf.SpaceAfter = to_points(space_after)
    if line_spacing is not None:
        rule, value = _coerce_line_spacing(line_spacing)
        # Set the value first: assigning LineSpacing forces the rule to Multiple,
        # so we write the rule last to land on Exactly/AtLeast/the named rules.
        if value is not None:
            pf.LineSpacing = value
        pf.LineSpacingRule = rule
    if page_break_before is not None:
        pf.PageBreakBefore = bool(page_break_before)
    if keep_together is not None:
        pf.KeepTogether = bool(keep_together)
    if keep_with_next is not None:
        pf.KeepWithNext = bool(keep_with_next)
    if widow_control is not None:
        pf.WidowControl = bool(widow_control)


# Reverse maps for the read mirror (`format_info`). Word stores alignment / line
# spacing as ints; these turn them back into the same keywords the write verbs
# accept, so read and write share one vocabulary. Unknown ints fall back to their
# string form rather than guessing.
_ALIGNMENT_BY_INT: dict[int, str] = {
    int(WdParagraphAlignment.LEFT): "left",
    int(WdParagraphAlignment.CENTER): "center",
    int(WdParagraphAlignment.RIGHT): "right",
    int(WdParagraphAlignment.JUSTIFY): "justify",
}


_LINE_SPACING_RULE_BY_INT: dict[int, str] = {
    int(WdLineSpacing.SINGLE): "single",
    int(WdLineSpacing.ONE_POINT_FIVE): "1.5",
    int(WdLineSpacing.DOUBLE): "double",
    int(WdLineSpacing.AT_LEAST): "at_least",
    int(WdLineSpacing.EXACTLY): "exactly",
    int(WdLineSpacing.MULTIPLE): "multiple",
}


# Word's "automatic" font colour sentinel (wdColorAutomatic) — a negative
# OLE_COLOR that `bgr_to_hex` (which expects 0x000000-0xFFFFFF) can't render, so
# `format_info` reports it as the keyword `"auto"`.
_WD_COLOR_AUTOMATIC = -16777216


def _pts(value: Any) -> float | None:
    """A `ParagraphFormat` length read as rounded points, or `None` if mixed.

    Word returns `WD_UNDEFINED` for a length that varies across a multi-paragraph
    range; surface that as `None` rather than a bogus 9999999.
    """
    v = float(value)
    if int(v) == WD_UNDEFINED:
        return None
    return round(v, 2)


def _tri(value: Any) -> bool | None:
    """A boolean `Font` property read as a tri-state: `True`/`False`, or `None`
    when Word reports `WD_UNDEFINED` (the property varies across runs)."""
    v = int(value)
    if v == WD_UNDEFINED:
        return None
    return bool(v)


def _line_spacing_repr(rule: int, value: float) -> str | None:
    """Render `(LineSpacingRule, LineSpacing)` as one comparable keyword/string.

    Mirrors `_coerce_line_spacing` in reverse: the named rules
    (`single`/`1.5`/`double`) drop their companion value; `multiple` renders as
    the multiple of single spacing (`13.8pt` -> `"1.15"`); `exactly`/`at_least`
    render as an exact length (`"14pt"` / `"at_least:14pt"`). `None` if the rule
    is unknown. One string keeps the `format_info` field shape uniform and the
    consistency-rule override compare a plain equality.
    """
    name = _LINE_SPACING_RULE_BY_INT.get(int(rule))
    if name in ("single", "1.5", "double"):
        return name
    if name == "multiple":
        return f"{round(float(value) / 12.0, 3):g}"
    if name == "exactly":
        return f"{round(float(value), 2):g}pt"
    if name == "at_least":
        return f"at_least:{round(float(value), 2):g}pt"
    return None


def _read_paragraph_format(pf: Any) -> dict[str, Any]:
    """Read a COM `ParagraphFormat` into the same field vocabulary `format_para-
    graph` writes — the effective values, no override annotation (that's added by
    `format_info` once it also has the style baseline)."""
    return {
        "alignment": _ALIGNMENT_BY_INT.get(int(pf.Alignment), str(int(pf.Alignment))),
        "left_indent": _pts(pf.LeftIndent),
        "right_indent": _pts(pf.RightIndent),
        "first_line_indent": _pts(pf.FirstLineIndent),
        "space_before": _pts(pf.SpaceBefore),
        "space_after": _pts(pf.SpaceAfter),
        "line_spacing": _line_spacing_repr(int(pf.LineSpacingRule), float(pf.LineSpacing)),
        "page_break_before": _tri(pf.PageBreakBefore),
        "keep_together": _tri(pf.KeepTogether),
        "keep_with_next": _tri(pf.KeepWithNext),
        "widow_control": _tri(pf.WidowControl),
    }


def _read_font(font: Any) -> tuple[dict[str, Any], list[str]]:
    """Read a COM `Font` into the same vocabulary `format_run` writes.

    Returns `(values, mixed)` — the effective character formatting, and the list
    of fields that read `WD_UNDEFINED` (i.e. vary across the range's runs). A
    mixed field's value is `None`; the field name appears in `mixed`. `color`
    renders as a `#RRGGBB` hex string, `"auto"` for Word's automatic colour, or
    `None` when mixed.
    """
    mixed: list[str] = []

    def _name() -> str | None:
        n = str(font.Name)
        # An empty name is Word's signal that the run fonts differ.
        if n == "":
            mixed.append("name")
            return None
        return n

    def _size() -> float | None:
        s = float(font.Size)
        if int(s) == WD_UNDEFINED:
            mixed.append("size")
            return None
        return round(s, 2)

    def _flag(field: str, raw: Any) -> bool | None:
        v = _tri(raw)
        if v is None:
            mixed.append(field)
        return v

    def _color() -> str | None:
        c = int(font.Color)
        if c == WD_UNDEFINED:
            mixed.append("color")
            return None
        if c < 0:  # wdColorAutomatic and friends are negative OLE_COLORs.
            return "auto"
        return bgr_to_hex(c)

    values = {
        "name": _name(),
        "size": _size(),
        "bold": _flag("bold", font.Bold),
        "italic": _flag("italic", font.Italic),
        "underline": _flag("underline", font.Underline),
        "strikethrough": _flag("strikethrough", font.StrikeThrough),
        "color": _color(),
        "subscript": _flag("subscript", font.Subscript),
        "superscript": _flag("superscript", font.Superscript),
        "small_caps": _flag("small_caps", font.SmallCaps),
        "all_caps": _flag("all_caps", font.AllCaps),
        "spacing": _pts(font.Spacing),
        "hidden": _flag("hidden", font.Hidden),
    }
    return values, mixed


# Border-side keywords -> the WdBorderType edges they cover. "all"/"box" hit the
# four outer edges; the named singles map to one edge each.
_BORDER_SIDES: dict[str, tuple[WdBorderType, ...]] = {
    "top": (WdBorderType.TOP,),
    "bottom": (WdBorderType.BOTTOM,),
    "left": (WdBorderType.LEFT,),
    "right": (WdBorderType.RIGHT,),
    "horizontal": (WdBorderType.HORIZONTAL,),
    "vertical": (WdBorderType.VERTICAL,),
    "all": (WdBorderType.TOP, WdBorderType.BOTTOM, WdBorderType.LEFT, WdBorderType.RIGHT),
    "box": (WdBorderType.TOP, WdBorderType.BOTTOM, WdBorderType.LEFT, WdBorderType.RIGHT),
}


_LINE_STYLES: dict[str, WdLineStyle] = {
    "none": WdLineStyle.NONE,
    "single": WdLineStyle.SINGLE,
    "dot": WdLineStyle.DOT,
    "dotted": WdLineStyle.DOT,
    "dash": WdLineStyle.DASH_LARGE_GAP,
    "dashed": WdLineStyle.DASH_LARGE_GAP,
    "dash-small": WdLineStyle.DASH_SMALL_GAP,
    "dash-dot": WdLineStyle.DASH_DOT,
    "dash-dot-dot": WdLineStyle.DASH_DOT_DOT,
    "double": WdLineStyle.DOUBLE,
}


_DROP_POSITIONS: dict[str, WdDropPosition] = {
    "none": WdDropPosition.NONE,
    "normal": WdDropPosition.DROPPED,
    "dropped": WdDropPosition.DROPPED,
    "margin": WdDropPosition.MARGIN,
}


_TAB_ALIGN: dict[str, WdTabAlignment] = {
    "left": WdTabAlignment.LEFT,
    "center": WdTabAlignment.CENTER,
    "centre": WdTabAlignment.CENTER,
    "right": WdTabAlignment.RIGHT,
    "decimal": WdTabAlignment.DECIMAL,
    "bar": WdTabAlignment.BAR,
}


_TAB_LEADERS: dict[str, WdTabLeader] = {
    "none": WdTabLeader.SPACES,
    "spaces": WdTabLeader.SPACES,
    "dots": WdTabLeader.DOTS,
    "dashes": WdTabLeader.DASHES,
    "lines": WdTabLeader.LINES,
    "heavy": WdTabLeader.HEAVY,
    "middle-dot": WdTabLeader.MIDDLE_DOT,
}


# Word's `Border.LineWidth` is a discrete `WdLineWidth` (points x 8). We accept a
# point value and snap to the nearest supported weight rather than rejecting the
# in-between ones, so `weight=1` (a hairline-ish 0.75pt..1pt) just works.
_LINE_WIDTHS: tuple[int, ...] = (2, 4, 6, 8, 12, 18, 24)  # 0.25, 0.5, 0.75, 1, 1.5, 2.25, 3 pt


def _resolve_border_sides(sides: Any) -> list[int]:
    """Map a `sides` argument (str or iterable of str) to WdBorderType ints."""
    if isinstance(sides, str):
        names = [sides]
    elif isinstance(sides, (list, tuple)):
        names = list(sides)
    else:
        raise TypeError(f"sides must be a string or list of strings; got {type(sides).__name__}")
    out: list[int] = []
    for n in names:
        key = str(n).lower()
        if key not in _BORDER_SIDES:
            raise ValueError(f"unknown border side {n!r}; expected one of {sorted(_BORDER_SIDES)}")
        for edge in _BORDER_SIDES[key]:
            if int(edge) not in out:
                out.append(int(edge))
    return out


# Line-spacing keywords -> the rule that needs no companion value.
_LINE_SPACING_NAMES: dict[str, WdLineSpacing] = {
    "single": WdLineSpacing.SINGLE,
    "1.5": WdLineSpacing.ONE_POINT_FIVE,
    "1.5x": WdLineSpacing.ONE_POINT_FIVE,
    "double": WdLineSpacing.DOUBLE,
}


def _coerce_line_spacing(value: Any) -> tuple[int, float | None]:
    """Map a `line_spacing` value to ``(LineSpacingRule, LineSpacing | None)``.

    - a **number** ``n`` → a multiple of single spacing (rule ``MULTIPLE``,
      ``LineSpacing = n × 12pt`` — Word's points-per-line);
    - ``"single"`` / ``"1.5"`` / ``"double"`` → the named multiple rules, with no
      companion value (``None``);
    - a **length string** carrying a unit (``"14pt"``, ``"1.5cm"``) → an *exact*
      line height (rule ``EXACTLY``, the value in points);
    - a unitless numeric string (``"1.5"``) → a multiple, as for a number.

    Raises `ValueError`/`TypeError` on anything else (translated to `OpError`).
    """
    if isinstance(value, bool):
        raise TypeError("line_spacing must be a number or string, not bool")
    if isinstance(value, (int, float)):
        return int(WdLineSpacing.MULTIPLE), float(value) * 12.0
    if isinstance(value, str):
        key = value.strip().lower()
        if key in _LINE_SPACING_NAMES:
            return int(_LINE_SPACING_NAMES[key]), None
        if any(key.endswith(u) for u in ("pt", "in", "cm", "mm")):
            return int(WdLineSpacing.EXACTLY), to_points(value)
        try:
            return int(WdLineSpacing.MULTIPLE), float(key) * 12.0
        except ValueError:
            raise ValueError(
                f"unknown line_spacing {value!r}; expected a number (a multiple of single "
                "spacing), one of single/1.5/double, or an exact length like '14pt'"
            ) from None
    raise TypeError(f"line_spacing must be a number or string; got {type(value).__name__}")


def _coerce_named(value: Any, table: dict[str, Any], label: str) -> int:
    if isinstance(value, int) and not isinstance(value, bool):
        return value
    if isinstance(value, str):
        try:
            return int(table[value.lower()])
        except KeyError:
            raise ValueError(
                f"unknown {label} {value!r}; expected one of {sorted(table)}"
            ) from None
    raise TypeError(f"{label} must be an int or str; got {type(value).__name__}")


def _coerce_line_weight(value: Any) -> int:
    """Points -> the nearest supported `WdLineWidth` (points x 8, snapped)."""
    pts = to_points(value)
    eighths = pts * 8
    return min(_LINE_WIDTHS, key=lambda w: abs(w - eighths))


def apply_borders(
    borders_com: Any,
    *,
    sides: Any = "all",
    style: Any = "single",
    weight: Any = 0.5,
    color: Any = None,
) -> None:
    """Apply border styling to a COM `Borders` collection (range, cell, or table).

    Shared by `Anchor.set_borders` (per-range / per-cell borders) and
    `Table.set_borders` (whole-grid borders): both index the same
    `Borders(WdBorderType)` shape, so the coerce-then-loop is identical. Coercion
    happens first, before any COM write, so a bad `sides`/`style`/`weight`/`color`
    raises `ValueError`/`TypeError` (the caller wraps it in `OpError`) without a
    partial edit.
    """
    edges = _resolve_border_sides(sides)
    line_style = _coerce_named(style, _LINE_STYLES, "border style")
    line_width = _coerce_line_weight(weight)
    bgr = to_bgr(color) if color is not None else None
    for edge in edges:
        b = borders_com(edge)
        b.LineStyle = line_style
        b.LineWidth = line_width
        if bgr is not None:
            b.Color = bgr


# Floating wrap keywords -> WdWrapType. "inline" and "auto" are handled
# specially by insert_image and are not in this map. Single source of truth:
# `_shapes.WRAP_NAMES` (the floating-shape mutators share the same vocabulary).
_WRAP_NAMES: dict[str, WdWrapType] = _shapes.WRAP_NAMES


_WRAP_VALUES: frozenset[str] = frozenset({"inline", "auto", *_WRAP_NAMES})


# Break keywords -> WdBreakType. `insert_break(kind=...)` accepts exactly these.
_BREAK_TYPES: dict[str, WdBreakType] = {
    "page": WdBreakType.PAGE,
    "column": WdBreakType.COLUMN,
    "section_next": WdBreakType.SECTION_NEXT_PAGE,
    "section_continuous": WdBreakType.SECTION_CONTINUOUS,
}


# Field keywords -> WdFieldType. `insert_field(kind=...)` accepts these names;
# "field" is the raw-code escape hatch (an empty field whose code is the `text`).
_FIELD_TYPES: dict[str, WdFieldType] = {
    "page": WdFieldType.PAGE,
    "numpages": WdFieldType.NUM_PAGES,
    "date": WdFieldType.DATE,
    "time": WdFieldType.TIME,
    "filename": WdFieldType.FILE_NAME,
    "author": WdFieldType.AUTHOR,
    "title": WdFieldType.TITLE,
    "field": WdFieldType.EMPTY,
}


def _resolve_wrap(wrap: str, inline_shape: Any, insert_rng: Any) -> WdWrapType:
    """Resolve a wrap keyword to a concrete `WdWrapType` for a floating shape.

    `"auto"` picks Square when the image is at most half the section's usable
    text width (`PageWidth - LeftMargin - RightMargin`), else top-and-bottom.
    """
    if wrap != "auto":
        return _WRAP_NAMES[wrap]
    ps = insert_rng.PageSetup
    usable = float(ps.PageWidth) - float(ps.LeftMargin) - float(ps.RightMargin)
    if float(inline_shape.Width) <= usable / 2:
        return WdWrapType.SQUARE
    return WdWrapType.TOP_BOTTOM


def _validate_table_data(data: Any, rows: int, cols: int) -> None:
    """Check a row-major `data` payload fits a `rows` × `cols` grid.

    Raised before any COM call so a bad shape is a clean `OpError` (exit 1)
    rather than a "subscript out of range" deep inside Word. Underfilling is
    allowed — fewer rows, or short rows — and leaves the trailing cells empty
    (matching `add_row`'s leniency); only *overflowing* the declared grid is an
    error, since that's the case that would otherwise blow up mid-insert.
    """
    if not isinstance(data, list):
        raise OpError(f"table data must be a list of rows; got {type(data).__name__}")
    if len(data) > rows:
        raise OpError(f"table data has {len(data)} rows but the table has only {rows}")
    for i, row in enumerate(data, start=1):
        if not isinstance(row, list):
            raise OpError(f"table data row {i} must be a list; got {type(row).__name__}")
        if len(row) > cols:
            raise OpError(
                f"table data row {i} has {len(row)} cells but the table has only {cols} column(s)"
            )


def _normalize_table_data(data: Any) -> tuple[list[list[Any]], bool]:
    """Coerce a table `data` payload into a row-major grid + a header flag.

    Two shapes are accepted so the caller can hand over tabular data however it
    has it:

    - **2-D array** — ``[[r1c1, r1c2], …]`` is returned unchanged with
      ``header=False`` (the caller's own `header` choice then wins).
    - **Records** — ``[{col: val, …}, …]`` (a list of dicts): the first record's
      keys become the header row and each dict contributes a body row, so the
      grid comes back ``header=True``. The first record fixes the column order;
      later records fill by key (a missing key is an empty cell, extra keys are
      ignored). This is the natural LLM "rows of objects" shape.

    Raises `OpError` for a non-list payload or a list that mixes dict and list
    rows (an ambiguous shape).
    """
    if not isinstance(data, list):
        raise OpError(f"table data must be a list; got {type(data).__name__}")
    if not data:
        return [], False
    dict_rows = [isinstance(row, dict) for row in data]
    if all(dict_rows):
        columns = list(data[0].keys())
        grid: list[list[Any]] = [list(columns)]
        grid.extend([rec.get(col, "") for col in columns] for rec in data)
        return grid, True
    if any(dict_rows):
        raise OpError("table data mixes object rows and array rows; use one shape")
    return data, False


def _within_table(doc_com: Any, start: int, end: int) -> bool:
    """Whether the `[start, end)` span sits inside a table.

    Used to detect when a new table's insertion point abuts an existing one —
    Word silently *merges* two tables with no paragraph mark between them, so
    `insert_table` drops a separator paragraph on any abutting side. A negative
    `start` (before the document) or a probe Word rejects reads as "not in a
    table".
    """
    if start < 0:
        return False
    try:
        return bool(doc_com.Range(start, end).Information(int(WdInformation.WITH_IN_TABLE)))
    except Exception:
        return False


def _final_paragraph_empty(doc_com: Any, final_mark: int) -> bool:
    """Whether the document's terminal paragraph holds no text (just its mark).

    Used by `insert_block` to decide, when appending at the very end, whether to
    *fill* the final paragraph (it's empty — the fresh-document case) or open a
    new one after it (it already has text — don't merge into it). If the probe
    can't read the paragraph's offsets (e.g. a stubbed COM in tests), default to
    "not empty" — the safe choice, which opens a new paragraph and never merges.
    """
    try:
        para = doc_com.Range(final_mark, final_mark).Paragraphs(1).Range
        return int(para.End) - int(para.Start) <= 1
    except (TypeError, ValueError, AttributeError):
        return False


def _equation_index_at(doc_com: Any, pos: int) -> int:
    """1-based document index of the OMath at `pos` — the just-inserted equation.

    Equations are addressed `equation:N` in document order, matching Word's own
    `OMaths(n)`. The native path probes a position *inside* the new zone, so a
    containing-zone match wins; the OMML path probes the insertion boundary, so
    we fall back to the first zone starting at or after `pos`. A document with no
    equations (shouldn't happen right after an insert) reports index 1.
    """
    omaths = doc_com.OMaths
    count = int(omaths.Count)
    for i in range(1, count + 1):
        rng = omaths.Item(i).Range
        if int(rng.Start) <= pos < int(rng.End):
            return i
    for i in range(1, count + 1):
        if int(omaths.Item(i).Range.Start) >= pos:
            return i
    return max(1, count)


def _markdown_segments(blocks: list[Any]) -> list[tuple[list[dict[str, Any]], str | None]]:
    """Group parsed Markdown `blocks` into `(insert_block_items, list_type)` runs.

    A maximal run of same-kind list items becomes one segment whose `list_type`
    is `"bulleted"`/`"numbered"` (applied over the span after insertion). A
    maximal run of heading/normal blocks becomes one segment with `list_type`
    `None`. Normal paragraphs are pinned to the built-in ``Normal`` style so they
    don't inherit a heading style from the insertion point.
    """
    from .._markdown import BULLET, HEADING, NORMAL, NUMBER

    segments: list[tuple[list[dict[str, Any]], str | None]] = []
    i, n = 0, len(blocks)
    while i < n:
        kind = blocks[i].kind
        items: list[dict[str, Any]] = []
        if kind in (BULLET, NUMBER):
            style = "List Bullet" if kind == BULLET else "List Number"
            list_type = "bulleted" if kind == BULLET else "numbered"
            while i < n and blocks[i].kind == kind:
                items.append({"text": blocks[i].text, "style": style})
                i += 1
            segments.append((items, list_type))
        else:
            while i < n and blocks[i].kind in (HEADING, NORMAL):
                b = blocks[i]
                style = f"Heading {b.level}" if b.kind == HEADING else "Normal"
                items.append({"text": b.text, "style": style})
                i += 1
            segments.append((items, None))
    return segments


def _chart_index_at(doc_com: Any, start: int) -> int:
    """1-based index, among the document's charts, of the chart at `start`.

    Charts are addressed `chart:N` in document order (over `HasChart` inline
    shapes). After an insert, the new chart sits at character position `start`;
    its index is the position of the first chart at or after `start`.
    """
    shapes = _charts.chart_shapes(doc_com)
    for i, shape in enumerate(shapes, start=1):
        try:
            if int(shape.Range.Start) >= start:
                return i
        except Exception:
            continue
    return len(shapes)


def _safe_float(obj: Any, attr: str) -> float | None:
    try:
        return float(getattr(obj, attr))
    except Exception:
        return None


def _safe_str(obj: Any, attr: str) -> str:
    try:
        return str(getattr(obj, attr) or "")
    except Exception:
        return ""


# ---------------------------------------------------------------------------
# Headings
# ---------------------------------------------------------------------------


def paragraph_text(para: Any) -> str:
    """Heading text minus the trailing paragraph mark, inline shapes tokenized."""
    return range_text(para.Range).rstrip("\r\n\x07")
