"""Anchor types — semantic handles for ranges inside a Word document.

Anchors target a `Range`, never the live `Selection`. Each public mutation
goes through the COM error translator. Operations are intentionally small;
they compose with `Document.edit()` for atomic-undo behaviour.
"""

from __future__ import annotations

import re
import secrets
from abc import ABC, abstractmethod
from collections.abc import Iterator
from typing import TYPE_CHECKING, Any

from . import _charts, _com, _equations, _images, _lists, _revisions, _shapes
from ._format import bgr_to_hex, to_bgr, to_points
from .constants import (
    WD_UNDEFINED,
    MsoTextOrientation,
    MsoTriState,
    WdBorderType,
    WdBreakType,
    WdCaptionPosition,
    WdCollapseDirection,
    WdColorIndex,
    WdContentControlType,
    WdDropPosition,
    WdFieldType,
    WdIndexType,
    WdInformation,
    WdLineSpacing,
    WdLineStyle,
    WdNumberType,
    WdParagraphAlignment,
    WdReferenceKind,
    WdReferenceType,
    WdTabAlignment,
    WdTabLeader,
    WdUnderline,
    WdWrapType,
)
from .exceptions import AnchorNotFoundError, EquationError, ExcelNotAvailableError, OpError

if TYPE_CHECKING:
    from pathlib import Path

    from ._document import Document
    from ._snapshot import Snapshot


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


def range_text(rng: Any) -> str:
    """Read a COM range's text with inline shapes surfaced as ``[image]`` tokens.

    Word represents each inline shape (embedded picture / OLE object) as a single
    placeholder character in the text stream. That character is *not* a reserved
    control code — it varies by build and is indistinguishable by value from real
    text (a forward slash, on some Word versions) — so a naive string replace
    would clobber genuine characters. Instead we locate the shapes via the
    ``InlineShapes`` collection and swap only the character at each shape's own
    position, leaving real text untouched. A range with no inline shapes returns
    its raw text unchanged.
    """
    raw = str(rng.Text or "")
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


# Bookmark names: must start with a letter, then letters/digits/underscores, and
# Word caps them at 40 characters. (Leading-underscore names are Word's hidden
# internal bookmarks, so user-created ones must lead with a letter.)
_BOOKMARK_NAME_RE = re.compile(r"^[A-Za-z][A-Za-z0-9_]*$")


def _validate_bookmark_name(name: Any) -> str:
    """Validate a new bookmark name against Word's rules. Raises `OpError`."""
    if not isinstance(name, str) or not name:
        raise OpError("bookmark name must be a non-empty string")
    if len(name) > 40:
        raise OpError(f"bookmark name must be at most 40 characters; got {len(name)}")
    if not _BOOKMARK_NAME_RE.match(name):
        raise OpError(
            f"invalid bookmark name {name!r}: must start with a letter and contain only "
            "letters, digits, and underscores (no spaces)"
        )
    return name


# Durable handles ("pins") are minted as Word-hidden bookmarks named `_wl_<code>`.
# The leading underscore makes Word treat them as internal (so they stay out of
# `bookmarks.list()`), and Word maintains the bookmark<->range association across
# inserts / deletes / edits natively — that native behaviour is what makes a
# `pin:` anchor durable where a positional `para:N` is not.
_WL_PREFIX = "_wl_"

# A user-supplied pin slug: lowercase alphanumeric words joined by single
# hyphens. No underscores — storage maps `-` -> `_`, so a slug carrying `_`
# would round-trip ambiguously.
_PIN_SLUG_RE = re.compile(r"^[a-z0-9]+(?:-[a-z0-9]+)*$")


def _new_pin_code() -> str:
    """A fresh random pin code — six lowercase hex chars (e.g. ``a3f9c2``).

    Random rather than sequential so codes don't leak document structure and
    two pins never collide by construction. Tests monkeypatch this for
    deterministic ids.
    """
    return secrets.token_hex(3)


def _pin_name_for(code: str) -> str:
    """Bookmark storage name for a pin display code (`pin:<code>` -> `_wl_…`)."""
    return _WL_PREFIX + code.replace("-", "_")


def _pin_id_for(name: str) -> str:
    """Pin display code for a `_wl_` bookmark name (inverse of `_pin_name_for`)."""
    return name[len(_WL_PREFIX) :].replace("_", "-")


def _validate_pin_slug(slug: Any) -> str:
    """Validate a user-supplied pin slug and return it. Raises `OpError`.

    A slug is lowercase alphanumeric words joined by single hyphens
    (``budget-intro``), and its stored bookmark name (`_wl_` + slug with hyphens
    mapped to underscores) must fit Word's 40-character bookmark-name cap.
    """
    if not isinstance(slug, str) or not slug:
        raise OpError("pin name must be a non-empty string")
    if not _PIN_SLUG_RE.match(slug):
        raise OpError(
            f"invalid pin name {slug!r}: use lowercase letters, digits, and single "
            "hyphens (e.g. 'budget-intro')"
        )
    storage = _pin_name_for(slug)
    if len(storage) > 40:
        raise OpError(
            f"pin name {slug!r} is too long: the stored bookmark name {storage!r} "
            "exceeds Word's 40-character limit"
        )
    return slug


def _mint_wl_bookmark(doc_com: Any, rng: Any, code: str) -> str:
    """Plant a hidden `_wl_<code>` bookmark over `rng`; return the bookmark name.

    Bypasses `_validate_bookmark_name` (which forbids the leading underscore) on
    purpose — these are wordlive-internal names, not user-chosen ones. Word still
    enforces its own 40-char cap, which `_validate_pin_slug` checks up front.
    """
    name = _pin_name_for(code)
    doc_com.Bookmarks.Add(Name=name, Range=rng)
    return name


def _stale_anchor_hint(doc_com: Any, kind: str, index: int) -> str | None:
    """A recovery hint for a positional `para:N` / `heading:N` that missed.

    Positional ids renumber under inserts / deletes, so a stale one is the
    common failure. One cheap pass over `Paragraphs` reports whether the index
    is out of range (renumbered / vanished) or — for `heading:N` — points at
    body text, and names the nearest heading. Always recommends pinning a
    durable handle. Returns `None` if the document can't be read (best-effort —
    the hint must never mask the original miss).
    """
    try:
        paras = list(doc_com.Paragraphs)
    except Exception:
        return None
    count = len(paras)
    pin_tip = "Pin a durable handle with `pin` to survive renumbering."

    def _level(p: Any) -> int:
        try:
            return int(p.OutlineLevel)
        except Exception:
            return 10

    def _nearest_heading() -> tuple[int, str] | None:
        best: tuple[int, int, str] | None = None  # (distance, idx, text)
        for i, p in enumerate(paras, start=1):
            if i == index:
                continue
            if _level(p) < 10:
                dist = abs(i - index)
                if best is None or dist < best[0]:
                    best = (dist, i, paragraph_text(p))
        return (best[1], best[2]) if best is not None else None

    nh = _nearest_heading()
    if index < 1 or index > count:
        base = f"{kind}:{index} out of range; document has {count} paragraph(s)"
        base += (
            f' (nearest heading is heading:{nh[0]} "{nh[1]}")'
            if nh
            else " (likely renumbered by an edit)"
        )
        return f"{base}. {pin_tip}"
    if kind == "heading":
        # In range but not a heading (body text at this index).
        base = f"heading:{index} is not a heading (body text)"
        if nh:
            base += f'; nearest heading is heading:{nh[0]} "{nh[1]}"'
        return f"{base}. Use a pin for a stable handle."
    return pin_tip


def _resolve_cross_ref_target(doc: Document, target: str) -> tuple[int, int | str]:
    """Map a cross-reference `target` anchor id to `(ReferenceType, ReferenceItem)`.

    `ReferenceItem` is what Word's `InsertCrossReference` expects, which differs
    by type: for `bookmark:NAME` it is the **bookmark name string** (Word looks
    it up by name, not position); for `heading:N` it is the 1-based ordinal among
    heading paragraphs; for `footnote:N` / `endnote:N` it is the 1-based index.
    Raises `AnchorNotFoundError` (exit 2) for an unknown/unresolvable target.
    """
    kind, _, value = str(target).partition(":")
    if kind == "bookmark":
        try:
            items = [
                str(it).strip()
                for it in doc.com.GetCrossReferenceItems(int(WdReferenceType.BOOKMARK))
            ]
        except Exception as e:
            raise AnchorNotFoundError("bookmark", target) from e
        if value not in items:
            raise AnchorNotFoundError("bookmark", target)
        # For bookmarks the ReferenceItem is the *name*, not a position index.
        return int(WdReferenceType.BOOKMARK), value
    if kind == "heading":
        try:
            n = int(value)
        except ValueError as e:
            raise AnchorNotFoundError("heading", target) from e
        ordinal = 0
        for idx, para in enumerate(doc.com.Paragraphs, start=1):
            try:
                level = int(para.OutlineLevel)
            except Exception:
                level = 10
            if level < 10:
                ordinal += 1
            if idx == n:
                if level >= 10:
                    raise AnchorNotFoundError("heading", target)
                return int(WdReferenceType.HEADING), ordinal
        raise AnchorNotFoundError("heading", target)
    if kind in ("footnote", "endnote"):
        try:
            n = int(value)
        except ValueError as e:
            raise AnchorNotFoundError(kind, target) from e
        ref_type = WdReferenceType.FOOTNOTE if kind == "footnote" else WdReferenceType.ENDNOTE
        coll = doc.footnotes if kind == "footnote" else doc.endnotes
        if not (1 <= n <= len(coll)):
            raise AnchorNotFoundError(kind, target)
        return int(ref_type), n
    raise AnchorNotFoundError(
        "cross-reference target",
        target,
        hint="target must be a bookmark:, heading:, footnote:, or endnote: anchor id",
    )


def _cross_ref_kind(kind: str, ref_type: int) -> int:
    """Map an `insert_cross_reference(kind=...)` string to a `WdReferenceKind`.

    Type-dependent in two places: a note has no "text" content to reference, so
    `"text"`/`"number"` both resolve to its number; for headings/bookmarks
    `"number"` is the paragraph number. Raises `ValueError` on an unknown kind.
    """
    is_note = ref_type in (int(WdReferenceType.FOOTNOTE), int(WdReferenceType.ENDNOTE))
    note_number = (
        int(WdReferenceKind.FOOTNOTE_NUMBER)
        if ref_type == int(WdReferenceType.FOOTNOTE)
        else int(WdReferenceKind.ENDNOTE_NUMBER)
    )
    if kind == "text":
        # wdContentText is invalid for footnotes/endnotes (a mark has no text);
        # fall back to the note number, which is the meaningful reference.
        return note_number if is_note else int(WdReferenceKind.CONTENT_TEXT)
    if kind == "page":
        return int(WdReferenceKind.PAGE_NUMBER)
    if kind == "above_below":
        return int(WdReferenceKind.POSITION)
    if kind == "number":
        return note_number if is_note else int(WdReferenceKind.NUMBER_NO_CONTEXT)
    raise ValueError(
        f"unknown cross-reference kind {kind!r}; expected text/page/number/above_below"
    )


def _caption_above(label: str, position: str | None) -> bool:
    """Resolve a caption's placement to a boolean (`True` = above the anchor).

    `position` is the user override (``"above"``/``"below"``, with
    ``"before"``/``"after"`` accepted as aliases). When it's `None` the
    *convention* applies: a ``"Table"`` caption goes **above**, every other
    label (Figure, Equation, …) goes **below**. Raises `ValueError` on a bad
    string.
    """
    if position is None:
        return str(label).strip().casefold() == "table"
    p = str(position).strip().casefold()
    if p in ("above", "before", "top"):
        return True
    if p in ("below", "after", "bottom"):
        return False
    raise ValueError(f"position must be 'above' or 'below'; got {position!r}")


def _markdown_segments(blocks: list[Any]) -> list[tuple[list[dict[str, Any]], str | None]]:
    """Group parsed Markdown `blocks` into `(insert_block_items, list_type)` runs.

    A maximal run of same-kind list items becomes one segment whose `list_type`
    is `"bulleted"`/`"numbered"` (applied over the span after insertion). A
    maximal run of heading/normal blocks becomes one segment with `list_type`
    `None`. Normal paragraphs are pinned to the built-in ``Normal`` style so they
    don't inherit a heading style from the insertion point.
    """
    from ._markdown import BULLET, HEADING, NORMAL, NUMBER

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


# ---------------------------------------------------------------------------
# Base
# ---------------------------------------------------------------------------


class Anchor(ABC):
    """Abstract base — subclasses know how to materialise their COM Range.

    Concrete subclasses must implement `_range()` and `set_text()`. Other
    operations (`text`, `insert_before`, `insert_after`, `delete`,
    `apply_style`, `format_paragraph`) are derived and inherited as-is.
    """

    kind: str = "anchor"
    name: str = ""

    def __init__(self, doc: Document, name: str) -> None:
        self._doc = doc
        self.name = name

    @property
    def com(self) -> Any:
        """Raw COM range. Subclasses override."""
        return self._range()

    @abstractmethod
    def _range(self) -> Any:
        """Return the COM Range that this anchor refers to. Must be overridden."""

    def _caption_object_range(self) -> Any | None:
        """Return a Range selecting a caption-able *object* for `insert_caption`.

        Word's `InsertCaption` only honours its above/below `Position` when the
        range selects a real object — a whole `Table`, an `InlineShape`, or a
        floating `Shape`. A plain text anchor isn't one, so the base returns
        `None` (the caption gets its own paragraph instead); `Cell` overrides
        this to return its parent table's range so a table caption lands above /
        below the table rather than inside a cell.
        """
        return None

    @property
    def text(self) -> str:
        with _com.translate_com_errors():
            return range_text(self._range())

    def _revision_runs(self, start: int, end: int) -> list[dict[str, Any]]:
        """`{change, start, end, text}` for each insert/delete revision overlapping `[start, end)`."""
        runs: list[dict[str, Any]] = []
        for row in self._doc.revisions.list():
            change = row["type"]
            if change not in ("insert", "delete"):
                continue
            r_start, r_end = int(row["start"]), int(row["end"])
            if r_end <= start or r_start >= end:
                continue
            runs.append({"change": change, "start": r_start, "end": r_end, "text": row["text"]})
        return runs

    def revision_segments(self) -> list[dict[str, Any]]:
        """The anchor's text split into tracked-change segments (revision-aware read).

        Returns `[{text, change}, …]` in document order, where `change` is
        ``"insert"``, ``"delete"``, or ``None`` (unchanged). Word's `text` read
        shows the *final* view (inserted runs present, deleted runs gone); this
        also surfaces the deleted runs, so you can see both sides of a tracked
        edit. [`text_final`][wordlive.Anchor.text_final] and
        [`text_original`][wordlive.Anchor.text_original] are the two flattened
        views. The structured, whole-document counterpart is `doc.revisions`.
        """
        with _com.translate_com_errors():
            rng = self._range()
            start, end = int(rng.Start), int(rng.End)
            final_text = range_text(rng)
        return _revisions.segment_runs(final_text, start, self._revision_runs(start, end))

    @property
    def text_final(self) -> str:
        """The anchor's text **as if every tracked change in it were accepted**.

        Inserted runs stay, deleted runs drop — the after-the-edits view. Equal to
        `text` when nothing tracked touches the range. The mirror is
        [`text_original`][wordlive.Anchor.text_original]; the per-segment breakdown
        is [`revision_segments`][wordlive.Anchor.revision_segments].
        """
        return "".join(s["text"] for s in self.revision_segments() if s["change"] != "delete")

    @property
    def text_original(self) -> str:
        """The anchor's text **as if every tracked change in it were rejected**.

        Deleted runs stay, inserted runs drop — the before-the-edits view. The
        mirror of [`text_final`][wordlive.Anchor.text_final].
        """
        return "".join(s["text"] for s in self.revision_segments() if s["change"] != "insert")

    @property
    @abstractmethod
    def anchor_id(self) -> str:
        """Stable string identifier for this anchor (e.g. `bookmark:Address`).

        Each anchor kind has its own scheme (`bookmark:`, `cc:`, `heading:`),
        so subclasses must declare theirs explicitly — no useful default
        exists at this level.
        """

    @abstractmethod
    def set_text(self, text: str) -> None:
        """Replace the anchor's text in place. Must be overridden."""

    def insert_before(self, text: str) -> None:
        with _com.translate_com_errors():
            rng = self._range()
            insert_rng = self._doc.com.Range(rng.Start, rng.Start)
            insert_rng.Text = text

    def insert_after(self, text: str) -> None:
        with _com.translate_com_errors():
            rng = self._range()
            insert_rng = self._doc.com.Range(rng.End, rng.End)
            insert_rng.Text = text

    def insert_paragraph_before(self, text: str, style: str | None = None) -> None:
        """Insert a new paragraph immediately before this anchor's range.

        If `style` is given it must name a style defined in the document;
        otherwise `StyleNotFoundError` is raised before any text is inserted.
        """
        style_obj = self._doc.styles[style] if style is not None else None
        with _com.translate_com_errors():
            doc_com = self._doc.com
            start = int(self._range().Start)
            insert_rng = doc_com.Range(start, start)
            insert_rng.Text = text + "\r"
            if style_obj is not None:
                # Word measures Range offsets in UTF-16 code units; Python's
                # len() under-counts surrogate pairs and leaves the tail unstyled.
                styled = doc_com.Range(start, start + _utf16_len(text))
                styled.Style = style_obj.com

    def insert_paragraph_after(self, text: str, style: str | None = None) -> None:
        """Insert a new paragraph immediately after this anchor's range.

        If `style` is given it must name a style defined in the document;
        otherwise `StyleNotFoundError` is raised before any text is inserted.

        When the anchor is (or ends at) the document's final paragraph there is
        no position *after* the terminal paragraph mark to write to — Word
        rejects `Range(end, end)` there with a "value out of range" COM error.
        In that case the new paragraph is split in just before the final mark
        instead, so appending to the end of a document — the common
        "build from scratch" case, where the only paragraph *is* the last one —
        just works.
        """
        style_obj = self._doc.styles[style] if style is not None else None
        with _com.translate_com_errors():
            doc_com = self._doc.com
            end = int(self._range().End)
            doc_end = int(doc_com.Content.End)
            if end >= doc_end:
                # Anchor ends at the final paragraph mark. Insert "<break><text>"
                # just before that mark: the leading break terminates the
                # anchor's paragraph and `text` becomes a new final paragraph
                # (the original final mark now closes it).
                anchor_pos = max(0, doc_end - 1)
                insert_rng = doc_com.Range(anchor_pos, anchor_pos)
                insert_rng.Text = "\r" + text
                text_start = anchor_pos + 1
            else:
                insert_rng = doc_com.Range(end, end)
                insert_rng.Text = text + "\r"
                text_start = end
            if style_obj is not None:
                # Word measures Range offsets in UTF-16 code units; Python's
                # len() under-counts surrogate pairs and leaves the tail unstyled.
                styled = doc_com.Range(text_start, text_start + _utf16_len(text))
                styled.Style = style_obj.com

    def insert_block(self, items: list[Any], *, where: str = "after") -> RangeAnchor:
        """Insert a contiguous run of styled paragraphs at this anchor, atomically.

        The multi-paragraph counterpart to `insert_paragraph_after` — drop a
        whole styled section (a feature list, a set of bullets, a heading plus
        its body) in **one** op, in natural reading order. Inserting paragraphs
        one at a time forces a reverse-order dance to dodge positional-anchor
        renumbering; this places them all at a single point so order is just the
        order of `items`.

        Each item is one paragraph, given as either a plain string or a dict:

        - ``"some text"`` — sugar for ``{"text": "some text"}``.
        - ``{"text": "**Bold lead** — rest", "style": "List Bullet"}`` — `text`
          carries the tiny inline markdown (`**bold**`, `*italic*`,
          `***both***`; escape a literal asterisk as ``\\*``), and `style` names
          the paragraph style.
        - ``{"runs": [{"text": "Bold lead", "bold": true}, {"text": " — rest"}],
          "style": "List Bullet"}`` — the structured form: each run is
          ``{text, bold?, italic?, underline?, style?}`` (a per-run character
          style). Use it when markup is ambiguous or you need a run `style`.

        Returns a [`RangeAnchor`][wordlive.RangeAnchor] spanning the inserted
        block (`range:START-END`), so a follow-up op can target the whole run —
        e.g. `apply_list` it into a bulleted section, or comment on it. `where`
        is ``"after"`` (default) or ``"before"`` this anchor's range. Resolves
        every paragraph/run style up front, so an unknown style name raises
        `StyleNotFoundError` before any text is inserted. Wrap in `doc.edit(...)`
        for atomic undo. Raises `OpError` for a malformed `items` payload.
        """
        from ._runs import normalize_block_items, runs_to_text

        if where not in ("before", "after"):
            raise ValueError(f"where must be 'before' or 'after'; got {where!r}")
        norm = normalize_block_items(items)
        # Resolve every paragraph + run style before touching the document, so a
        # bad name fails the whole block cleanly rather than leaving a partial,
        # half-styled run behind.
        para_styles = [self._doc.styles[s] if s else None for _, s in norm]
        run_styles: dict[str, Any] = {}
        for runs, _ in norm:
            for r in runs:
                if r.style and r.style not in run_styles:
                    run_styles[r.style] = self._doc.styles[r.style]
        para_texts = [runs_to_text(runs) for runs, _ in norm]
        joined = "\r".join(para_texts)
        with _com.translate_com_errors():
            doc_com = self._doc.com
            if where == "before":
                start = int(self._range().Start)
                doc_com.Range(start, start).Text = joined + "\r"
            else:
                end = int(self._range().End)
                doc_end = int(doc_com.Content.End)
                final_mark = max(0, doc_end - 1)
                if end >= final_mark:
                    # We're at the document's terminal paragraph mark — the one
                    # position you can't write *past* (`doc.end` resolves here,
                    # as does an anchor ending at the last paragraph). The old
                    # `end >= doc_end` guard missed `doc.end` (whose range ends at
                    # doc_end - 1), so the block was written *before* the final
                    # mark and merged into a non-empty last paragraph — stealing
                    # its style too. Decide by whether that final paragraph holds
                    # text, so the block neither merges nor leaves a stray empty:
                    if _final_paragraph_empty(doc_com, final_mark):
                        # Reuse the empty final paragraph as the block's last one:
                        # write without a trailing break so the existing mark
                        # closes it (no leftover empty paragraph).
                        doc_com.Range(final_mark, final_mark).Text = joined
                        start = final_mark
                    else:
                        # Open a fresh paragraph after the final one: the leading
                        # break terminates it and the block becomes new trailing
                        # paragraphs (the original final mark closes the last one).
                        doc_com.Range(final_mark, final_mark).Text = "\r" + joined
                        start = final_mark + 1
                else:
                    doc_com.Range(end, end).Text = joined + "\r"
                    start = end
            span_end = start + _utf16_len(joined)
            # Paragraph styling and run formatting both preserve text length, so
            # offsets stay valid throughout: walk the block by deterministic
            # UTF-16 offset (Word counts code units) rather than re-querying
            # Paragraphs() after each mutation.
            off = start
            for (runs, _), style_obj, ptext in zip(norm, para_styles, para_texts, strict=True):
                plen = _utf16_len(ptext)
                if style_obj is not None:
                    doc_com.Range(off, off + plen).Paragraphs(1).Range.Style = style_obj.com
                roff = off
                for run in runs:
                    rlen = _utf16_len(run.text)
                    if run.formatted():
                        sub = doc_com.Range(roff, roff + rlen)
                        if run.bold is not None:
                            sub.Bold = bool(run.bold)
                        if run.italic is not None:
                            sub.Italic = bool(run.italic)
                        if run.underline is not None:
                            sub.Underline = 1 if run.underline else 0
                        if run.style:
                            sub.Style = run_styles[run.style].com
                    roff += rlen
                off += plen + 1  # + the paragraph mark (CR)
        return RangeAnchor(self._doc, start, span_end)

    def insert_section(
        self, heading: str, body: Any, *, level: int = 1, where: str = "after"
    ) -> RangeAnchor:
        """Insert a heading plus its body in one atomic op.

        The opinionated common case over `insert_block`: a single
        ``Heading {level}`` paragraph followed by `body`, placed in reading
        order at one point. `heading` carries the same inline markdown a block
        item's `text` does (`**bold**`, `*italic*`); `body` is the `insert_block`
        items shape — a list of plain strings or ``{text|runs, style?}`` dicts
        (a bare string is sugar for a one-paragraph body). `level` is 1–9 and
        selects the built-in ``Heading {level}`` style (validated before any
        mutation; an absent style raises `StyleNotFoundError` via `insert_block`).

        Returns the section's spanning [`RangeAnchor`][wordlive.RangeAnchor]
        (`range:START-END`). Wrap in `doc.edit(...)` for atomic undo.
        """
        if not isinstance(level, int) or isinstance(level, bool) or not 1 <= level <= 9:
            raise ValueError(f"level must be an integer 1–9; got {level!r}")
        if isinstance(body, str):
            body = [body]
        if not isinstance(body, list):
            raise OpError(
                f"insert_section body must be a string or list; got {type(body).__name__}"
            )
        items = [{"text": heading, "style": f"Heading {level}"}, *body]
        return self.insert_block(items, where=where)

    def insert_markdown(self, md: str, *, where: str = "after") -> RangeAnchor:
        """Insert a constrained-Markdown block as real Word structure, atomically.

        Maps a deliberately tiny block dialect (see `_markdown`) to paragraphs,
        headings, and lists: ``#``/``##``/``###`` → `Heading 1/2/3`, ``-``/``*``
        → a bulleted list, ``1.`` → a numbered list, blank-line-separated text →
        `Normal` paragraphs, with inline ``**bold**``/``*italic*`` spans honoured.
        It is **a subset, not CommonMark** — no code fences, nested lists, block
        quotes, or tables in v1; anything unrecognised is literal paragraph text.

        The whole block is one `insert_block` (one contiguous write); each
        same-kind list run is then `apply_list`-ed over its own span, so a
        numbered list reads 1..N. `where` is ``"after"`` (default) or ``"before"``
        this anchor's range. Returns the [`RangeAnchor`][wordlive.RangeAnchor]
        spanning everything inserted. Raises `OpError` for empty markdown.
        """
        from ._markdown import parse_markdown
        from ._runs import normalize_block_items, runs_to_text

        blocks = parse_markdown(md)
        if not blocks:
            raise OpError("insert_markdown requires non-empty markdown")
        # Flatten every block into ONE insert_block (a single contiguous write —
        # chaining separate inserts would land each list before the previous
        # block's paragraph mark and merge them). Record which paragraph runs are
        # lists so we can apply_list over their spans afterwards.
        segments = _markdown_segments(blocks)
        items: list[dict[str, Any]] = []
        list_groups: list[tuple[int, int, str]] = []  # (first_para, last_para, list_type)
        for seg_items, list_type in segments:
            start_idx = len(items)
            items.extend(seg_items)
            if list_type is not None:
                list_groups.append((start_idx, len(items) - 1, list_type))
        rng = self.insert_block(items, where=where)
        if not list_groups:
            return rng
        # Recompute each paragraph's offset exactly as insert_block walks them
        # (UTF-16 text length + one CR each, from the block's start), so a list
        # group's span can be addressed without re-querying the document.
        texts = [runs_to_text(runs) for runs, _ in normalize_block_items(items)]
        offsets: list[int] = []
        off = rng.start
        for t in texts:
            offsets.append(off)
            off += _utf16_len(t) + 1
        for first, last, list_type in list_groups:
            span = RangeAnchor(self._doc, offsets[first], offsets[last] + _utf16_len(texts[last]))
            span.apply_list(list_type)
        return rng

    def insert_image(
        self,
        image: str | Path | bytes,
        *,
        wrap: str,
        where: str = "after",
        block: bool = False,
        width: float | None = None,
        height: float | None = None,
        alt_text: str | None = None,
        lock_aspect: bool = True,
    ) -> ShapeAnchor | None:
        """Insert an image at this anchor (atomic-undo when inside `doc.edit()`).

        `image` is a file path, raw image bytes, or a base64 string — a `str`
        is treated as a path when it names an existing file, otherwise as
        base64. Word embeds the picture (`SaveWithDocument=True`) and
        auto-detects its natural size, so `width`/`height` (points) are optional
        overrides. `alt_text` sets the image's accessibility text.

        `wrap` is required — there is no default — so layout intent is always
        explicit:

        - ``"inline"`` keeps the image in the text flow (an `InlineShape`).
        - ``"auto"`` floats it: Square when its width is at most half the
          section's usable text width, else top-and-bottom.
        - ``"square" | "tight" | "through" | "top-bottom" | "front" | "behind"``
          floats it with that wrap type.

        `where` is ``"after"`` (default) or ``"before"`` the anchor's range.

        `block` places the image in its own new paragraph (reset to ``Normal``)
        rather than embedding it in the anchor's text run — so
        ``heading.insert_image(..., wrap="inline", where="before", block=True)``
        drops the image on its own line *above* the heading instead of joining
        the heading text. Without it, an inline image anchored at a heading lands
        mid-run and the heading text trails it on the same line.

        A floating image (any `wrap` other than ``"inline"``) leaves the inline
        text flow, so `image:N` no longer addresses it — this returns its floating
        [`ShapeAnchor`][wordlive.ShapeAnchor] (`shape:N`) for restyle
        (re-wrap / reposition / resize / `replace_image`). An ``"inline"`` image
        stays an `InlineShape` (addressed as `image:N`) and returns ``None``.

        Raises `ImageSourceError` for a missing/unreadable/invalid image and
        `ValueError` for an unknown `wrap` or `where`.
        """
        if wrap not in _WRAP_VALUES:
            raise ValueError(f"unknown wrap {wrap!r}; expected one of {sorted(_WRAP_VALUES)}")
        if where not in ("before", "after"):
            raise ValueError(f"where must be 'before' or 'after'; got {where!r}")
        # New paragraphs inherit the anchor's style — a block image above a
        # heading would otherwise become a heading-styled (and outline-polluting)
        # paragraph. Reset it to the body default, like insert_table does.
        normal_obj = self._doc.styles["Normal"] if block and "Normal" in self._doc.styles else None
        with _images.image_on_disk(image) as disk_path:
            with _com.translate_com_errors():
                doc_com = self._doc.com
                rng = self._range()
                pos = int(rng.Start) if where == "before" else int(rng.End)
                if block:
                    # Open a fresh paragraph at the insertion point and target it,
                    # so the image sits on its own line instead of in the run.
                    doc_com.Range(pos, pos).Text = "\r"
                    if normal_obj is not None:
                        doc_com.Range(pos, pos).Paragraphs(1).Range.Style = normal_obj.com
                insert_rng = doc_com.Range(pos, pos)
                ish = insert_rng.InlineShapes.AddPicture(
                    FileName=disk_path,
                    LinkToFile=False,
                    SaveWithDocument=True,
                    Range=insert_rng,
                )
                ish.LockAspectRatio = int(MsoTriState.TRUE if lock_aspect else MsoTriState.FALSE)
                if width is not None:
                    ish.Width = float(width)
                if height is not None:
                    ish.Height = float(height)
                if alt_text is not None:
                    ish.AlternativeText = alt_text
                if wrap == "inline":
                    return None
                wrap_type = _resolve_wrap(wrap, ish, insert_rng)
                shape = ish.ConvertToShape()
                shape.WrapFormat.Type = int(wrap_type)
                if alt_text is not None:
                    # AlternativeText doesn't always survive the conversion.
                    shape.AlternativeText = alt_text
                # The picture left InlineShapes (image:N no longer addresses it),
                # so hand back its floating shape:N handle for restyle. Locate by a
                # unique temp name — don't assume "last" (other floats can reorder).
                orig_name = str(shape.Name or "")
                probe_name = f"_wl_shape_{secrets.token_hex(8)}"
                shape.Name = probe_name
                index = _shapes.index_of_named(doc_com, probe_name)
                if orig_name:
                    shape.Name = orig_name
            return ShapeAnchor(self._doc, index)

    def insert_text_box(
        self,
        text: str,
        *,
        width: Any = 200,
        height: Any = 100,
        wrap: str = "square",
        where: str = "after",
        font: str | None = None,
        size: Any = None,
        bold: bool | None = None,
        italic: bool | None = None,
        alignment: str | None = None,
        fill: str | None = None,
        border: str | bool | None = None,
    ) -> ShapeAnchor:
        """Insert a floating text box (a pull quote / call-out) anchored here.

        A `Shapes.AddTextbox` floating shape is anchored to this anchor's
        paragraph and seeded with `text`. `width` / `height` are points or a unit
        string (``"3in"`` / ``"8cm"``). `wrap` is how body text flows around it —
        ``"square"`` (default), ``"tight"``, ``"through"``, ``"top-bottom"``,
        ``"front"``, or ``"behind"`` (the same vocabulary as `insert_image`, minus
        ``"inline"``). `where` places the anchor ``"after"`` (default) or
        ``"before"`` this anchor's range.

        The remaining kwargs style the box and its text, each optional:
        `font` / `size` (points or unit string) / `bold` / `italic` set the
        character format; `alignment` (``"left"``/``"center"``/``"right"``/
        ``"justify"``) the paragraph; `fill` is a background colour
        (``"#eeeeff"`` / ``"navy"``) and `border` is ``False`` for no outline, a
        colour string for a coloured outline, or ``True`` for the default.

        Returns the text box's floating [`ShapeAnchor`][wordlive.ShapeAnchor]
        (`shape:N`) so it can be restyled in place afterwards (`set_text` /
        `set_wrap` / `set_position` / `set_size` / `format`); discover text boxes
        later via [`doc.text_boxes`][wordlive.Document.text_boxes]. Wrap in
        `doc.edit(...)` for atomic undo; raises `ValueError` for an unknown
        `wrap` / `where`.
        """
        if wrap not in _WRAP_NAMES:
            raise ValueError(
                f"unknown wrap {wrap!r}; expected one of {sorted(_WRAP_NAMES)} (text boxes float)"
            )
        if where not in ("before", "after"):
            raise ValueError(f"where must be 'before' or 'after'; got {where!r}")
        wd_align = _ALIGNMENT_NAMES[alignment] if alignment is not None else None
        try:
            w = to_points(width)
            h = to_points(height)
            font_size = to_points(size) if size is not None else None
            fill_bgr = to_bgr(fill) if fill is not None else None
            border_bgr = to_bgr(border) if isinstance(border, str) else None
        except (ValueError, TypeError) as e:
            raise OpError(str(e)) from e
        with _com.translate_com_errors():
            doc_com = self._doc.com
            rng = self._range()
            pos = int(rng.Start) if where == "before" else int(rng.End)
            anchor_rng = doc_com.Range(pos, pos)
            shape = doc_com.Shapes.AddTextbox(
                Orientation=int(MsoTextOrientation.HORIZONTAL),
                Left=0.0,
                Top=0.0,
                Width=w,
                Height=h,
                Anchor=anchor_rng,
            )
            text_range = shape.TextFrame.TextRange
            text_range.Text = text
            font_obj = text_range.Font
            if font is not None:
                font_obj.Name = font
            if font_size is not None:
                font_obj.Size = font_size
            if bold is not None:
                font_obj.Bold = int(MsoTriState.TRUE if bold else MsoTriState.FALSE)
            if italic is not None:
                font_obj.Italic = int(MsoTriState.TRUE if italic else MsoTriState.FALSE)
            if wd_align is not None:
                text_range.ParagraphFormat.Alignment = int(wd_align)
            shape.WrapFormat.Type = int(_WRAP_NAMES[wrap])
            if fill_bgr is not None:
                shape.Fill.Visible = int(MsoTriState.TRUE)
                shape.Fill.Solid()
                shape.Fill.ForeColor.RGB = fill_bgr
            if border is False:
                shape.Line.Visible = int(MsoTriState.FALSE)
            elif border_bgr is not None:
                shape.Line.Visible = int(MsoTriState.TRUE)
                shape.Line.ForeColor.RGB = border_bgr
            # Hand back the new text box's shape:N handle (locate by a unique temp
            # name — don't assume "last", other floats can reorder).
            orig_name = str(shape.Name or "")
            probe_name = f"_wl_shape_{secrets.token_hex(8)}"
            shape.Name = probe_name
            index = _shapes.index_of_named(doc_com, probe_name)
            if orig_name:
                shape.Name = orig_name
        return ShapeAnchor(self._doc, index)

    def insert_chart(
        self,
        kind: str,
        data: Any,
        *,
        title: str | None = None,
        where: str = "after",
    ) -> ChartAnchor:
        """Insert an Excel-backed chart at this anchor and return it.

        `kind` is one of ``"bar"`` (clustered columns), ``"pie"``, ``"line"``, or
        ``"scatter"``. `data` is either an object mapping ``{label: value}`` (for
        bar / pie / line) or an array of ``[x, y]`` pairs (for ``scatter`` — both
        axes numeric, duplicate x preserved — and ``line``). `title` sets the
        chart title and series name; ``None`` leaves it untitled. `where` places
        the chart ``"after"`` (default) or ``"before"`` this anchor's range.

        Charts are Excel-backed: this embeds a chart whose data lives in a hidden
        Excel workbook, then breaks the link so the data is **static** — no live
        workbook ships in the document and the series data can't be read back
        (deferred). Requires Excel installed: raises `ExcelNotAvailableError`
        (CLI exit 6), checked up front so the document is untouched on a missing
        Excel. Raises `OpError` for malformed `data` and `ValueError` for an
        unknown `kind` / `where`.

        Word's chart API only inserts off the live `Selection`, so this moves the
        cursor to the insertion point; wrap in `doc.edit(...)` (as the CLI / exec
        / MCP surfaces do) for atomic undo and to restore the user's selection.
        """
        if kind not in _charts.KIND_TO_XL:
            raise ValueError(
                f"unknown chart kind {kind!r}; expected one of {sorted(_charts.KIND_TO_XL)}"
            )
        if where not in ("before", "after"):
            raise ValueError(f"where must be 'before' or 'after'; got {where!r}")
        xs, ys = _charts.normalize_chart_data(kind, data)
        if not _charts.probe_excel_available():
            raise ExcelNotAvailableError()
        xl_type = int(_charts.KIND_TO_XL[kind])
        with _com.translate_com_errors():
            doc_com = self._doc.com
            rng = self._range()
            pos = int(rng.Start) if where == "before" else int(rng.End)
            # AddChart2 only works off the Selection, never an arbitrary Range
            # (a Range raises "Requested object is not available"). doc.edit()
            # restores the user's selection on exit.
            doc_com.Range(pos, pos).Select()
            shape = doc_com.Application.Selection.InlineShapes.AddChart2(-1, xl_type)
            try:
                _charts.populate_chart(shape.Chart, kind, xs, ys, title)
            except Exception:
                # Don't leave a half-built placeholder chart behind on failure.
                try:
                    shape.Delete()
                except Exception:
                    pass
                raise
            index = _chart_index_at(doc_com, int(shape.Range.Start))
        return ChartAnchor(self._doc, index)

    def insert_equation(
        self,
        *,
        unicodemath: str | None = None,
        latex: str | None = None,
        mathml: str | None = None,
        where: str = "after",
        display: bool = True,
    ) -> EquationAnchor:
        """Insert a mathematical equation at this anchor and return it.

        The equation is given in exactly one of three input dialects:

        - ``unicodemath=`` — Word's native **UnicodeMath** linear form, e.g.
          ``"x=(-b±√(b^2-4ac))/(2a)"`` or ``"a^2+b^2=c^2"``. Zero-dependency: the
          string is typed into a math zone and *built up* into the 2-D form by
          Word itself.
        - ``latex=`` — a **LaTeX** math string, e.g.
          ``r"\\frac{-b\\pm\\sqrt{b^2-4ac}}{2a}"``. Converted LaTeX→MathML→OMML;
          the LaTeX→MathML hop needs the optional ``latex`` extra
          (`pip install "wordlive[latex]"`) and raises `EquationError` without it.
        - ``mathml=`` — a **MathML** (``<math>…</math>``) string. Converted
          MathML→OMML through Office's own transform (no extra needed).

        The equation always lands on its **own paragraph**, and that paragraph's
        style is pinned so it never inherits the style of whatever it was
        inserted next to (an equation dropped before a `Heading 2` used to come
        out *styled* `Heading 2` and land in the outline/TOC). `display` (default
        ``True``) gives it the dedicated centred ``Equation`` paragraph style
        (created on first use, based on ``Normal`` — a stable hook for later
        equation numbering); ``display=False`` resets the paragraph to ``Normal``
        and left-aligns it (it is still its own paragraph — wordlive does not
        place math mid-sentence — but reads as body text, not centred display
        math). `where` is ``"after"`` (default) or ``"before"`` this anchor's
        range — so ``doc.headings["Derivation"].insert_equation(...)`` drops an
        equation under a heading and ``doc.end.insert_equation(...)`` appends one.

        Returns an [`EquationAnchor`][wordlive.EquationAnchor] (`equation:N`);
        read it back as MathML with `equation.mathml`, or discover every equation
        via [`doc.equations`][wordlive.Document.equations]. Wrap in
        `doc.edit(...)` for atomic undo. Raises `EquationError` for malformed
        input (none, or more than one, of the three dialects; unparseable
        MathML/LaTeX; a missing LaTeX backend) and `ValueError` for a bad `where`.
        """
        given = [
            name
            for name, value in (
                ("unicodemath", unicodemath),
                ("latex", latex),
                ("mathml", mathml),
            )
            if value is not None
        ]
        if len(given) != 1:
            raise EquationError(
                "insert_equation needs exactly one of unicodemath=, latex=, or mathml="
                + (f"; got {', '.join(given)}" if given else "")
            )
        if where not in ("before", "after"):
            raise ValueError(f"where must be 'before' or 'after'; got {where!r}")
        if unicodemath is not None:
            return self._insert_equation_native(unicodemath, where=where, display=display)
        mathml_src = _equations.latex_to_mathml(latex) if latex is not None else (mathml or "")
        omml_inner = _equations.mathml_to_omml(mathml_src)
        return self._insert_equation_omml(omml_inner, where=where, display=display)

    def _equation_paragraph_span(self, where: str) -> tuple[int, int]:
        """Return the `(start, end)` of the document paragraph the equation attaches to.

        An equation always lands on its own paragraph, so insertion targets a
        *paragraph mark*, never a mid-paragraph offset — addressing off the
        anchor's raw range would land inside a math zone (`equation:N`) or
        mid-sentence (a bookmark). We resolve the paragraph containing the
        relevant edge of the anchor: its **start** for ``"before"``, its last real
        character (``End - 1``, clamped off the terminal mark) for ``"after"``.
        """
        rng = self._range()
        doc_com = self._doc.com
        doc_end = int(doc_com.Content.End)
        if where == "before":
            probe = max(0, int(rng.Start))
        else:
            probe = min(max(int(rng.Start), int(rng.End) - 1), max(0, doc_end - 1))
        para = doc_com.Range(probe, probe).Paragraphs(1).Range
        return int(para.Start), int(para.End)

    def _insert_equation_native(
        self, unicodemath: str, *, where: str, display: bool
    ) -> EquationAnchor:
        """Native UnicodeMath path: type the linear string, wrap it, BuildUp.

        Opens a fresh paragraph at the containing paragraph's boundary, writes the
        linear string into it, wraps the run in an `OMaths.Add` zone, and asks
        Word to build it up into the 2-D form. No XML, no extra dependency.
        """
        with _com.translate_com_errors():
            doc_com = self._doc.com
            pstart, pend = self._equation_paragraph_span(where)
            doc_end = int(doc_com.Content.End)
            if where == "before":
                # Write "<text>\r" at the paragraph start: the string becomes a new
                # paragraph and pushes the anchor's paragraph down. Clean for any
                # position, including the very start of the document (prepend).
                doc_com.Range(pstart, pstart).Text = unicodemath + "\r"
                ms = pstart
            elif pend >= doc_end:
                # The anchor's paragraph is the last; there's no position past the
                # undeletable terminal mark, so split "\r<text>" in just before it.
                pos = max(0, doc_end - 1)
                doc_com.Range(pos, pos).Text = "\r" + unicodemath
                ms = pos + 1
            else:
                # Open a new paragraph after the containing one and write into it.
                doc_com.Range(pend, pend).Text = unicodemath + "\r"
                ms = pend
            me = ms + _utf16_len(unicodemath)
            zone_rng = doc_com.Range(ms, me)
            zone_rng.OMaths.Add(zone_rng)
            zone = _equations.omath_in_range(doc_com, ms)
            if zone is not None:
                zone.BuildUp()
                zone.Type = 1 if display else 0
            index = _equation_index_at(doc_com, ms)
            self._style_equation_paragraph(ms, display=display)
        return EquationAnchor(self._doc, index)

    def _insert_equation_omml(
        self, omml_inner: str, *, where: str, display: bool
    ) -> EquationAnchor:
        """OMML path (latex/mathml): splice into a live template and InsertXML.

        `Range.InsertXML` only accepts a full, valid WordprocessingML package, so
        we take a live `Range.WordOpenXML` at a paragraph mark as the template and
        inject one math paragraph there. ``"after"`` targets the containing
        paragraph's mark; ``"before"`` targets the *preceding* paragraph's mark.
        Prepending before the first paragraph has no preceding mark to split
        against, so we open a leading paragraph first and trim the stray empty
        paragraph afterwards.
        """
        with _com.translate_com_errors():
            doc_com = self._doc.com
            doc_end = int(doc_com.Content.End)
            pstart, pend = self._equation_paragraph_span(where)
            prepend = where == "before" and pstart <= 0
            if prepend:
                doc_com.Range(0, 0).Text = "\r"
                t = 0
            elif where == "before":
                t = pstart - 1
            else:
                t = min(pend - 1, max(0, doc_end - 1))
            package = _equations.equation_package(
                str(doc_com.Range(t, t).WordOpenXML), omml_inner, display=display
            )
            doc_com.Range(t, t).InsertXML(package)
            if prepend and str(doc_com.Content.Text).startswith("\r"):
                # Trim the leading empty paragraph opened to anchor the prepend.
                doc_com.Range(0, 1).Delete()
            eq_pos = t if prepend else t + 1
            index = _equation_index_at(doc_com, eq_pos)
            self._style_equation_paragraph(eq_pos, display=display)
        return EquationAnchor(self._doc, index)

    def _ensure_equation_style(self) -> Any | None:
        """Return the COM ``Equation`` paragraph style, creating it if absent.

        A centred, ``Normal``-based paragraph style dedicated to display
        equations. Applying it to every display equation means an inserted
        equation can never inherit a heading style from its insertion point
        (which would drop the equation into the navigation outline / TOC), and
        gives a stable, named hook for future equation numbering and
        cross-references. Returns ``None`` for a degenerate document with no
        ``Normal`` to base it on — the caller then falls back to ``Normal``.
        """
        styles = self._doc.styles
        if "Equation" in styles:
            return styles["Equation"].com
        if "Normal" not in styles:
            return None
        style = styles.add("Equation", based_on="Normal", next_style="Normal")
        style.format_paragraph(alignment="center")
        return style.com

    def _style_equation_paragraph(self, pos: int, *, display: bool) -> None:
        """Pin the style/alignment of the paragraph an equation just landed on.

        Without this, an equation written at a paragraph boundary inherits the
        *following* paragraph's style — so an equation inserted before a
        ``Heading 2`` came out styled ``Heading 2`` and polluted the outline/TOC.
        A **display** equation gets the dedicated centred ``Equation`` style; an
        **inline** (``display=False``) equation is reset to ``Normal`` and
        left-aligned (it still lands on its own paragraph, but reads as body
        text, not centred display math). Best-effort — a COM hiccup here must not
        sink an otherwise-successful insert.
        """
        doc_com = self._doc.com
        try:
            para = doc_com.Range(pos, pos).Paragraphs(1).Range
            if display:
                eq_style = self._ensure_equation_style()
                if eq_style is not None:
                    para.Style = eq_style
                # Centring comes from the Equation style, so a redefined style
                # still drives it — no competing direct alignment.
            else:
                if "Normal" in self._doc.styles:
                    para.Style = self._doc.styles["Normal"].com
                para.ParagraphFormat.Alignment = int(WdParagraphAlignment.LEFT)
        except Exception:  # noqa: BLE001 — styling is a finishing touch, not the insert
            pass

    def insert_table(
        self,
        rows: int | None = None,
        cols: int | None = None,
        *,
        where: str = "after",
        style: str | None = None,
        data: list[Any] | None = None,
        header: bool = False,
    ) -> Any:
        """Create a `rows` × `cols` table at this anchor and return it.

        The structural counterpart to `insert_image` — it *creates* new
        document structure rather than editing existing structure. Returns the
        new [`Table`][wordlive.Table] wrapper so create → fill → read closes on
        one object; the table's 1-based document index is on `.index`.

        `where` is ``"after"`` (default) or ``"before"`` this anchor's range —
        so `doc.headings["Pricing"].insert_table(...)` drops a table just under
        a heading, and `doc.end.insert_table(...)` (i.e.
        [`Document.add_table`][wordlive.Document.add_table]) appends one.

        `style` names a table style defined in the document (e.g. ``"Table
        Grid"``); an unknown name raises `StyleNotFoundError` before anything is
        inserted. `style=None` applies the built-in ``"Table Grid"`` when it's
        available, so a table has visible borders by default rather than the
        invisible cell gridlines of a styleless table.

        `data` populates the cells at creation and can be given two ways:

        - a **row-major 2-D list** (``[[r1c1, r1c2], …]``); or
        - **records** — a list of dicts (``[{"Item": "Travel", "Cost": "$400"},
          …]``), where the first record's keys become a header row and each
          dict a body row (so `header` is forced on). The natural shape for
          tabular data an LLM already has as rows of objects.

        When `data` is given, `rows`/`cols` are **optional** — they're inferred
        from the data's shape — so the common case is just
        ``end.insert_table(data=…)``. Pass them explicitly to pad the grid
        larger than the data; `data` is validated against the final `rows` ×
        `cols` up front (`OpError` on overflow) and a short payload leaves the
        trailing cells empty. Filling at creation keeps the whole grid in one
        atomic undo and beats a `set_cell` storm. With no `data`, both `rows`
        and `cols` are required.

        `header=True` bolds the first row as a header (records imply it). Wrap
        in `doc.edit(...)` for atomic undo. Raises `ValueError` for an unknown
        `where` and `OpError` for a non-positive `rows`/`cols`, a missing
        dimension with no data to infer it from, or a bad `data` shape.
        """
        from ._tables import Table, index_of

        if where not in ("before", "after"):
            raise ValueError(f"where must be 'before' or 'after'; got {where!r}")
        # Normalise data first so rows/cols can be inferred from its shape.
        grid: list[list[Any]] | None = None
        if data is not None:
            grid, header_from_data = _normalize_table_data(data)
            header = header or header_from_data
            if rows is None:
                rows = len(grid)
            if cols is None:
                cols = max((len(r) for r in grid), default=0)
        if rows is None or cols is None:
            raise OpError("insert_table needs rows and cols, or a data payload to infer them from")
        if isinstance(rows, bool) or not isinstance(rows, int) or rows < 1:
            raise OpError(f"table rows must be a positive integer; got {rows!r}")
        if isinstance(cols, bool) or not isinstance(cols, int) or cols < 1:
            raise OpError(f"table cols must be a positive integer; got {cols!r}")
        if grid is not None:
            _validate_table_data(grid, rows, cols)
        # Resolve the style up-front so a bad name fails before any mutation.
        if style is not None:
            style_obj = self._doc.styles[style]  # StyleNotFoundError (exit 2) if missing
        elif "Table Grid" in self._doc.styles:
            style_obj = self._doc.styles["Table Grid"]
        else:
            style_obj = None
        # New cells inherit the *paragraph* style at the insertion point — drop a
        # table right after a `Heading 2` and Word makes every cell Heading 2,
        # which renders as large heading text and pollutes the navigation
        # outline. Reset the cells to the body default (`Normal`) so a table
        # looks like a table regardless of where it was anchored. The table
        # `style` above (borders etc.) and `header` bolding still apply on top.
        normal_obj = self._doc.styles["Normal"] if "Normal" in self._doc.styles else None
        with _com.translate_com_errors():
            doc_com = self._doc.com
            rng = self._range()
            pos = int(rng.Start) if where == "before" else int(rng.End)
            # Word's final paragraph mark is undeletable and Tables.Add needs a
            # paragraph *after* the insertion point to anchor the table; at/after
            # that mark there is none, so the add raises COM 0x80020009. Push a
            # trailing paragraph first so the table lands before it (a document
            # can't end with a table anyway — Word keeps a paragraph after one).
            doc_end = int(doc_com.Content.End)
            if pos >= doc_end - 1:
                pos = max(0, doc_end - 1)
                doc_com.Range(pos, pos).Text = "\r"
            # Word merges two tables that touch with no paragraph mark between
            # them, so a table appended at the end (or dropped next to another)
            # would silently fuse into its neighbour. Push a separator paragraph
            # onto whichever side abuts an existing table; untouched insertions
            # into ordinary text get no stray paragraph.
            if _within_table(doc_com, pos - 1, pos):
                doc_com.Range(pos, pos).Text = "\r"
                pos += 1
            if _within_table(doc_com, pos, pos + 1):
                doc_com.Range(pos, pos).Text = "\r"
            insert_rng = doc_com.Range(pos, pos)
            table_com = doc_com.Tables.Add(insert_rng, rows, cols)
            if style_obj is not None:
                table_com.Style = style_obj.com
            if normal_obj is not None:
                # Per-cell rather than table_com.Range.Style: a paragraph style
                # set on the whole table range can bleed onto the paragraph that
                # follows the table; the cell loop is contained and explicit.
                normal_com = normal_obj.com
                for r in range(1, rows + 1):
                    for c in range(1, cols + 1):
                        table_com.Cell(r, c).Range.Style = normal_com
            if grid:
                for r, row in enumerate(grid, start=1):
                    for c, val in enumerate(row, start=1):
                        table_com.Cell(r, c).Range.Text = str(val)
            if header:
                table_com.Rows(1).Range.Bold = True
            index = index_of(self._doc.com, table_com)
        return Table(self._doc, table_com, index)

    def insert_break(self, kind: str = "page", *, where: str = "after") -> None:
        """Insert a page, column, or section break at this anchor.

        The explicit one-off break — the clean alternative to appending a
        paragraph whose text is a literal form-feed. `kind` is one of:

        - ``"page"`` (default) — a manual page break (the 90% case).
        - ``"column"`` — a column break (multi-column layouts).
        - ``"section_next"`` — a section break that starts the new section on
          the next page.
        - ``"section_continuous"`` — a section break with no page break, so the
          new section flows on the same page.

        Section breaks pair with [`Document.sections`][wordlive.Document.sections]:
        each new section gets its own headers/footers and page setup. To make a
        *style* (e.g. every `Heading 1`) open a new page without a stray break
        character, prefer
        [`format_paragraph(page_break_before=True)`][wordlive.Anchor.format_paragraph]
        instead — it survives reflow.

        `where` is ``"after"`` (default) or ``"before"`` this anchor's range.
        Wrap in `doc.edit(...)` for atomic undo. Raises `ValueError` for an
        unknown `kind` or `where`.
        """
        if kind not in _BREAK_TYPES:
            raise ValueError(f"unknown break kind {kind!r}; expected one of {sorted(_BREAK_TYPES)}")
        if where not in ("before", "after"):
            raise ValueError(f"where must be 'before' or 'after'; got {where!r}")
        break_type = _BREAK_TYPES[kind]
        # A section break creates a *new* paragraph to carry the break, and that
        # paragraph inherits the anchor's style — drop one before a `Heading 1`
        # and Word makes the break paragraph a heading, leaving a spurious empty
        # entry in the navigation outline / TOC. Reset it to `Normal` so the break
        # is invisible to the outline. (Page/column breaks are an in-paragraph
        # character and create no such paragraph, so they need no reset.)
        is_section = kind in ("section_next", "section_continuous")
        normal_obj = (
            self._doc.styles["Normal"] if is_section and "Normal" in self._doc.styles else None
        )
        with _com.translate_com_errors():
            rng = self._range()
            pos = int(rng.Start) if where == "before" else int(rng.End)
            insert_rng = self._doc.com.Range(pos, pos)
            insert_rng.InsertBreak(Type=int(break_type))
            if normal_obj is not None:
                # The break now occupies the position we inserted at; the
                # paragraph containing `pos` is the break paragraph.
                break_para = self._doc.com.Range(pos, pos).Paragraphs(1)
                break_para.Range.Style = normal_obj.com

    def insert_field(self, kind: str, *, text: str | None = None, where: str = "after") -> None:
        """Insert a Word field at this anchor — a self-updating value, not literal text.

        A field shows a computed value Word keeps current: a page number, the
        page count, today's date, the file name, a document property. The named
        kinds are:

        - ``"page"`` — the current page number (`{ PAGE }`).
        - ``"numpages"`` — the total page count (`{ NUMPAGES }`); pair with
          ``"page"`` for "Page X of Y".
        - ``"date"`` / ``"time"`` — the current date / time.
        - ``"filename"`` — the document's file name.
        - ``"author"`` / ``"title"`` — document-property fields.

        For anything else, ``kind="field"`` is the escape hatch: pass the raw
        field code as `text` (e.g.
        ``insert_field("field", text="REF myBookmark \\\\h")``) and Word inserts an
        empty field carrying that code.

        Page numbers belong in a header or footer — because a `HeaderFooter`
        *is* an anchor, ``doc.sections[1].footer().insert_field("page")`` works,
        and [`HeaderFooter.insert_page_number()`][wordlive.HeaderFooter] is the
        sugar for it. Newly inserted fields render once; call
        [`Document.update_fields()`][wordlive.Document] (or take a `snapshot`,
        which repaginates) to refresh them after later edits.

        `where` is ``"after"`` (default) or ``"before"`` this anchor's range.
        Bad input raises `OpError`. Wrap in `doc.edit(...)` for atomic undo.
        """
        try:
            if where not in ("before", "after"):
                raise ValueError(f"where must be 'before' or 'after'; got {where!r}")
            wd_type = _coerce_named(kind, _FIELD_TYPES, "field kind")
            if wd_type == int(WdFieldType.EMPTY) and not text:
                raise ValueError(
                    'field kind "field" requires the raw field code via text= (e.g. text="PAGE")'
                )
            with _com.translate_com_errors():
                # Collapse a *duplicate* of the anchor's own range, so the field
                # lands in the same story — critical for header/footer anchors,
                # whose offsets are not main-document positions (a `doc.Range`
                # there would target the body instead).
                insert_rng = self._range().Duplicate
                insert_rng.Collapse(
                    int(WdCollapseDirection.START if where == "before" else WdCollapseDirection.END)
                )
                # Positional args: the Type=/Text= keywords are dropped under
                # pywin32 late binding (same gotcha as TabStops.Add / Footnotes).
                if text is not None:
                    insert_rng.Fields.Add(insert_rng, wd_type, text)
                else:
                    insert_rng.Fields.Add(insert_rng, wd_type)
        except (ValueError, TypeError) as e:
            raise OpError(str(e)) from e

    def insert_footnote(self, text: str, *, where: str = "after") -> Any:
        """Insert a footnote at this anchor and return it as a `Footnote` anchor.

        A footnote drops a reference mark in the main text and puts `text` in the
        note body at the bottom of the page; Word auto-numbers the mark. The
        returned [`Footnote`][wordlive.Footnote] is addressed `footnote:N`, so
        `note.set_text(...)` edits the body and `note.delete()` removes the mark
        and body together. Discover existing footnotes with
        [`doc.footnotes`][wordlive.Document.footnotes].

        `where` is ``"after"`` (default) or ``"before"`` this anchor's range —
        the side the reference mark lands on. Wrap in `doc.edit(...)` for atomic
        undo. Bad input raises `OpError`.
        """
        return self._insert_note("Footnotes", "footnote", text, where=where)

    def insert_endnote(self, text: str, *, where: str = "after") -> Any:
        """Insert an endnote at this anchor and return it as an `Endnote` anchor.

        The endnote mirror of [`insert_footnote`][wordlive.Anchor.insert_footnote]:
        the reference mark lands in the main text and `text` collects at the end
        of the document (or section). The returned
        [`Endnote`][wordlive.Endnote] is addressed `endnote:N`; discover existing
        endnotes with [`doc.endnotes`][wordlive.Document.endnotes].

        `where` is ``"after"`` (default) or ``"before"`` this anchor's range.
        Wrap in `doc.edit(...)` for atomic undo. Bad input raises `OpError`.
        """
        return self._insert_note("Endnotes", "endnote", text, where=where)

    def _insert_note(self, attr: str, scheme: str, text: str, *, where: str) -> Any:
        """Shared footnote/endnote insertion (`attr` is the COM collection name)."""
        from ._notes import Endnote, Footnote, index_of_note

        cls = Footnote if scheme == "footnote" else Endnote
        try:
            if where not in ("before", "after"):
                raise ValueError(f"where must be 'before' or 'after'; got {where!r}")
            with _com.translate_com_errors():
                rng = self._range()
                # A note's reference mark always lands in the main text story, so
                # a plain document Range at the anchor's edge is correct (unlike
                # insert_field, which can target a footer's own story).
                pos = int(rng.Start) if where == "before" else int(rng.End)
                insert_rng = self._doc.com.Range(pos, pos)
                coll = getattr(self._doc.com, attr)
                # Positional args: an empty Reference auto-numbers, and the
                # Reference=/Text= keywords are dropped under pywin32 late binding
                # (same gotcha as Fields.Add / TabStops.Add).
                coll.Add(insert_rng, "", text)
                index = index_of_note(coll, pos)
            return cls(self._doc, index)
        except (ValueError, TypeError) as e:
            raise OpError(str(e)) from e

    def insert_toc(
        self,
        *,
        levels: tuple[int, int] = (1, 3),
        use_heading_styles: bool = True,
        hyperlinks: bool = True,
        where: str = "after",
    ) -> Any:
        """Insert a table of contents at this anchor and return it as a `Toc`.

        Builds a TOC from the document's heading paragraphs over the given
        `levels` (a ``(upper, lower)`` pair — `(1, 3)` covers Heading 1–3).
        `use_heading_styles=True` sources entries from the built-in Heading
        styles; `hyperlinks=True` makes each entry a clickable jump (and a real
        hyperlink in exported PDFs). Returns a [`Toc`][wordlive.Toc].

        A TOC's page numbers populate only after repagination — call
        `toc.update()` (or [`Document.update_fields`][wordlive.Document.update_fields],
        or take a `snapshot`, which forces print layout) before reading them.
        Most documents want the TOC at the top: `doc.add_toc(...)` is the sugar
        for `doc.start.insert_toc(...)`.

        `where` is ``"after"`` (default) or ``"before"`` this anchor's range.
        Wrap in `doc.edit(...)` for atomic undo. Bad input raises `OpError`.
        """
        from ._toc import Toc

        try:
            if where not in ("before", "after"):
                raise ValueError(f"where must be 'before' or 'after'; got {where!r}")
            try:
                upper, lower = int(levels[0]), int(levels[1])
            except (TypeError, IndexError, ValueError, KeyError) as e:
                raise ValueError(
                    f"levels must be a (upper, lower) pair of ints; got {levels!r}"
                ) from e
            if not (1 <= upper <= lower <= 9):
                raise ValueError(
                    f"levels must satisfy 1 <= upper <= lower <= 9; got {(upper, lower)}"
                )
            with _com.translate_com_errors():
                rng = self._range()
                pos = int(rng.Start) if where == "before" else int(rng.End)
                insert_rng = self._doc.com.Range(pos, pos)
                # Positional args (keyword names are dropped under pywin32 late
                # binding). Order: Range, UseHeadingStyles, UpperHeadingLevel,
                # LowerHeadingLevel, UseFields, TableID, RightAlignPageNumbers,
                # IncludePageNumbers, AddedStyles, UseHyperlinks,
                # HidePageNumbersInWeb, UseOutlineLevels.
                toc_com = self._doc.com.TablesOfContents.Add(
                    insert_rng,
                    bool(use_heading_styles),
                    upper,
                    lower,
                    False,
                    "",
                    True,
                    True,
                    "",
                    bool(hyperlinks),
                    True,
                    True,
                )
            return Toc(self._doc, toc_com)
        except (ValueError, TypeError) as e:
            raise OpError(str(e)) from e

    def link_to(
        self,
        address: str | None = None,
        *,
        bookmark: str | None = None,
        text: str | None = None,
        screen_tip: str | None = None,
    ) -> None:
        """Turn this anchor into a hyperlink (or insert new linked text).

        Pass exactly one destination: `address` for an external link (a URL,
        `mailto:`, or file path) or `bookmark` for an internal jump to a named
        bookmark in this document. With `text=None` the anchor's existing range
        becomes the clickable link; pass `text=...` to **insert** new linked text
        at the end of the anchor's range (so linking a heading or a `range:`
        phrase with `text=...` adds the link rather than overwriting the content).
        `screen_tip` is the hover tooltip.

        Pair it with [`doc.bookmarks.add(...)`][wordlive.BookmarkCollection.add]
        to build internal navigation, or a `range:START-END` id (from `find`) to
        link an existing phrase. Wrap in `doc.edit(...)` for atomic undo. Bad
        input (not exactly one destination) raises `OpError`.
        """
        try:
            if (address is None) == (bookmark is None):
                raise ValueError("link_to requires exactly one of 'address' or 'bookmark'")
            with _com.translate_com_errors():
                rng = self._range()
                if text is not None:
                    # Insert *new* linked text rather than overwriting the
                    # anchor's range: collapse to its end so a heading / phrase
                    # keeps its content and the link is added after it.
                    rng = rng.Duplicate
                    rng.Collapse(int(WdCollapseDirection.END))
                addr_arg = address or ""
                sub_arg = bookmark or ""
                tip_arg = screen_tip or ""
                # Positional args (Anchor, Address, SubAddress, ScreenTip,
                # TextToDisplay) — keep keywords out for late-binding safety.
                if text is not None:
                    self._doc.com.Hyperlinks.Add(rng, addr_arg, sub_arg, tip_arg, text)
                else:
                    self._doc.com.Hyperlinks.Add(rng, addr_arg, sub_arg, tip_arg)
        except (ValueError, TypeError) as e:
            raise OpError(str(e)) from e

    def insert_cross_reference(
        self,
        target: str,
        *,
        kind: str = "text",
        hyperlink: bool = True,
        where: str = "after",
    ) -> None:
        """Insert a cross-reference to another anchor at this anchor.

        `target` is the anchor id to point at: `bookmark:NAME`, `heading:N`,
        `footnote:N`, or `endnote:N`. `kind` selects what the reference shows:
        ``"text"`` (the heading/bookmark text — the default), ``"page"`` ("see
        page 5"), ``"number"`` (the paragraph or note number), or
        ``"above_below"`` ("above"/"below"). `hyperlink=True` makes the inserted
        reference a clickable jump.

        An unresolvable `target` raises `AnchorNotFoundError` (exit 2) before
        anything is inserted. `where` is ``"after"`` (default) or ``"before"``
        this anchor's range. Wrap in `doc.edit(...)` for atomic undo; other bad
        input raises `OpError`.
        """
        try:
            if where not in ("before", "after"):
                raise ValueError(f"where must be 'before' or 'after'; got {where!r}")
            # Resolve outside translate_com_errors so an AnchorNotFoundError for a
            # bad target propagates as exit 2 rather than being masked.
            ref_type, ref_item = _resolve_cross_ref_target(self._doc, target)
            ref_kind = _cross_ref_kind(kind, ref_type)
            with _com.translate_com_errors():
                insert_rng = self._range().Duplicate
                insert_rng.Collapse(
                    int(WdCollapseDirection.START if where == "before" else WdCollapseDirection.END)
                )
                # Positional args: IncludePositionInformation as a keyword raises
                # under pywin32 late binding, so pass only the first four.
                insert_rng.InsertCrossReference(ref_type, ref_kind, ref_item, bool(hyperlink))
        except (ValueError, TypeError) as e:
            raise OpError(str(e)) from e

    def insert_caption(
        self, label: str = "Figure", *, text: str | None = None, position: str | None = None
    ) -> None:
        """Insert a numbered caption as its **own paragraph** at this anchor.

        `label` is a caption label — built-in ``"Figure"`` / ``"Table"`` /
        ``"Equation"`` or any custom string; Word auto-numbers per label
        (Figure 1, Figure 2, …). `text` is the caption title shown after the
        label and number. Pairs with
        [`insert_cross_reference`][wordlive.Anchor.insert_cross_reference] for
        "see Figure 2".

        `position` is ``"above"`` or ``"below"`` the anchor. Left as `None` it
        follows convention: a ``"Table"`` caption goes **above**, every other
        label goes **below**. The caption always becomes its own
        `Caption`-styled paragraph — it never fuses into the target paragraph.
        On a table cell (`table:N:R:C`) the caption is placed above / below the
        **whole table**, not inside the cell.

        Wrap in `doc.edit(...)` for atomic undo. Bad input raises `OpError`.
        """
        try:
            above = _caption_above(label, position)
            title = text if text is not None else ""
            pos = int(WdCaptionPosition.ABOVE if above else WdCaptionPosition.BELOW)
            with _com.translate_com_errors():
                obj_rng = self._caption_object_range()
                if obj_rng is not None:
                    # A caption-able object (e.g. a table): let Word place the
                    # caption on its own line above/below the object natively.
                    obj_rng.InsertCaption(str(label), title, pos, False)
                else:
                    # Text/paragraph anchor: carve out a dedicated empty
                    # paragraph (before or after the anchor) and drop the
                    # caption into it, so it never fuses into the host paragraph.
                    insert_rng = self._range().Duplicate
                    insert_rng.Collapse(
                        int(WdCollapseDirection.START if above else WdCollapseDirection.END)
                    )
                    insert_rng.InsertParagraphBefore()
                    insert_rng.Collapse(int(WdCollapseDirection.START))
                    # Positional args (Label, Title, Position, ExcludeLabel) for
                    # late-binding safety; a string Label matches a built-in or
                    # defines a custom one.
                    insert_rng.InsertCaption(str(label), title, pos, False)
        except (ValueError, TypeError) as e:
            raise OpError(str(e)) from e

    def insert_content_control(
        self,
        kind: str = "rich_text",
        *,
        title: str | None = None,
        tag: str | None = None,
        items: list[Any] | None = None,
        where: str = "wrap",
        lock_contents: bool = False,
        lock_control: bool = False,
    ) -> ContentControl:
        """Wrap (or insert) a content control at this anchor and return it.

        Content controls are the building blocks of form-like documents — a
        labelled region the user fills in. `kind` selects the type:
        ``"rich_text"`` (the default — formatted text), ``"text"`` (plain text),
        ``"picture"``, ``"combo_box"`` / ``"dropdown"`` (a pick list — pass
        `items`), ``"date"``, ``"checkbox"`` (Word 2013+), ``"building_block"``,
        ``"group"``, or ``"repeating_section"`` (Word 2013+).

        `where` is ``"wrap"`` (default — the control surrounds this anchor's
        existing range, so a `range:START-END` from `find` wraps that phrase) or
        ``"before"`` / ``"after"`` (insert a fresh empty control at the anchor's
        start / end). Set `title` and/or `tag` to give the control a name: a
        titled control is addressable later as `cc:TITLE` (falling back to the
        tag) and shows a label in Word's UI. `items` populates a combo box or
        dropdown — each is a string, or a `{"text": ..., "value": ...}` dict.
        `lock_contents` stops the user editing the value; `lock_control` stops
        them deleting the control.

        Returns the new [`ContentControl`][wordlive.ContentControl] (usable even
        when unnamed — it caches the live control). Wrap in `doc.edit(...)` for
        atomic undo. Bad input raises `OpError`.
        """
        try:
            if where not in ("wrap", "before", "after"):
                raise ValueError(f"where must be 'wrap', 'before', or 'after'; got {where!r}")
            try:
                cc_type = _CC_TYPE_NAMES[str(kind).lower()]
            except (KeyError, AttributeError) as e:
                raise ValueError(
                    f"unknown content control kind {kind!r}; "
                    f"expected one of {sorted(_CC_TYPE_NAMES)}"
                ) from e
            if items is not None and cc_type not in _CC_LIST_TYPES:
                raise ValueError("items is only valid for a 'combo_box' or 'dropdown' control")
            with _com.translate_com_errors():
                target = self._range().Duplicate
                if where != "wrap":
                    target.Collapse(
                        int(
                            WdCollapseDirection.START
                            if where == "before"
                            else WdCollapseDirection.END
                        )
                    )
                # Positional args (Type, Range) for late-binding safety.
                cc = self._doc.com.ContentControls.Add(int(cc_type), target)
                if title is not None:
                    cc.Title = str(title)
                if tag is not None:
                    cc.Tag = str(tag)
                if lock_contents:
                    cc.LockContents = True
                if lock_control:
                    cc.LockContentControl = True
                if items:
                    entries = cc.DropdownListEntries
                    for entry_text, value in _normalize_cc_items(items):
                        entries.Add(entry_text, value)  # positional (Text, Value)
            name = title or tag or ""
            return ContentControl(self._doc, name, com=cc)
        except (ValueError, TypeError) as e:
            raise OpError(str(e)) from e

    def mark_index_entry(
        self,
        entry: str,
        *,
        cross_reference: str | None = None,
        bold: bool = False,
        italic: bool = False,
    ) -> None:
        """Mark this anchor's range as a back-of-book index entry (an `XE` field).

        `entry` is the text that appears in the index; use ``"main:sub"`` to file
        it as a subentry under ``main`` (Word's colon convention). `bold` /
        `italic` style the entry's *page number* in the built index.
        `cross_reference` replaces the page number with a "see …" pointer (e.g.
        ``cross_reference="Widgets"`` → *"see Widgets"*).

        This is the per-term half of indexing; once entries are marked, build the
        list with [`insert_index`][wordlive.Anchor.insert_index] /
        [`Document.add_index`][wordlive.Document.add_index]. The `XE` field is
        hidden text and doesn't disturb the visible flow. Wrap in `doc.edit(...)`
        for atomic undo. Bad input raises `OpError`.
        """
        try:
            if not str(entry).strip():
                raise ValueError("entry must be a non-empty string")
            with _com.translate_com_errors():
                rng = self._range()
                # Indexes.MarkEntry(Range, Entry, EntryAutoText, CrossReference,
                # CrossReferenceAutoText, BookmarkName, Bold, Italic) — positional
                # for late-binding safety.
                self._doc.com.Indexes.MarkEntry(
                    rng,
                    str(entry),
                    "",
                    str(cross_reference) if cross_reference is not None else "",
                    "",
                    "",
                    bool(bold),
                    bool(italic),
                )
        except (ValueError, TypeError) as e:
            raise OpError(str(e)) from e

    def insert_index(
        self,
        *,
        columns: int = 2,
        run_in: bool = False,
        right_align_page_numbers: bool = False,
        where: str = "after",
    ) -> Any:
        """Insert a back-of-book index at this anchor and return it as an `Index`.

        Gathers every `XE` entry marked with
        [`mark_index_entry`][wordlive.Anchor.mark_index_entry] into an
        alphabetised, page-numbered list. `columns` is the number of newspaper
        columns the index is laid out in (2 is the book default). `run_in=True`
        packs subentries into a single paragraph instead of one per line;
        `right_align_page_numbers=True` flushes page numbers to the right margin.

        Returns an [`Index`][wordlive.Index]; like a TOC it's a field block whose
        page numbers populate only after repagination — call `index.update()`,
        [`Document.update_fields`][wordlive.Document.update_fields], or take a
        `snapshot`. Most documents want the index at the end:
        `doc.add_index(...)` is the sugar for `doc.end.insert_index(...)`.

        `where` is ``"after"`` (default) or ``"before"`` this anchor's range.
        Wrap in `doc.edit(...)` for atomic undo. Bad input raises `OpError`.
        """
        from ._index import Index

        try:
            if where not in ("before", "after"):
                raise ValueError(f"where must be 'before' or 'after'; got {where!r}")
            cols = int(columns)
            if cols < 1:
                raise ValueError(f"columns must be >= 1; got {columns!r}")
            idx_type = int(WdIndexType.RUNIN if run_in else WdIndexType.INDENT)
            with _com.translate_com_errors():
                rng = self._range()
                pos = int(rng.Start) if where == "before" else int(rng.End)
                insert_rng = self._doc.com.Range(pos, pos)
                # Indexes.Add(Range, HeadingSeparator, RightAlignPageNumbers, Type,
                # NumberOfColumns, AccentedLetters, SortBy, IndexLanguage). Positional;
                # HeadingSeparator must be a WdHeadingSeparator value (0 = none) — an
                # empty string raises a type-mismatch on a makepy-typed Word wrapper.
                idx_com = self._doc.com.Indexes.Add(
                    insert_rng, 0, bool(right_align_page_numbers), idx_type, cols
                )
            return Index(self._doc, idx_com)
        except (ValueError, TypeError) as e:
            raise OpError(str(e)) from e

    def insert_table_of_figures(
        self,
        *,
        label: str = "Figure",
        include_label: bool = True,
        hyperlinks: bool = True,
        right_align_page_numbers: bool = True,
        where: str = "after",
    ) -> Any:
        """Insert a table of figures at this anchor and return it.

        The caption-driven sibling of [`insert_toc`][wordlive.Anchor.insert_toc]:
        it lists every caption of one `label` — ``"Figure"`` (the default),
        ``"Table"``, ``"Equation"``, or any custom label you passed to
        [`insert_caption`][wordlive.Anchor.insert_caption] — with its page number.
        `include_label=True` keeps the "Figure 1" prefix in each entry;
        `hyperlinks=True` makes entries clickable jumps;
        `right_align_page_numbers=True` flushes page numbers right.

        Returns a [`TableOfFigures`][wordlive.TableOfFigures]; like a TOC its page
        numbers populate only after repagination — call `tof.update()`,
        [`Document.update_fields`][wordlive.Document.update_fields], or take a
        `snapshot`. `where` is ``"after"`` (default) or ``"before"`` this anchor's
        range. Wrap in `doc.edit(...)` for atomic undo. Bad input raises `OpError`.
        """
        from ._toc import TableOfFigures

        try:
            if where not in ("before", "after"):
                raise ValueError(f"where must be 'before' or 'after'; got {where!r}")
            with _com.translate_com_errors():
                rng = self._range()
                pos = int(rng.Start) if where == "before" else int(rng.End)
                insert_rng = self._doc.com.Range(pos, pos)
                # Keyword args here: the optional string Variants (TableID,
                # Caption2, AddedStyles) raise a type-mismatch when passed
                # positionally as "" on a makepy-typed Word wrapper, so we name
                # only the flags we set and let the rest default (Range + Caption
                # stay positional, matching the Word signature).
                tof_com = self._doc.com.TablesOfFigures.Add(
                    insert_rng,
                    str(label),
                    IncludeLabel=bool(include_label),
                    UseHeadingStyles=False,
                    RightAlignPageNumbers=bool(right_align_page_numbers),
                    UseHyperlinks=bool(hyperlinks),
                )
            return TableOfFigures(self._doc, tof_com)
        except (ValueError, TypeError) as e:
            raise OpError(str(e)) from e

    def insert_citation(
        self,
        tag: str,
        *,
        pages: str | None = None,
        prefix: str | None = None,
        suffix: str | None = None,
        volume: str | None = None,
        suppress_author: bool = False,
        suppress_year: bool = False,
        suppress_title: bool = False,
        locale: int = 1033,
        where: str = "after",
    ) -> Any:
        """Insert an in-text citation at this anchor and return it as a `Citation`.

        References a source in the document's store (add one with
        `doc.sources.add`) by its `tag` and
        renders it in the current [`bibliography_style`][wordlive.Document.bibliography_style]
        — e.g. *(Smith 2020)*. `pages` adds a page locator (*(Smith 2020, 15)*);
        `prefix` / `suffix` wrap the citation (*"see "* / *", at 12"*); `volume`
        adds a volume. `suppress_author` / `suppress_year` / `suppress_title` drop
        those parts. `locale` is the LCID the style formats under (1033 = en-US).

        Returns a [`Citation`][wordlive.Citation]; a citation to an unknown tag
        renders *"Invalid source specified."* rather than failing. `where` is
        ``"after"`` (default) or ``"before"`` this anchor's range. Wrap in
        `doc.edit(...)` for atomic undo. Bad input raises `OpError`.
        """
        from ._citations import Citation

        try:
            if where not in ("before", "after"):
                raise ValueError(f"where must be 'before' or 'after'; got {where!r}")
            if not str(tag).strip():
                raise ValueError("tag must be a non-empty string")
            # CITATION field code (switches confirmed against live Word): \p page,
            # \v volume, \f prefix, \s suffix, \n/\y/\t suppress author/year/title.
            parts = [f"CITATION {tag} \\l {int(locale)}"]
            if pages:
                parts.append(f'\\p "{pages}"')
            if volume:
                parts.append(f"\\v {volume}")
            if prefix:
                parts.append(f'\\f "{prefix}"')
            if suffix:
                parts.append(f'\\s "{suffix}"')
            if suppress_author:
                parts.append("\\n")
            if suppress_year:
                parts.append("\\y")
            if suppress_title:
                parts.append("\\t")
            code = " ".join(parts)
            with _com.translate_com_errors():
                # Collapse a *duplicate* of the anchor's own range so the field
                # lands in the same story (header/footer-safe — see insert_field).
                insert_rng = self._range().Duplicate
                insert_rng.Collapse(
                    int(WdCollapseDirection.START if where == "before" else WdCollapseDirection.END)
                )
                # EMPTY (-1) raw-code insert — positional, the proven path Word
                # parses into a typed CITATION field (its numeric, 96, is fragile
                # to pass directly).
                field = insert_rng.Fields.Add(insert_rng, int(WdFieldType.EMPTY), code, False)
            return Citation(self._doc, field)
        except (ValueError, TypeError) as e:
            raise OpError(str(e)) from e

    def insert_bibliography(self, *, where: str = "after") -> Any:
        """Insert a bibliography at this anchor and return it as a `Bibliography`.

        Inserts a ``BIBLIOGRAPHY`` field — the reference list of every source
        *cited* in the document, formatted in the current
        [`bibliography_style`][wordlive.Document.bibliography_style]. Most
        documents want it at the end: `doc.add_bibliography()` is the sugar for
        `doc.end.insert_bibliography()`.

        Returns a [`Bibliography`][wordlive.Bibliography]; like a TOC it's a field
        block — call `bibliography.update()`,
        [`Document.update_fields`][wordlive.Document.update_fields], or take a
        `snapshot` after adding citations. `where` is ``"after"`` (default) or
        ``"before"`` this anchor's range. Wrap in `doc.edit(...)` for atomic undo.
        Bad input raises `OpError`.
        """
        from ._citations import Bibliography

        try:
            if where not in ("before", "after"):
                raise ValueError(f"where must be 'before' or 'after'; got {where!r}")
            with _com.translate_com_errors():
                insert_rng = self._range().Duplicate
                insert_rng.Collapse(
                    int(WdCollapseDirection.START if where == "before" else WdCollapseDirection.END)
                )
                field = insert_rng.Fields.Add(
                    insert_rng, int(WdFieldType.EMPTY), "BIBLIOGRAPHY", False
                )
            return Bibliography(self._doc, field)
        except (ValueError, TypeError) as e:
            raise OpError(str(e)) from e

    def mark_citation(
        self,
        long_citation: str,
        *,
        short_citation: str | None = None,
        category: str | int = "cases",
        where: str = "after",
    ) -> None:
        """Mark this anchor's range as a table-of-authorities citation (a `TA` field).

        The legal analog of [`mark_index_entry`][wordlive.Anchor.mark_index_entry]:
        `long_citation` is the full citation as it appears in the table (e.g.
        *"Smith v. Jones, 1 U.S. 1 (2020)"*), `short_citation` the abbreviated
        form Word matches elsewhere in the text (defaults to `long_citation`), and
        `category` the section it files under — ``"cases"`` (the default),
        ``"statutes"``, ``"other"``, ``"rules"``, ``"treatises"``,
        ``"regulations"``, ``"constitutional"``, or a category number (1-16).

        This is the per-authority half; build the table with
        [`insert_table_of_authorities`][wordlive.Anchor.insert_table_of_authorities]
        / [`Document.add_table_of_authorities`][wordlive.Document.add_table_of_authorities].
        The `TA` field is hidden and doesn't disturb the visible flow. Wrap in
        `doc.edit(...)` for atomic undo. Bad input raises `OpError`.
        """
        from ._toa import _TOA_CATEGORIES

        try:
            if where not in ("before", "after"):
                raise ValueError(f"where must be 'before' or 'after'; got {where!r}")
            if not str(long_citation).strip():
                raise ValueError("long_citation must be a non-empty string")
            cat = _coerce_named(category, _TOA_CATEGORIES, "TOA category")
            short = short_citation if short_citation is not None else long_citation
            code = f'TA \\l "{long_citation}" \\s "{short}" \\c {cat}'
            with _com.translate_com_errors():
                insert_rng = self._range().Duplicate
                insert_rng.Collapse(
                    int(WdCollapseDirection.START if where == "before" else WdCollapseDirection.END)
                )
                insert_rng.Fields.Add(insert_rng, int(WdFieldType.EMPTY), code, False)
        except (ValueError, TypeError) as e:
            raise OpError(str(e)) from e

    def insert_table_of_authorities(
        self,
        *,
        category: str | int = "all",
        passim: bool = True,
        keep_entry_formatting: bool = True,
        entry_separator: str | None = None,
        page_range_separator: str | None = None,
        where: str = "after",
    ) -> Any:
        """Insert a table of authorities at this anchor and return it.

        Gathers the `TA` citations marked with
        [`mark_citation`][wordlive.Anchor.mark_citation] into a page-numbered
        table. `category` selects which authorities to include — ``"all"`` (the
        default), ``"cases"``, ``"statutes"``, … or a number (1-16). `passim=True`
        replaces five-or-more page references for one authority with *"passim"*;
        `keep_entry_formatting=True` preserves each citation's character
        formatting. `entry_separator` / `page_range_separator` override the
        defaults between a citation and its first page / between page ranges.

        Returns a [`TableOfAuthorities`][wordlive.TableOfAuthorities]; like a TOC
        it's a field block whose page numbers populate only after repagination —
        call `toa.update()`, [`Document.update_fields`][wordlive.Document.update_fields],
        or take a `snapshot`. `doc.add_table_of_authorities(...)` is the sugar for
        `doc.end.insert_table_of_authorities(...)`. `where` is ``"after"``
        (default) or ``"before"`` this anchor's range. Wrap in `doc.edit(...)` for
        atomic undo. Bad input raises `OpError`.
        """
        from ._toa import _TOA_CATEGORIES, TableOfAuthorities

        try:
            if where not in ("before", "after"):
                raise ValueError(f"where must be 'before' or 'after'; got {where!r}")
            cat = _coerce_named(category, _TOA_CATEGORIES, "TOA category")
            with _com.translate_com_errors():
                rng = self._range()
                pos = int(rng.Start) if where == "before" else int(rng.End)
                insert_rng = self._doc.com.Range(pos, pos)
                # TablesOfAuthorities.Add(Range, Category, ...): Range + int
                # Category positional; the string-Variant optionals need keyword
                # form (same gotcha as TablesOfFigures).
                kwargs: dict[str, Any] = {
                    "Passim": bool(passim),
                    "KeepEntryFormatting": bool(keep_entry_formatting),
                }
                if entry_separator is not None:
                    kwargs["EntrySeparator"] = str(entry_separator)
                if page_range_separator is not None:
                    kwargs["PageRangeSeparator"] = str(page_range_separator)
                toa_com = self._doc.com.TablesOfAuthorities.Add(insert_rng, cat, **kwargs)
            return TableOfAuthorities(self._doc, toa_com)
        except (ValueError, TypeError) as e:
            raise OpError(str(e)) from e

    def snapshot(
        self, out: str | Path | None = None, *, dpi: int = 150, max_dim: int | None = None
    ) -> list[Snapshot]:
        """Render the page(s) this anchor sits on to PNG — let a model *see* it.

        A heading expands to its whole section; any other anchor renders the
        page(s) its range spans. Returns a list of
        [`Snapshot`][wordlive.Snapshot] (one per page); pass `out` to also write
        the image(s) to disk. `max_dim` caps each page's long edge in pixels (for
        a cheaper render). Sugar for
        [`Document.snapshot_anchor`][wordlive.Document.snapshot_anchor]; see it
        for the full semantics. Requires the `snapshot` extra (PyMuPDF).
        """
        return self._doc.snapshot_anchor(self, out, dpi=dpi, max_dim=max_dim)

    def read_image(self) -> tuple[bytes, str]:
        """Extract the image embedded in this anchor's range as `(bytes, mime_type)`.

        The read side of the image story — pull an embedded picture's original
        bytes back out (e.g. to hand to a vision model), the counterpart to
        [`insert_image`][wordlive.Anchor.insert_image]. The range must contain
        exactly one picture: an [`image:N`][wordlive.ImageAnchor] anchor (or any
        single-image text anchor) reads cleanly, while a range with no image — or
        more than one — raises `ImageSourceError`. `bytes` is the picture's raw
        encoded data (PNG/JPEG/…); `mime_type` is its content type
        (``"image/png"``, ``"image/jpeg"``, …). Discover what's there first with
        [`doc.images`][wordlive.Document.images]. Read-only — nothing is mutated.
        """
        with _com.translate_com_errors():
            return _images.read_image_from_range(self._range())

    def delete(self) -> None:
        with _com.translate_com_errors():
            self._range().Delete()

    def apply_style(self, name: str) -> None:
        """Apply the named paragraph or character style to this anchor's range.

        Word selects paragraph- vs. character-style behaviour from the style's
        own `Type`; we don't model that distinction. Raises `StyleNotFoundError`
        if the style isn't defined in the document.
        """
        style = self._doc.styles[name]  # raises StyleNotFoundError if missing
        with _com.translate_com_errors():
            self._range().Style = style.com

    def format_paragraph(
        self,
        *,
        alignment: Any = None,
        left_indent: float | None = None,
        right_indent: float | None = None,
        first_line_indent: float | None = None,
        space_before: float | None = None,
        space_after: float | None = None,
        line_spacing: Any = None,
        page_break_before: bool | None = None,
        keep_together: bool | None = None,
        keep_with_next: bool | None = None,
        widow_control: bool | None = None,
    ) -> None:
        """Set paragraph-formatting properties on this anchor's range.

        All kwargs are optional; only the ones explicitly passed are written.
        Indent and spacing values are in points (Word's native unit for
        `ParagraphFormat.LeftIndent` etc.). `alignment` accepts a
        `WdParagraphAlignment` enum, its int value, or a string
        (`"left"`/`"center"`/`"right"`/`"justify"`).

        `line_spacing` sets the leading between lines *within* the paragraph
        (distinct from `space_before`/`space_after`, which space paragraphs
        apart). It accepts a **number** — a multiple of single spacing (`1`
        single, `1.5`, `2` double) — one of the keywords `"single"`/`"1.5"`/
        `"double"`, or an **exact length string** (`"14pt"`, `"1.5cm"`) for a
        fixed line height.

        `page_break_before=True` forces the paragraph to begin on a new page —
        the *clean* way to page-break (e.g. apply it to every `Heading 1`): it's
        a paragraph property that survives reflow and leaves no stray break
        character, unlike [`insert_break`][wordlive.Anchor.insert_break].
        `False` clears the property. Indents/spacing accept a number (points) or
        a unit string (`"0.5in"`).

        The remaining flags are Word's *pagination* controls (all tri-state —
        `True`/`False` set, `None` leaves untouched), for clean multi-page
        layout: `keep_together` keeps every line of the paragraph on one page;
        `keep_with_next` keeps it on the same page as the following paragraph
        (e.g. a heading with its first body line); `widow_control` prevents a
        lone first/last line stranded at the bottom/top of a page (on by default
        in Word).
        """
        try:
            with _com.translate_com_errors():
                _apply_paragraph_format(
                    self._range().ParagraphFormat,
                    alignment=alignment,
                    left_indent=left_indent,
                    right_indent=right_indent,
                    first_line_indent=first_line_indent,
                    space_before=space_before,
                    space_after=space_after,
                    line_spacing=line_spacing,
                    page_break_before=page_break_before,
                    keep_together=keep_together,
                    keep_with_next=keep_with_next,
                    widow_control=widow_control,
                )
        except (ValueError, TypeError) as e:
            raise OpError(str(e)) from e

    def drop_cap(
        self,
        lines: int = 3,
        *,
        position: str = "dropped",
        distance: Any = 0.0,
        font: str | None = None,
    ) -> None:
        """Turn the first letter of this anchor's paragraph into a drop cap.

        The editorial oversized initial — a real Word `DropCap`, not a faked
        big-font run, so it reflows and re-wraps the body text around it
        natively. Applies to the **first paragraph** of the anchor's range.

        `position` is ``"dropped"`` (the default — the letter sits *into* the
        text, the common magazine style), ``"margin"`` (it hangs out in the left
        margin), or ``"none"`` (remove an existing drop cap; `lines`/`distance`/
        `font` are then ignored). `lines` is how many lines tall the letter is
        (Word's default is 3). `distance` is the gap between the letter and the
        body text, in points (or a unit string like ``"2pt"``). `font` optionally
        sets the dropped letter's font family.

        Word rejects a drop cap on an **empty** paragraph (there's no letter to
        drop) — that surfaces as a `ComError`. Wrap in `doc.edit(...)` for atomic
        undo. Raises `OpError` for an unknown `position` or a bad `distance`.
        """
        try:
            pos = _coerce_named(position, _DROP_POSITIONS, "drop-cap position")
            dist = to_points(distance)
            if not isinstance(lines, int) or isinstance(lines, bool) or lines < 1:
                raise ValueError(f"lines must be a positive integer; got {lines!r}")
            with _com.translate_com_errors():
                dc = self._range().Paragraphs(1).DropCap
                # Enable the cap first: Word resets LinesToDrop/DistanceFromText/
                # FontName to its defaults when Position changes, so the geometry
                # must be written *after* the position or it's silently dropped.
                dc.Position = pos
                if pos == int(WdDropPosition.NONE):
                    return
                dc.LinesToDrop = lines
                dc.DistanceFromText = dist
                if font is not None:
                    dc.FontName = str(font)
        except (ValueError, TypeError) as e:
            raise OpError(str(e)) from e

    def format_run(
        self,
        *,
        bold: bool | None = None,
        italic: bool | None = None,
        underline: bool | None = None,
        strikethrough: bool | None = None,
        font: str | None = None,
        size: Any = None,
        color: Any = None,
        highlight: Any = None,
        subscript: bool | None = None,
        superscript: bool | None = None,
        small_caps: bool | None = None,
        all_caps: bool | None = None,
        spacing: Any = None,
    ) -> None:
        """Set character-formatting (run-level) properties on this anchor's range.

        Direct formatting — the *bold this phrase* layer, distinct from
        [`apply_style`][wordlive.Anchor.apply_style] (named styles) and
        [`format_paragraph`][wordlive.Anchor.format_paragraph] (paragraph-scope).
        Pairs naturally with `range:START-END` to style a sub-paragraph span.

        All kwargs are optional and tri-state; only the ones explicitly passed
        are written (`None` leaves the property untouched). `bold`/`italic`/
        `underline`/`strikethrough`/`subscript`/`superscript`/`small_caps`/
        `all_caps` are booleans. `font` is a family name; `size` and `spacing`
        accept a number (points) or a unit string (`"12pt"`, `"1.5mm"`).
        `color` accepts a named colour, hex (`"#FF0000"`), or `(r, g, b)`.
        `highlight` is a named text-highlight colour (`"yellow"`, `"green"`, …,
        or `"none"`/`"auto"` to clear it) — a palette index, *not* an RGB.

        Bad colour/length/highlight input raises `OpError` (bad-input). Wrap in
        `doc.edit(...)` for atomic undo.
        """
        try:
            with _com.translate_com_errors():
                rng = self._range()
                _apply_font(
                    rng.Font,
                    bold=bold,
                    italic=italic,
                    underline=underline,
                    strikethrough=strikethrough,
                    font_name=font,
                    size=size,
                    color=color,
                    subscript=subscript,
                    superscript=superscript,
                    small_caps=small_caps,
                    all_caps=all_caps,
                    spacing=spacing,
                )
                if highlight is not None:
                    rng.HighlightColorIndex = _coerce_highlight(highlight)
        except (ValueError, TypeError) as e:
            raise OpError(str(e)) from e

    def format_info(self) -> dict[str, Any]:
        """The effective paragraph + character formatting on this anchor — the
        read mirror of [`format_paragraph`][wordlive.Anchor.format_paragraph] and
        [`format_run`][wordlive.Anchor.format_run]. Pure read.

        Returns `{anchor_id, style, paragraph, font}`. `style` is the applied
        paragraph style's name. `paragraph` and `font` carry one entry per field,
        each `{value, style, override}`:

        - `value` — the *effective* value (what's actually rendered);
        - `style` — the value the applied **style** would give on its own;
        - `override` — `True` when `value != style`, i.e. a **direct override**
          sits on top of the style (the signal the consistency linter rules act
          on). A mixed field (`value is None`) is never flagged as an override.

        `font.mixed` lists the character fields that read `wdUndefined` because
        they vary across the range's runs (e.g. a heading with one bold word) —
        those carry `value: null` rather than a bogus number. Lengths are in
        points; `color` is `#RRGGBB` (or `"auto"`); `alignment`/`line_spacing`
        use the same keywords the write verbs accept.

        The field vocabulary is identical to the write side, so a value read here
        can be written straight back through `format_paragraph`/`format_run`.
        """
        with _com.translate_com_errors():
            rng = self._range()
            style = rng.ParagraphStyle
            eff_para = _read_paragraph_format(rng.ParagraphFormat)
            sty_para = _read_paragraph_format(style.ParagraphFormat)
            eff_font, mixed = _read_font(rng.Font)
            sty_font, _ = _read_font(style.Font)
            style_name = str(style.NameLocal)

        def _annotate(eff: dict[str, Any], sty: dict[str, Any]) -> dict[str, Any]:
            return {
                key: {
                    "value": eff[key],
                    "style": sty[key],
                    "override": eff[key] is not None and eff[key] != sty[key],
                }
                for key in eff
            }

        font = _annotate(eff_font, sty_font)
        font["mixed"] = mixed
        return {
            "anchor_id": self.anchor_id,
            "style": style_name,
            "paragraph": _annotate(eff_para, sty_para),
            "font": font,
        }

    def set_shading(self, *, fill: Any = None, pattern: Any = None) -> None:
        """Set the background (fill) shading of this anchor's range.

        `fill` is a named colour, hex (`"#FFFF00"`), or `(r, g, b)` — applied to
        `Range.Shading.BackgroundPatternColor`. Because a `Cell` is an `Anchor`,
        this is also how you shade a table cell. `pattern` (a shading pattern/
        texture) is accepted for forward-compatibility but not yet applied —
        deferred. Bad colour input raises `OpError`. Wrap in `doc.edit(...)`.
        """
        try:
            with _com.translate_com_errors():
                if fill is not None:
                    self._range().Shading.BackgroundPatternColor = to_bgr(fill)
        except (ValueError, TypeError) as e:
            raise OpError(str(e)) from e

    def set_borders(
        self,
        *,
        sides: Any = "all",
        style: Any = "single",
        weight: Any = 0.5,
        color: Any = None,
    ) -> None:
        """Draw borders on this anchor's range (or cell — a `Cell` is an `Anchor`).

        `sides` is `"all"`/`"box"` (the default — four outer edges), a single
        edge (`"top"`/`"bottom"`/`"left"`/`"right"`), an interior gridline
        (`"horizontal"`/`"vertical"`, for multi-cell ranges), or a list of those.
        `style` is a line style (`"single"`, `"double"`, `"dot"`, `"dash"`, …, or
        `"none"` to remove). `weight` is the line width in points, snapped to
        Word's discrete set (0.25/0.5/0.75/1/1.5/2.25/3 pt). `color` is an
        optional border colour (name/hex/RGB).

        Page borders (`Section.Borders`) and table-wide borders (`Table.Borders`)
        are out of scope here — this sets per-range/per-cell borders. Bad input
        raises `OpError`. Wrap in `doc.edit(...)`.
        """
        try:
            edges = _resolve_border_sides(sides)
            line_style = _coerce_named(style, _LINE_STYLES, "border style")
            line_width = _coerce_line_weight(weight)
            bgr = to_bgr(color) if color is not None else None
            with _com.translate_com_errors():
                borders = self._range().Borders
                for edge in edges:
                    b = borders(edge)
                    b.LineStyle = line_style
                    b.LineWidth = line_width
                    if bgr is not None:
                        b.Color = bgr
        except (ValueError, TypeError) as e:
            raise OpError(str(e)) from e

    def add_tab_stop(self, position: Any, *, align: Any = "left", leader: Any = None) -> None:
        """Add a tab stop to this anchor's paragraph(s).

        `position` is the distance from the left margin in points (or a unit
        string like `"3in"`). `align` is `"left"`/`"center"`/`"right"`/
        `"decimal"`/`"bar"`. `leader` is an optional fill drawn up to the stop —
        `"dots"` (price lists / tables of contents), `"dashes"`, `"lines"`, … —
        defaulting to none. Maps to `ParagraphFormat.TabStops.Add`. Bad input
        raises `OpError`. Wrap in `doc.edit(...)`.
        """
        try:
            pos = to_points(position)
            al = _coerce_named(align, _TAB_ALIGN, "tab alignment")
            ld = (
                _coerce_named(leader, _TAB_LEADERS, "tab leader")
                if leader is not None
                else int(WdTabLeader.SPACES)
            )
            with _com.translate_com_errors():
                # Positional args: the `Leader=` keyword is dropped under pywin32
                # late binding, so pass Position, Alignment, Leader positionally.
                self._range().ParagraphFormat.TabStops.Add(pos, al, ld)
        except (ValueError, TypeError) as e:
            raise OpError(str(e)) from e

    def apply_list(self, list_type: str = "bulleted", *, continue_previous: bool = False) -> None:
        """Turn this anchor's paragraphs into a list.

        `list_type` is `"bulleted"`, `"numbered"`, or `"outline"` (the three
        `ListGalleries`). By default numbering starts fresh at 1; pass
        `continue_previous=True` to continue from a list immediately above.
        Raises `ValueError` for an unknown `list_type`.
        """
        gallery_type = _lists.gallery_for(list_type)  # ValueError before any mutation
        with _com.translate_com_errors():
            _lists.apply_list_template(
                self._range(), gallery_type, continue_previous=continue_previous
            )

    def remove_list(self) -> None:
        """Strip list formatting (bullets / numbers) from this anchor's paragraphs."""
        with _com.translate_com_errors():
            self._range().ListFormat.RemoveNumbers(NumberType=int(WdNumberType.ALL_NUMBERS))

    def list_info(self) -> dict[str, Any]:
        """Describe the list this anchor sits in: `{type, level, number, string}`.

        `type` is `"none"` when there's no list formatting, otherwise one of
        `"bulleted"`, `"numbered"`, `"outline"`, `"number-only"`, or `"mixed"`.
        `number` is the first paragraph's value, `string` its rendered marker.
        """
        with _com.translate_com_errors():
            return _lists.read_list_info(self._range())

    def location(self) -> dict[str, Any]:
        """Where this anchor sits in the laid-out document — a pure read.

        Returns `{page, end_page, line, column, in_table}`:

        - `page` / `end_page` — the 1-based pages the anchor's **first** and
          **last** characters fall on (equal for a collapsed/single-line anchor);
          the pair is the anchor's *page span*, so a section/table/image that
          straddles a page boundary reports both. `page` is what answers "what
          page is this on"; scan `paragraphs` and watch `page` step up to find
          "which paragraph starts page 2".
        - `line` / `column` — the first character's 1-based line and column in
          the page's text grid (`Range.Information`).
        - `in_table` — whether the anchor sits inside a table.

        Page/line numbers are only meaningful in print layout, so the document
        is **repaginated first** (content-neutral — it touches neither the
        user's selection, scroll, nor view), mirroring the guarantee a
        `snapshot` gives. No politeness concern: this mutates nothing — the
        document's `Saved` state is snapshotted and restored around the
        repaginate, which would otherwise flip Word's dirty bit.
        """
        with _com.translate_com_errors(), _com.preserve_saved(self._doc.com):
            rng = self._range()
            self._doc.com.Repaginate()
            start, end = int(rng.Start), int(rng.End)
            doc_com = self._doc.com
            head = doc_com.Range(start, start)
            tail = doc_com.Range(end, end)
            return {
                "page": int(head.Information(int(WdInformation.ACTIVE_END_PAGE_NUMBER))),
                "end_page": int(tail.Information(int(WdInformation.ACTIVE_END_PAGE_NUMBER))),
                "line": int(head.Information(int(WdInformation.FIRST_CHARACTER_LINE_NUMBER))),
                "column": int(head.Information(int(WdInformation.FIRST_CHARACTER_COLUMN_NUMBER))),
                "in_table": bool(rng.Information(int(WdInformation.WITH_IN_TABLE))),
            }

    def restart_numbering(self) -> None:
        """Restart this list's numbering at 1.

        Re-applies the range's current list template with "continue previous"
        off. Raises `ValueError` if the range isn't part of a list.
        """
        with _com.translate_com_errors():
            _lists.restart_numbering(self._range())

    def indent_list(self) -> None:
        """Demote this list item one level (e.g. level 1 -> 2)."""
        with _com.translate_com_errors():
            self._range().ListFormat.ListIndent()

    def outdent_list(self) -> None:
        """Promote this list item one level (e.g. level 2 -> 1)."""
        with _com.translate_com_errors():
            self._range().ListFormat.ListOutdent()

    def __repr__(self) -> str:
        return f"<{type(self).__name__} {self.name!r}>"


# ---------------------------------------------------------------------------
# Arbitrary ranges
# ---------------------------------------------------------------------------


class RangeAnchor(Anchor):
    """An anchor over an arbitrary character range — `doc.range(start, end)`.

    Unlike bookmarks/headings/cells, a range anchor names nothing in the
    document: it's a pair of absolute character offsets (UTF-16 code units, the
    same coordinates Word's `Document.Range(start, end)` uses and that
    `Document.find()` emits as `range:START-END`). It's the generic target when
    no named anchor exists — feed a `find()` hit straight into a `replace`, or
    drop a comment on an offset span.

    The anchor is ephemeral: offsets resolve live against the document on each
    access, so an edit elsewhere that shifts the text can leave it pointing at
    the wrong span. Resolve, act, discard. `set_text` keeps the anchor's own
    `end` in sync with the replacement so chained ops on the same instance stay
    consistent.
    """

    kind = "range"

    def __init__(self, doc: Document, start: int, end: int) -> None:
        start = int(start)
        end = int(end)
        if start < 0 or end < start:
            raise ValueError(f"invalid range offsets: start={start}, end={end}")
        super().__init__(doc, name=f"range:{start}-{end}")
        self._start = start
        self._end = end

    @property
    def start(self) -> int:
        return self._start

    @property
    def end(self) -> int:
        return self._end

    @property
    def anchor_id(self) -> str:
        return f"range:{self._start}-{self._end}"

    def _range(self) -> Any:
        return self._doc.com.Range(self._start, self._end)

    def set_text(self, text: str) -> None:
        with _com.translate_com_errors():
            rng = self._doc.com.Range(self._start, self._end)
            rng.Text = text
        # A Range.Text assignment resizes the span; keep our end in sync so a
        # follow-up read/op on the same anchor sees the replacement rather than
        # the stale coordinates. Word counts UTF-16 code units, not code points.
        self._end = self._start + _utf16_len(text)


# ---------------------------------------------------------------------------
# Start / end of document
# ---------------------------------------------------------------------------


class StartAnchor(Anchor):
    """A zero-width anchor at the very start of the document body — `doc.start`.

    The mirror of [`EndAnchor`][wordlive.EndAnchor]: the insertion point before
    the first paragraph. `doc.start` returns it and `anchor_by_id("start")`
    resolves it, so "prepend to the document" composes with the usual verbs and
    the CLI `--anchor-id` plumbing.

    Only the *prepend* direction is meaningful at a single start-point, so every
    insert verb lands text at the start: `insert_paragraph_before` /
    `insert_paragraph_after` add a new first paragraph (delegating to
    [`Document.prepend_paragraph`][wordlive.Document.prepend_paragraph]), and
    `insert_before` / `insert_after` / `set_text` prepend inline (delegating to
    [`Document.prepend`][wordlive.Document.prepend]). `text` is always empty and
    `delete()` is a no-op. `insert_image` and `apply_style` are inherited: they
    resolve to the collapsed start position.
    """

    kind = "start"

    def __init__(self, doc: Document) -> None:
        super().__init__(doc, name="start")

    @property
    def anchor_id(self) -> str:
        return "start"

    def _range(self) -> Any:
        # Collapsed at offset 0 — the position Document.prepend* writes to.
        return self._doc.com.Range(0, 0)

    def set_text(self, text: str) -> None:
        # Nothing to replace at the start-point — prepend instead.
        self._doc.prepend(text)

    def insert_after(self, text: str) -> None:
        self._doc.prepend(text)

    def insert_before(self, text: str) -> None:
        # A single start-point has no distinct "after"; prepending is the only
        # sensible reading, and it keeps `--anchor-id start` honest either way.
        self._doc.prepend(text)

    def insert_paragraph_after(self, text: str, style: str | None = None) -> None:
        self._doc.prepend_paragraph(text, style=style)

    def insert_paragraph_before(self, text: str, style: str | None = None) -> None:
        self._doc.prepend_paragraph(text, style=style)


class EndAnchor(Anchor):
    """A zero-width anchor at the very end of the document body — `doc.end`.

    The one position no content names: the insertion point past the last
    paragraph. `doc.end` returns it and `anchor_by_id("end")` resolves it, so
    "append to the document" composes with the same verbs and the same CLI
    `--anchor-id` plumbing as every other anchor — no `.com` drop needed.

    Only the *append* direction is meaningful at a single end-point, so every
    insert verb lands text at the end: `insert_paragraph_after` /
    `insert_paragraph_before` add a new final paragraph (delegating to
    [`Document.append_paragraph`][wordlive.Document.append_paragraph]), and
    `insert_after` / `insert_before` / `set_text` append inline (delegating to
    [`Document.append`][wordlive.Document.append]). `text` is always empty and
    `delete()` is a no-op — there is no content here to read or remove.
    `insert_image` and `apply_style` are inherited: they resolve to the
    collapsed end position, so an image lands at the end and a style falls on
    the final paragraph.
    """

    kind = "end"

    def __init__(self, doc: Document) -> None:
        super().__init__(doc, name="end")

    @property
    def anchor_id(self) -> str:
        return "end"

    def _range(self) -> Any:
        # Collapsed just before the final paragraph mark — the position
        # Document.append* writes to, and a safe target for the inherited verbs
        # (a zero-width span reads "" and deletes nothing).
        with _com.translate_com_errors():
            end = int(self._doc.com.Content.End)
        pos = max(0, end - 1)
        return self._doc.com.Range(pos, pos)

    def set_text(self, text: str) -> None:
        # Nothing to replace at the end-point — append instead.
        self._doc.append(text)

    def insert_after(self, text: str) -> None:
        self._doc.append(text)

    def insert_before(self, text: str) -> None:
        # A single end-point has no distinct "before"; appending is the only
        # sensible reading, and it keeps `--anchor-id end` honest either way.
        self._doc.append(text)

    def insert_paragraph_after(self, text: str, style: str | None = None) -> None:
        self._doc.append_paragraph(text, style=style)

    def insert_paragraph_before(self, text: str, style: str | None = None) -> None:
        self._doc.append_paragraph(text, style=style)


# ---------------------------------------------------------------------------
# Paragraphs
# ---------------------------------------------------------------------------


class Paragraph(Anchor):
    """A paragraph located by 1-based index over `doc.Paragraphs`.

    `para:N` addresses *any* paragraph — body text, headings, list items alike.
    `heading:N` is the same index space narrowed to heading paragraphs, so
    `para:5` and `heading:5` resolve to the same paragraph when paragraph 5 is a
    heading. A `Paragraph` inherits every anchor verb (`set_text`, `apply_style`,
    `format_paragraph`, `apply_list`, `insert_paragraph_before/after`, …).
    """

    kind = "paragraph"

    def __init__(self, doc: Document, index: int) -> None:
        super().__init__(doc, name=f"para:{index}")
        self._index = index

    @property
    def anchor_id(self) -> str:
        return f"para:{self._index}"

    @property
    def index(self) -> int:
        return self._index

    def _paragraph(self) -> Any:
        for idx, para in enumerate(self._doc.com.Paragraphs, start=1):
            if idx == self._index:
                # Keep .name informative for repr / error messages.
                self.name = paragraph_text(para) or self.name
                return para
        raise AnchorNotFoundError(
            "paragraph",
            f"para:{self._index}",
            hint=_stale_anchor_hint(self._doc.com, "para", self._index),
        )

    @property
    def level(self) -> int:
        with _com.translate_com_errors():
            return int(self._paragraph().OutlineLevel)

    @property
    def is_heading(self) -> bool:
        return self.level < 10

    def _range(self) -> Any:
        return self._paragraph().Range

    @property
    def text(self) -> str:
        with _com.translate_com_errors():
            return paragraph_text(self._paragraph())

    def set_text(self, text: str) -> None:
        with _com.translate_com_errors():
            para_range = self._paragraph().Range
            start = int(para_range.Start)
            end = int(para_range.End)
            # Preserve the trailing paragraph mark so the paragraph isn't merged
            # with the next one (same approach as Heading.set_text).
            inner = self._doc.com.Range(start, max(start, end - 1))
            inner.Text = text


class ParagraphCollection:
    """Indexable, iterable view over every paragraph in the document.

    Unlike `headings`, this includes body paragraphs and list items, not just
    heading paragraphs. Index by 1-based position (`doc.paragraphs[2]`); iterate
    for a `Paragraph` per paragraph. `list()` emits each paragraph's `start` /
    `end` offsets, so a body paragraph can be turned into a `range:START-END`
    insertion point for mid-paragraph edits.
    """

    def __init__(self, doc: Document) -> None:
        self._doc = doc

    def _count(self) -> int:
        with _com.translate_com_errors():
            return sum(1 for _ in self._doc.com.Paragraphs)

    def __len__(self) -> int:
        return self._count()

    def __getitem__(self, index: int) -> Paragraph:
        if isinstance(index, bool) or not isinstance(index, int):
            raise TypeError(f"paragraph index must be int, got {type(index).__name__}")
        if index < 1 or index > self._count():
            raise AnchorNotFoundError(
                "paragraph",
                f"para:{index}",
                hint=_stale_anchor_hint(self._doc.com, "para", index),
            )
        return Paragraph(self._doc, index)

    def __iter__(self) -> Iterator[Paragraph]:
        with _com.translate_com_errors():
            count = sum(1 for _ in self._doc.com.Paragraphs)
        for idx in range(1, count + 1):
            yield Paragraph(self._doc, idx)

    def at(self, offset: int) -> Paragraph | None:
        """Return the paragraph whose range contains `offset`, or None.

        Used to map a character offset (e.g. the cursor position) back to a
        `para:N` anchor.
        """
        with _com.translate_com_errors():
            for idx, para in enumerate(self._doc.com.Paragraphs, start=1):
                rng = para.Range
                if int(rng.Start) <= offset < int(rng.End):
                    return Paragraph(self._doc, idx)
        return None

    def list(self) -> list[dict[str, Any]]:
        """Every paragraph as `[{index, anchor_id, level, style, is_heading, start, end, text}, ...]`.

        `style` is the paragraph's applied Word style name (e.g. ``"Normal"``,
        ``"List Number"``, ``"Heading 2"``) — the handle to feed back into
        `apply_style` / a write's `style=` to mirror existing formatting, since
        `level` (Word's `OutlineLevel`) is `10` for *all* non-heading paragraphs
        and so can't distinguish a list item from body text. It's `None` if the
        style can't be read.
        """
        out: list[dict[str, Any]] = []
        with _com.translate_com_errors():
            for idx, para in enumerate(self._doc.com.Paragraphs, start=1):
                try:
                    level = int(para.OutlineLevel)
                except Exception:
                    level = 10
                try:
                    style: str | None = str(para.Range.Style.NameLocal)
                except Exception:
                    style = None
                rng = para.Range
                out.append(
                    {
                        "index": idx,
                        "anchor_id": f"para:{idx}",
                        "level": level,
                        "style": style,
                        "is_heading": level < 10,
                        "start": int(rng.Start),
                        "end": int(rng.End),
                        "text": paragraph_text(para),
                    }
                )
        return out


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

    def _apply(self, fn: Any, *args: Any, **kwargs: Any) -> None:
        """Run a `_shapes` size helper on this inline picture (`InlineShape` shares
        the `Width` / `Height` / `LockAspectRatio` surface), translating COM and
        bad-input errors into the wordlive hierarchy (`OpError`)."""
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
        """Every image as `{index, anchor_id, mime, width, height, alt_text, para}`.

        `mime` is the picture's content type (read from its package XML —
        ``None`` if the shape isn't a raster image, e.g. an embedded chart or OLE
        object). `width`/`height` are in points. `para` is the `para:N` anchor of
        the paragraph the image sits in (or ``None`` if it can't be located).
        Reads each image's content type but not its (potentially large) bytes —
        call [`read_image`][wordlive.Anchor.read_image] for those.
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
                        "alt_text": _safe_str(shape, "AlternativeText"),
                        "para": para_id,
                    }
                )
        return out


class EquationAnchor(Anchor):
    """A mathematical equation located by 1-based index — `equation:N`.

    Mirrors Word's own `OMaths(N)` ordering (document order). The anchor resolves
    to the equation's range, so `mathml` round-trips it back to MathML (via
    Office's own transform, without mutating the document) and `linear` reads its
    UnicodeMath form. `type` is ``"display"`` or ``"inline"``. Create equations
    with [`Anchor.insert_equation`][wordlive.Anchor.insert_equation]; discover
    them via [`doc.equations`][wordlive.Document.equations]. An equation isn't
    plain text, so `set_text` raises — delete and re-insert to change it.
    """

    kind = "equation"

    def __init__(self, doc: Document, index: int) -> None:
        super().__init__(doc, name=f"equation:{index}")
        self._index = index

    @property
    def index(self) -> int:
        return self._index

    @property
    def anchor_id(self) -> str:
        return f"equation:{self._index}"

    def _omath(self) -> Any:
        omaths = self._doc.com.OMaths
        n = int(omaths.Count)
        if not (1 <= self._index <= n):
            raise AnchorNotFoundError("equation", f"equation:{self._index}")
        return omaths.Item(self._index)

    def _range(self) -> Any:
        return self._omath().Range

    @property
    def type(self) -> str:
        """``"display"`` (its own centred line) or ``"inline"`` (in the text flow)."""
        with _com.translate_com_errors():
            # WdOMathType: wdOMathDisplay == 1, wdOMathInline == 0.
            return "display" if int(self._omath().Type) == 1 else "inline"

    @property
    def mathml(self) -> str:
        """The equation as MathML — a non-mutating read via Office's OMML→MathML transform."""
        with _com.translate_com_errors():
            package = str(self._omath().Range.WordOpenXML)
        return _equations.omml_to_mathml(package)

    @property
    def linear(self) -> str:
        """The equation's text in Word's built-up linear form (a compact preview).

        Reads the zone's text with the internal structure markers collapsed — a
        readable approximation of the math, not a precise round-trip. For
        fidelity use [`mathml`][wordlive.EquationAnchor.mathml].
        """
        with _com.translate_com_errors():
            raw = str(self._omath().Range.Text or "")
        return raw.replace("\r", "").replace("\x0b", "").strip()

    def set_text(self, text: str) -> None:
        raise OpError(
            "an equation anchor has no plain text to set; delete it and "
            "insert_equation(...) again to change it"
        )


class EquationCollection:
    """Read-only, iterable view over the document's equations (`doc.equations`).

    Index an equation by 1-based position (`doc.equations[2]`) to get an
    [`EquationAnchor`][wordlive.EquationAnchor] (`equation:N`), then `mathml` /
    `linear` to read it. `list()` summarises every equation — id, type, a linear
    preview, and the `para:N` it sits in. Positions match Word's own `OMaths(n)`
    ordering. The write mirror is any anchor's
    [`insert_equation`][wordlive.Anchor.insert_equation].
    """

    def __init__(self, doc: Document) -> None:
        self._doc = doc

    def __len__(self) -> int:
        with _com.translate_com_errors():
            return int(self._doc.com.OMaths.Count)

    def __getitem__(self, index: int) -> EquationAnchor:
        if isinstance(index, bool) or not isinstance(index, int):
            raise TypeError(f"equation index must be int, got {type(index).__name__}")
        n = len(self)
        if not (1 <= index <= n):
            raise AnchorNotFoundError("equation", str(index))
        return EquationAnchor(self._doc, index)

    def __iter__(self) -> Iterator[EquationAnchor]:
        with _com.translate_com_errors():
            count = int(self._doc.com.OMaths.Count)
        for i in range(1, count + 1):
            yield EquationAnchor(self._doc, i)

    def list(self) -> list[dict[str, Any]]:
        """Every equation as `{index, anchor_id, type, linear, para}`.

        `type` is ``"display"`` / ``"inline"``; `linear` is the built-up text as
        a compact preview (read [`EquationAnchor.mathml`][wordlive.EquationAnchor]
        for fidelity); `para` is the `para:N` the equation sits in (or ``None``).
        Reads no XML, so this is cheap to call over a whole document.
        """
        out: list[dict[str, Any]] = []
        with _com.translate_com_errors():
            omaths = self._doc.com.OMaths
            count = int(omaths.Count)
            for i in range(1, count + 1):
                zone = omaths.Item(i)
                rng = zone.Range
                try:
                    start = int(rng.Start)
                except Exception:
                    start = None
                try:
                    eq_type = "display" if int(zone.Type) == 1 else "inline"
                except Exception:
                    eq_type = "inline"
                linear = str(rng.Text or "").replace("\r", "").replace("\x0b", "").strip()
                para_id: str | None = None
                if start is not None:
                    para = self._doc.paragraphs.at(start)
                    para_id = para.anchor_id if para is not None else None
                out.append(
                    {
                        "index": i,
                        "anchor_id": f"equation:{i}",
                        "type": eq_type,
                        "linear": linear,
                        "para": para_id,
                    }
                )
        return out


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


class ChartAnchor(Anchor):
    """An Excel-backed chart located by 1-based index — `chart:N`.

    Indexes the document's chart inline shapes in document order. The anchor
    resolves to the chart's range, so it inherits `apply_style` / formatting like
    any anchor. `chart_type` reports the kind string (``"bar"`` / ``"pie"`` /
    ``"line"`` / ``"scatter"``) and `title` the chart title — metadata only:
    charts are inserted with a broken data link, so the underlying series data is
    static and isn't read back. Create charts with
    [`Anchor.insert_chart`][wordlive.Anchor.insert_chart]; discover them via
    [`doc.charts`][wordlive.Document.charts]. A chart isn't plain text, so
    `set_text` raises.
    """

    kind = "chart"

    def __init__(self, doc: Document, index: int) -> None:
        super().__init__(doc, name=f"chart:{index}")
        self._index = index

    @property
    def index(self) -> int:
        return self._index

    @property
    def anchor_id(self) -> str:
        return f"chart:{self._index}"

    def _shape(self) -> Any:
        shapes = _charts.chart_shapes(self._doc.com)
        if not (1 <= self._index <= len(shapes)):
            raise AnchorNotFoundError("chart", f"chart:{self._index}")
        return shapes[self._index - 1]

    def _range(self) -> Any:
        return self._shape().Range

    @property
    def chart_type(self) -> str:
        """The chart kind — ``"bar"`` / ``"pie"`` / ``"line"`` / ``"scatter"`` (or the raw int)."""
        with _com.translate_com_errors():
            ctype = int(self._shape().Chart.ChartType)
        return _charts.XL_TO_KIND.get(ctype, str(ctype))

    @property
    def title(self) -> str | None:
        """The chart's title, or ``None`` if it has none."""
        with _com.translate_com_errors():
            chart = self._shape().Chart
            if not bool(chart.HasTitle):
                return None
            return str(chart.ChartTitle.Text or "")

    @property
    def chart_style(self) -> int:
        """The built-in design-gallery style id (`Chart.ChartStyle`)."""
        with _com.translate_com_errors():
            return int(self._shape().Chart.ChartStyle)

    @property
    def has_legend(self) -> bool:
        """Whether the chart currently shows a legend."""
        with _com.translate_com_errors():
            return bool(self._shape().Chart.HasLegend)

    def format(
        self,
        *,
        title: Any = _charts._UNSET,
        legend: bool | None = None,
        legend_position: str | None = None,
        chart_style: int | None = None,
        background: Any = None,
        plot_background: Any = None,
        font: str | None = None,
        font_size: Any = None,
        font_color: Any = None,
        data_labels: bool | None = None,
        data_label_format: str | None = None,
        chart_type: str | None = None,
    ) -> ChartAnchor:
        """Apply whole-chart / design formatting — Word's chart "Design" tab.

        All kwargs are optional and tri-state; only the ones you pass are written.
        `title=None` clears the chart title (omit it to leave it). `legend`
        toggles the legend; `legend_position` (`"right"`/`"left"`/`"top"`/
        `"bottom"`/`"corner"`) implies it's shown. `chart_style` is the built-in
        design-gallery int. `background` / `plot_background` fill the chart and
        plot areas; `font` / `font_size` / `font_color` set the whole-chart font.
        `data_labels` toggles point labels on every series, `data_label_format`
        is their number format. `chart_type` (`"bar"`/`"pie"`/`"line"`/
        `"scatter"`) re-types the chart in place. Operates on the static chart —
        no Excel needed. Returns `self` (chainable); wrap in `doc.edit(...)` for
        atomic undo. Bad input raises `OpError`.
        """
        self._apply(
            _charts.apply_chart_format,
            title=title,
            legend=legend,
            legend_position=legend_position,
            chart_style=chart_style,
            background=background,
            plot_background=plot_background,
            font=font,
            font_size=font_size,
            font_color=font_color,
            data_labels=data_labels,
            data_label_format=data_label_format,
            chart_type=chart_type,
        )
        return self

    def set_axis(
        self,
        which: str,
        *,
        title: Any = _charts._UNSET,
        minimum: Any = None,
        maximum: Any = None,
        scale: str | None = None,
        number_format: str | None = None,
        gridlines: bool | None = None,
    ) -> ChartAnchor:
        """Format one axis. `which` is ``"value"``/``"y"`` or ``"category"``/``"x"``.

        Tri-state: `title=None` clears the axis title; `minimum`/`maximum` set the
        scale bounds; `scale` is ``"linear"`` or ``"log"`` (log is ideal for
        order-of-magnitude data); `number_format` is the tick-label format string;
        `gridlines` toggles major gridlines. Returns `self`. Bad input raises
        `OpError`.
        """
        self._apply(
            _charts.apply_axis_format,
            which,
            title=title,
            minimum=minimum,
            maximum=maximum,
            scale=scale,
            number_format=number_format,
            gridlines=gridlines,
        )
        return self

    def add_trendline(
        self,
        *,
        series: int = 1,
        kind: str = "linear",
        display_equation: bool = False,
        display_r_squared: bool = False,
        forward: Any = None,
        backward: Any = None,
    ) -> ChartAnchor:
        """Fit a trendline to a series (1-based `series`).

        `kind` is ``"linear"``, ``"exponential"``, ``"logarithmic"``,
        ``"moving_average"``, ``"polynomial"``, or ``"power"``. `display_equation`
        / `display_r_squared` annotate the fit — a power/exponential fit with the
        equation literally draws the law of best fit. `forward` / `backward`
        extend the line that many units past the data. Returns `self`. Bad input
        raises `OpError`.
        """
        self._apply(
            _charts.add_trendline,
            series=series,
            kind=kind,
            display_equation=display_equation,
            display_r_squared=display_r_squared,
            forward=forward,
            backward=backward,
        )
        return self

    def set_series_color(
        self, color: Any, *, series: int = 1, point: int | None = None
    ) -> ChartAnchor:
        """Recolour a whole series, or a single 1-based `point` (bar / pie slice).

        `color` is a named colour, hex (`"#2E86C1"`), or `(r, g, b)`. Omit `point`
        to colour the entire series; pass it to vary one bar / slice / marker.
        Sets the line/marker colour too where the series has one (line/scatter).
        Returns `self`. Bad input raises `OpError`.
        """
        self._apply(_charts.set_series_color, color, series=series, point=point)
        return self

    def _apply(self, fn: Any, *args: Any, **kwargs: Any) -> None:
        """Run a `_charts` formatting helper on this chart's `Chart`, translating
        COM and bad-input errors into the wordlive hierarchy (`OpError`)."""
        chart = self._shape().Chart
        try:
            with _com.translate_com_errors():
                fn(chart, *args, **kwargs)
        except (ValueError, TypeError) as e:
            raise OpError(str(e)) from e

    def set_text(self, text: str) -> None:
        raise OpError(
            "a chart anchor has no plain text to set; delete it and "
            "insert_chart(...) again to change it"
        )


class ChartCollection:
    """Read-only, iterable view over the document's charts (`doc.charts`).

    Index a chart by 1-based position (`doc.charts[2]`) to get a
    [`ChartAnchor`][wordlive.ChartAnchor] (`chart:N`); `list()` summarises every
    chart — id, kind, title, and the `para:N` it sits in. Positions follow
    document order. Metadata only — charts are inserted with their data link
    broken (static data), so reading the series back is deferred. The write
    mirror is any anchor's [`insert_chart`][wordlive.Anchor.insert_chart].
    """

    def __init__(self, doc: Document) -> None:
        self._doc = doc

    def __len__(self) -> int:
        with _com.translate_com_errors():
            return len(_charts.chart_shapes(self._doc.com))

    def __getitem__(self, index: int) -> ChartAnchor:
        if isinstance(index, bool) or not isinstance(index, int):
            raise TypeError(f"chart index must be int, got {type(index).__name__}")
        n = len(self)
        if not (1 <= index <= n):
            raise AnchorNotFoundError("chart", str(index))
        return ChartAnchor(self._doc, index)

    def __iter__(self) -> Iterator[ChartAnchor]:
        for i in range(1, len(self) + 1):
            yield ChartAnchor(self._doc, i)

    def list(self) -> list[dict[str, Any]]:
        """Every chart as `{index, anchor_id, kind, title, chart_style, has_legend, para}`.

        `kind` is the chart-type string; `title` the chart title (or ``None``);
        `chart_style` the design-gallery id; `has_legend` whether a legend shows;
        `para` the `para:N` the chart sits in. Touches only `ChartType` /
        `ChartTitle` / `ChartStyle` / `HasLegend` (never the series data), so it's
        cheap and Word-stable.
        """
        out: list[dict[str, Any]] = []
        with _com.translate_com_errors():
            shapes = _charts.chart_shapes(self._doc.com)
            for i, shape in enumerate(shapes, start=1):
                chart = shape.Chart
                try:
                    kind: str | None = _charts.XL_TO_KIND.get(
                        int(chart.ChartType), str(int(chart.ChartType))
                    )
                except Exception:
                    kind = None
                try:
                    title = str(chart.ChartTitle.Text) if bool(chart.HasTitle) else None
                except Exception:
                    title = None
                try:
                    chart_style: int | None = int(chart.ChartStyle)
                except Exception:
                    chart_style = None
                try:
                    has_legend: bool | None = bool(chart.HasLegend)
                except Exception:
                    has_legend = None
                try:
                    start = int(shape.Range.Start)
                except Exception:
                    start = None
                para_id: str | None = None
                if start is not None:
                    para = self._doc.paragraphs.at(start)
                    para_id = para.anchor_id if para is not None else None
                out.append(
                    {
                        "index": i,
                        "anchor_id": f"chart:{i}",
                        "kind": kind,
                        "title": title,
                        "chart_style": chart_style,
                        "has_legend": has_legend,
                        "para": para_id,
                    }
                )
        return out


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

    def set_wrap(self, wrap: str) -> ShapeAnchor:
        """Set how body text flows around the shape — ``"square"`` / ``"tight"`` /
        ``"through"`` / ``"top-bottom"`` / ``"front"`` / ``"behind"``. Returns
        `self` (chainable). Bad input raises `OpError`."""
        self._apply(_shapes.apply_shape_wrap, wrap)
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
        height, rotation, z_order, wrap, alt_text, has_text, para}`.

        `shape_type` is the kind string; `width` / `height` are points; `rotation`
        the clockwise angle in degrees; `z_order` the 1-based stacking position;
        `wrap` the text-wrap keyword; `has_text` whether a text frame holds text;
        `para` the `para:N` the shape is anchored in.
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
# Bookmarks
# ---------------------------------------------------------------------------


class Bookmark(Anchor):
    kind = "bookmark"

    # Set when this bookmark is a wordlive durable handle (`_wl_<code>`): the
    # anchor then reports `pin:<code>` instead of `bookmark:_wl_<code>`.
    _pin_code: str | None = None

    @classmethod
    def pin(cls, doc: Document, code: str) -> Bookmark:
        """A `Bookmark` over the hidden `_wl_<code>` pin, reporting `pin:<code>`."""
        bm = cls(doc, _pin_name_for(code))
        bm._pin_code = code
        return bm

    @property
    def anchor_id(self) -> str:
        if self._pin_code is not None:
            return f"pin:{self._pin_code}"
        return f"bookmark:{self.name}"

    def _range(self) -> Any:
        doc_com = self._doc.com
        if not doc_com.Bookmarks.Exists(self.name):
            raise AnchorNotFoundError("bookmark", self.name)
        return doc_com.Bookmarks(self.name).Range

    def set_text(self, text: str) -> None:
        with _com.translate_com_errors():
            doc_com = self._doc.com
            if not doc_com.Bookmarks.Exists(self.name):
                raise AnchorNotFoundError("bookmark", self.name)
            rng = doc_com.Bookmarks(self.name).Range
            start = int(rng.Start)
            rng.Text = text
            # Setting Range.Text deletes the bookmark; re-add covering the new content.
            # Word measures Range offsets in UTF-16 code units, not Python code points.
            new_end = start + _utf16_len(text)
            new_rng = doc_com.Range(start, new_end)
            doc_com.Bookmarks.Add(Name=self.name, Range=new_rng)


def _bookmarks_including_hidden(doc_com: Any) -> list[Any]:
    """Every bookmark, *including* Word's hidden (leading-underscore) ones.

    Word omits underscore-prefixed bookmarks — its own `_Toc`/`_Ref` anchors and
    wordlive's `_wl_` pins — from `Document.Bookmarks` iteration unless the
    collection's `ShowHidden` flag is on (a real-Word behaviour the fake COM
    fixture doesn't model). We flip it on for the read and restore it after, so
    pin enumeration / idempotency see the `_wl_` handles. The fake has no
    `ShowHidden`, so the toggle is best-effort.
    """
    bms = doc_com.Bookmarks
    previous: bool | None
    try:
        previous = bool(bms.ShowHidden)
        bms.ShowHidden = True
    except Exception:
        previous = None
    try:
        return list(bms)
    finally:
        if previous is not None:
            try:
                bms.ShowHidden = previous
            except Exception:
                pass


def _is_user_bookmark(name: str) -> bool:
    """Word auto-creates internal bookmarks for TOC entries, cross-references,
    and form-field anchors — all of them named with a leading underscore. Those
    are noise for the user-facing `list()` / iteration paths; agents addressing
    them by exact name (via `bookmarks[name]`) still work.
    """
    return not name.startswith("_")


class BookmarkCollection:
    """Indexable view over a document's bookmarks.

    `list()` and iteration return only user-visible bookmarks. Word's hidden
    bookmarks (`_Toc...`, `_Ref...`, etc.) are filtered out by default; address
    them by their exact name through `bookmarks[name]` if you need them.
    """

    def __init__(self, doc: Document) -> None:
        self._doc = doc

    def __getitem__(self, name: str) -> Bookmark:
        with _com.translate_com_errors():
            if not self._doc.com.Bookmarks.Exists(name):
                raise AnchorNotFoundError("bookmark", name)
        return Bookmark(self._doc, name)

    def __contains__(self, name: object) -> bool:
        if not isinstance(name, str):
            return False
        with _com.translate_com_errors():
            return bool(self._doc.com.Bookmarks.Exists(name))

    def add(self, name: str, anchor: Anchor | str) -> Bookmark:
        """Create a bookmark named `name` over `anchor`'s range and return it.

        `anchor` is an [`Anchor`][wordlive.Anchor] or an anchor id string
        (resolved via `doc.anchor_by_id`). `name` is validated against Word's
        rules — it must start with a letter and contain only letters, digits, and
        underscores (no spaces), max 40 characters — and an invalid name raises
        `OpError` *before* anything is created. Adding a bookmark with an existing
        name moves it to the new range (Word's own behaviour). This is the
        prerequisite for internal hyperlinks
        ([`Anchor.link_to`][wordlive.Anchor.link_to]) and cross-references
        ([`Anchor.insert_cross_reference`][wordlive.Anchor.insert_cross_reference]).
        Wrap in `doc.edit(...)` for atomic undo.
        """
        _validate_bookmark_name(name)
        resolved = self._doc.anchor_by_id(anchor) if isinstance(anchor, str) else anchor
        with _com.translate_com_errors():
            rng = resolved.com
            self._doc.com.Bookmarks.Add(Name=name, Range=rng)
        return Bookmark(self._doc, name)

    def list(self, *, include_hidden: bool = False) -> list[str]:
        """Names of every user-visible bookmark in document order.

        Set `include_hidden=True` to also return Word's internal bookmarks
        (TOC entries, cross-references, etc.) whose names start with `_`.
        """
        with _com.translate_com_errors():
            if include_hidden:
                return [str(bm.Name) for bm in _bookmarks_including_hidden(self._doc.com)]
            names = [str(bm.Name) for bm in self._doc.com.Bookmarks]
        return [n for n in names if _is_user_bookmark(n)]

    def __iter__(self) -> Iterator[Bookmark]:
        for name in self.list():
            yield Bookmark(self._doc, name)


# ---------------------------------------------------------------------------
# Content controls
# ---------------------------------------------------------------------------


def _cc_by_name(doc_com: Any, name: str) -> Any | None:
    """Find a content control by its Title (Tag falls back). Returns None if missing.

    Reject empty `name` explicitly — many content controls have neither a
    Title nor a Tag, and the naive `cc.Title or "" == ""` test would match
    the first such control. Callers asking for `""` get `None` instead.
    """
    if not name:
        return None
    for cc in doc_com.ContentControls:
        if str(cc.Title or "") == name or str(cc.Tag or "") == name:
            return cc
    return None


class ContentControl(Anchor):
    kind = "content_control"

    def __init__(self, doc: Document, name: str, *, com: Any | None = None) -> None:
        super().__init__(doc, name)
        # A freshly created control caches its live COM object, so the returned
        # wrapper resolves even when unnamed (and survives edits — Word maintains
        # the control's identity). Named lookups still go through `_cc_by_name`.
        self._cc_com = com

    @property
    def anchor_id(self) -> str:
        return f"cc:{self.name}"

    def _cc(self) -> Any:
        if self._cc_com is not None:
            return self._cc_com
        cc = _cc_by_name(self._doc.com, self.name)
        if cc is None:
            raise AnchorNotFoundError("content_control", self.name)
        return cc

    def _range(self) -> Any:
        return self._cc().Range

    @property
    def text(self) -> str:
        with _com.translate_com_errors():
            cc = self._cc()
            return range_text(cc.Range)

    def set_text(self, text: str) -> None:
        with _com.translate_com_errors():
            cc = self._cc()
            cc.Range.Text = text

    def set_properties(
        self,
        *,
        title: Any = _charts._UNSET,
        tag: Any = _charts._UNSET,
        lock_contents: bool | None = None,
        lock_control: bool | None = None,
    ) -> ContentControl:
        """Re-set this control's metadata in place — no delete + reinsert.

        Tri-state: omit a field to leave it untouched, pass a string to set it,
        or `None` (equivalently ``""``) to clear `title` / `tag`. `lock_contents`
        stops the user editing the value; `lock_control` stops them deleting the
        control — pass a bool to set either, omit to leave. Renaming the title
        (or the tag, when untitled) changes the control's `cc:NAME` anchor id;
        the returned handle keeps working regardless. Returns `self` (chainable);
        wrap in `doc.edit(...)` for atomic undo. Bad input raises `OpError`.
        """
        try:
            with _com.translate_com_errors():
                cc = self._cc()
                if title is not _charts._UNSET:
                    cc.Title = "" if title is None else str(title)
                if tag is not _charts._UNSET:
                    cc.Tag = "" if tag is None else str(tag)
                if lock_contents is not None:
                    cc.LockContents = bool(lock_contents)
                if lock_control is not None:
                    cc.LockContentControl = bool(lock_control)
                if title is not _charts._UNSET or tag is not _charts._UNSET:
                    # Keep anchor_id honest after a rename (title beats tag).
                    self.name = str(cc.Title or cc.Tag or "")
        except (ValueError, TypeError) as e:
            raise OpError(str(e)) from e
        return self

    def set_items(self, items: list[Any]) -> ContentControl:
        """Replace a combo-box / dropdown's choice list — no delete + reinsert.

        `items` is the full new list (it replaces the existing entries, not
        appends); each is a string or a `{"text": ..., "value": ...}` dict.
        Only valid on a ``combo_box`` / ``dropdown`` control — any other kind
        raises `OpError`. Returns `self` (chainable); wrap in `doc.edit(...)`
        for atomic undo. Bad input raises `OpError`.
        """
        try:
            with _com.translate_com_errors():
                cc = self._cc()
                if cc.Type not in _CC_LIST_TYPES:
                    raise ValueError(
                        "set_items is only valid for a 'combo_box' or 'dropdown' control"
                    )
                pairs = _normalize_cc_items(items)
                entries = cc.DropdownListEntries
                entries.Clear()
                for entry_text, value in pairs:
                    entries.Add(entry_text, value)  # positional (Text, Value)
        except (ValueError, TypeError) as e:
            raise OpError(str(e)) from e
        return self


class ContentControlCollection:
    def __init__(self, doc: Document) -> None:
        self._doc = doc

    def __getitem__(self, name: str) -> ContentControl:
        with _com.translate_com_errors():
            if _cc_by_name(self._doc.com, name) is None:
                raise AnchorNotFoundError("content_control", name)
        return ContentControl(self._doc, name)

    def __contains__(self, name: object) -> bool:
        if not isinstance(name, str):
            return False
        with _com.translate_com_errors():
            return _cc_by_name(self._doc.com, name) is not None

    def list(self) -> list[str]:
        with _com.translate_com_errors():
            names: list[str] = []
            for cc in self._doc.com.ContentControls:
                names.append(str(cc.Title or cc.Tag or ""))
            return names

    def __iter__(self) -> Iterator[ContentControl]:
        for name in self.list():
            if name:
                yield ContentControl(self._doc, name)

    def add(self, anchor: Anchor | str, kind: str = "rich_text", **kwargs: Any) -> ContentControl:
        """Create a content control over an anchor and return it.

        Symmetric with [`bookmarks.add`][wordlive.BookmarkCollection.add]: a
        document-level entry point for the per-anchor
        [`insert_content_control`][wordlive.Anchor.insert_content_control].
        `anchor` is an [`Anchor`][wordlive.Anchor] or an anchor-id string
        (`range:START-END`, `cc:NAME`, `heading:N`, …); `kind` and the keyword
        options (`title`, `tag`, `items`, `where`, `lock_contents`,
        `lock_control`) pass straight through. Wrap in `doc.edit(...)` for atomic
        undo.
        """
        target = self._doc.anchor_by_id(anchor) if isinstance(anchor, str) else anchor
        return target.insert_content_control(kind, **kwargs)


# ---------------------------------------------------------------------------
# Headings
# ---------------------------------------------------------------------------


def paragraph_text(para: Any) -> str:
    """Heading text minus the trailing paragraph mark, inline shapes tokenized."""
    return range_text(para.Range).rstrip("\r\n\x07")


def _find_heading_paragraph(doc_com: Any, name: str) -> tuple[Any, int] | None:
    """Locate a heading paragraph by visible text. Returns (Paragraph, 1-based index)."""
    for idx, para in enumerate(doc_com.Paragraphs, start=1):
        try:
            level = int(para.OutlineLevel)
        except Exception:
            continue
        if level >= 10:  # WdOutlineLevel: 1-9 are headings; 10 is body text
            continue
        if paragraph_text(para) == name:
            return para, idx
    return None


def _section_range(doc_com: Any, target_para: Any, target_level: int) -> Any:
    """COM Range from the end of `target_para` to the next paragraph whose
    OutlineLevel is a heading and `<= target_level` — or to the end of the
    document's last paragraph if no such boundary exists.
    """
    paragraphs = list(doc_com.Paragraphs)
    target_start = int(target_para.Range.Start)

    idx: int | None = None
    for i, p in enumerate(paragraphs):
        try:
            if int(p.Range.Start) == target_start:
                idx = i
                break
        except Exception:
            continue
    if idx is None:
        end = int(target_para.Range.End)
        return doc_com.Range(end, end)

    section_start = int(target_para.Range.End)
    section_end: int | None = None
    for p in paragraphs[idx + 1 :]:
        try:
            lvl = int(p.OutlineLevel)
        except Exception:
            continue
        if lvl < 10 and lvl <= target_level:
            section_end = int(p.Range.Start)
            break
    if section_end is None:
        try:
            section_end = int(paragraphs[-1].Range.End)
        except Exception:
            section_end = section_start
    return doc_com.Range(section_start, section_end)


class Heading(Anchor):
    kind = "heading"

    def _paragraph(self) -> Any:
        found = _find_heading_paragraph(self._doc.com, self.name)
        if found is None:
            raise AnchorNotFoundError("heading", self.name)
        return found[0]

    def _paragraph_and_index(self) -> tuple[Any, int]:
        """Default lookup goes by visible text; subclasses can override."""
        found = _find_heading_paragraph(self._doc.com, self.name)
        if found is None:
            raise AnchorNotFoundError("heading", self.name)
        return found

    @property
    def anchor_id(self) -> str:
        with _com.translate_com_errors():
            _, idx = self._paragraph_and_index()
        return f"heading:{idx}"

    @property
    def level(self) -> int:
        with _com.translate_com_errors():
            return int(self._paragraph().OutlineLevel)

    def section_range(self) -> Any:
        """COM Range covering the body under this heading.

        Spans from the end of the heading paragraph to the start of the next
        heading whose level is `<=` this one's (or to the end of the document
        if no such heading exists). Excludes the heading paragraph itself.
        """
        with _com.translate_com_errors():
            para = self._paragraph()
            level = int(para.OutlineLevel)
            return _section_range(self._doc.com, para, level)

    def section_text(self) -> str:
        """Plain text of the body under this heading."""
        with _com.translate_com_errors():
            return str(self.section_range().Text or "")

    def replace_section_body(self, body: Any, *, markdown: bool = False) -> RangeAnchor:
        """Rewrite this heading's body, leaving the heading paragraph intact.

        The "rewrite section X" workflow: clears the span under this heading
        (`section_range`, up to the next same-or-higher heading) and inserts
        `body` after the heading. With ``markdown=False`` (default) `body` is the
        `insert_block` items shape (or a bare string); with ``markdown=True``
        `body` is a constrained-Markdown string routed through `insert_markdown`.
        Returns the new body's spanning [`RangeAnchor`][wordlive.RangeAnchor].
        Wrap in `doc.edit(...)` for atomic undo.
        """
        with _com.translate_com_errors():
            span = self.section_range()
            doc_com = self._doc.com
            doc_com.Range(int(span.Start), int(span.End)).Delete()
        if markdown:
            if not isinstance(body, str):
                raise OpError("replace_section_body with markdown=True requires a string body")
            return self.insert_markdown(body, where="after")
        if isinstance(body, str):
            body = [body]
        if not isinstance(body, list):
            raise OpError(
                f"replace_section_body body must be a string or list; got {type(body).__name__}"
            )
        return self.insert_block(body, where="after")

    def _range(self) -> Any:
        return self._paragraph().Range

    @property
    def text(self) -> str:
        with _com.translate_com_errors():
            return paragraph_text(self._paragraph())

    def set_text(self, text: str) -> None:
        with _com.translate_com_errors():
            para_range = self._paragraph().Range
            start = int(para_range.Start)
            end = int(para_range.End)
            # Preserve the trailing paragraph mark.
            inner = self._doc.com.Range(start, max(start, end - 1))
            inner.Text = text

    # insert_paragraph_after / insert_paragraph_before are inherited from Anchor;
    # for a Heading, _range() is the heading paragraph, so "after" lands a new
    # paragraph just below the heading (the original v0 behaviour).


class HeadingCollection:
    """Iterable, indexable view over a document's headings.

    Symmetric with `BookmarkCollection` and `ContentControlCollection`:

        for h in doc.headings:           # iteration → Heading per heading paragraph
            ...
        doc.headings["Risks"]            # by visible text
        doc.headings[3]                  # by 1-based paragraph index
        "Risks" in doc.headings          # membership
        doc.headings.list()              # same shape as doc.outline()
    """

    def __init__(self, doc: Document) -> None:
        self._doc = doc

    def __getitem__(self, key: str | int) -> Heading:
        if isinstance(key, bool):
            # bool is a subclass of int; reject before the int branch matches.
            raise TypeError(f"heading key must be str or int, got {type(key).__name__}")
        if isinstance(key, int):
            return _IndexedHeading(self._doc, key)
        if isinstance(key, str):
            with _com.translate_com_errors():
                if _find_heading_paragraph(self._doc.com, key) is None:
                    raise AnchorNotFoundError("heading", key)
            return Heading(self._doc, key)
        raise TypeError(f"heading key must be str or int, got {type(key).__name__}")

    def __contains__(self, key: object) -> bool:
        if isinstance(key, bool):
            return False
        if isinstance(key, int):
            # 1-based paragraph index must reference an actual heading paragraph.
            with _com.translate_com_errors():
                for idx, para in enumerate(self._doc.com.Paragraphs, start=1):
                    if idx != key:
                        continue
                    try:
                        lvl = int(para.OutlineLevel)
                    except Exception:
                        return False
                    return lvl < 10
            return False
        if not isinstance(key, str):
            return False
        with _com.translate_com_errors():
            return _find_heading_paragraph(self._doc.com, key) is not None

    def list(self) -> list[dict[str, Any]]:
        """Same shape as `Document.outline()` — `[{level, text, anchor_id}, ...]`."""
        out: list[dict[str, Any]] = []
        with _com.translate_com_errors():
            for idx, para in enumerate(self._doc.com.Paragraphs, start=1):
                try:
                    level = int(para.OutlineLevel)
                except Exception:
                    continue
                if level >= 10:
                    continue
                out.append(
                    {
                        "level": level,
                        "text": paragraph_text(para),
                        "anchor_id": f"heading:{idx}",
                    }
                )
        return out

    def __iter__(self) -> Iterator[Heading]:
        for entry in self.list():
            # Each entry's anchor_id is `heading:N`; index-based heading
            # disambiguates duplicate visible text.
            idx = int(entry["anchor_id"].split(":", 1)[1])
            yield _IndexedHeading(self._doc, idx)


class _IndexedHeading(Heading):
    """A Heading located by 1-based paragraph index — used by anchor_by_id('heading:N').

    Disambiguates duplicate heading text. The display name is set to the resolved
    heading text the first time `_paragraph()` succeeds so error messages and
    `.name` reads stay informative.
    """

    def __init__(self, doc: Document, paragraph_index: int) -> None:
        super().__init__(doc, name=f"heading:{paragraph_index}")
        self._paragraph_index = paragraph_index

    @property
    def anchor_id(self) -> str:
        return f"heading:{self._paragraph_index}"

    def _paragraph(self) -> Any:
        for idx, para in enumerate(self._doc.com.Paragraphs, start=1):
            if idx != self._paragraph_index:
                continue
            try:
                level = int(para.OutlineLevel)
            except Exception:
                break
            if level >= 10:
                break
            self.name = paragraph_text(para) or self.name
            return para
        raise AnchorNotFoundError(
            "heading",
            f"heading:{self._paragraph_index}",
            hint=_stale_anchor_hint(self._doc.com, "heading", self._paragraph_index),
        )

    def _paragraph_and_index(self) -> tuple[Any, int]:
        return self._paragraph(), self._paragraph_index
