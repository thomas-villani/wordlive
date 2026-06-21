"""Lists & numbering — apply / read / restart bullets and numbered lists.

List operations act on a *range's paragraphs*, so the verbs themselves live on
the base `Anchor` (`apply_list`, `remove_list`, `list_info`, `restart_numbering`,
`indent_list`, `outdent_list`) and delegate to the small COM helpers here. This
module also holds the document-scoped `doc.lists` discovery collection.

Word's list model is notoriously fiddly: numbering comes from a `ListTemplate`
pulled out of one of three `Application.ListGalleries` (bullet / number / outline),
and "restart vs. continue the previous list" is a boolean on
`ListFormat.ApplyListTemplate`. We expose the 80%-useful shape — apply a fresh
bulleted/numbered/outline list, read what list a range is in, restart numbering,
and step a list item in/out a level.

For the remaining 20%, `apply_list_format` **authors a custom list template**
from per-level specs (`Document.ListTemplates.Add` + per-`ListLevel` number
format / style / bullet glyph / indentation / marker font), and
`read_list_levels` is its read mirror. Per-level mutation is settable under late
binding (live-probed 2026-06-21) — the one trap is that a bullet level must be
authored via `NumberFormat` + a symbol font, never `NumberStyle=bullet` (which
raises).
"""

from __future__ import annotations

from collections.abc import Iterator
from typing import TYPE_CHECKING, Any

from . import _com
from ._format import to_bgr, to_points
from .constants import (
    WdDefaultListBehavior,
    WdListApplyTo,
    WdListGalleryType,
    WdListLevelAlignment,
    WdListNumberStyle,
    WdListType,
    WdTrailingCharacter,
)
from .exceptions import AnchorNotFoundError, OpError

if TYPE_CHECKING:
    from ._anchors import RangeAnchor
    from ._document import Document


# Accepted `list_type` strings -> the gallery they apply from. The three
# canonical names are bulleted / numbered / outline; common variants alias on.
_GALLERY_FOR: dict[str, WdListGalleryType] = {
    "bulleted": WdListGalleryType.BULLET,
    "bullet": WdListGalleryType.BULLET,
    "bullets": WdListGalleryType.BULLET,
    "numbered": WdListGalleryType.NUMBER,
    "number": WdListGalleryType.NUMBER,
    "numbers": WdListGalleryType.NUMBER,
    "outline": WdListGalleryType.OUTLINE_NUMBER,
    "outline-number": WdListGalleryType.OUTLINE_NUMBER,
    "outline_number": WdListGalleryType.OUTLINE_NUMBER,
}

_CANONICAL_TYPES = ("bulleted", "numbered", "outline")

# Per-level number-style names (apply_list_format) -> WdListNumberStyle. Bullet
# is deliberately absent: a bullet level is authored via NumberFormat + a symbol
# font, never NumberStyle=23 (which raises on a multi-level template).
_NUMBER_STYLE_FOR: dict[str, WdListNumberStyle] = {
    "arabic": WdListNumberStyle.ARABIC,
    "decimal": WdListNumberStyle.ARABIC,
    "upper-roman": WdListNumberStyle.UPPERCASE_ROMAN,
    "uppercase-roman": WdListNumberStyle.UPPERCASE_ROMAN,
    "lower-roman": WdListNumberStyle.LOWERCASE_ROMAN,
    "lowercase-roman": WdListNumberStyle.LOWERCASE_ROMAN,
    "upper-letter": WdListNumberStyle.UPPERCASE_LETTER,
    "uppercase-letter": WdListNumberStyle.UPPERCASE_LETTER,
    "lower-letter": WdListNumberStyle.LOWERCASE_LETTER,
    "lowercase-letter": WdListNumberStyle.LOWERCASE_LETTER,
    "ordinal": WdListNumberStyle.ORDINAL,
    "cardinal-text": WdListNumberStyle.CARDINAL_TEXT,
    "ordinal-text": WdListNumberStyle.ORDINAL_TEXT,
}
_NUMBER_STYLE_NAME: dict[int, str] = {
    int(WdListNumberStyle.ARABIC): "arabic",
    int(WdListNumberStyle.UPPERCASE_ROMAN): "upper-roman",
    int(WdListNumberStyle.LOWERCASE_ROMAN): "lower-roman",
    int(WdListNumberStyle.UPPERCASE_LETTER): "upper-letter",
    int(WdListNumberStyle.LOWERCASE_LETTER): "lower-letter",
    int(WdListNumberStyle.ORDINAL): "ordinal",
    int(WdListNumberStyle.CARDINAL_TEXT): "cardinal-text",
    int(WdListNumberStyle.ORDINAL_TEXT): "ordinal-text",
    int(WdListNumberStyle.BULLET): "bullet",
}
_TRAILING_FOR: dict[str, WdTrailingCharacter] = {
    "tab": WdTrailingCharacter.TAB,
    "space": WdTrailingCharacter.SPACE,
    "none": WdTrailingCharacter.NONE,
}
_TRAILING_NAME: dict[int, str] = {int(v): k for k, v in _TRAILING_FOR.items()}
_LEVEL_ALIGN_FOR: dict[str, WdListLevelAlignment] = {
    "left": WdListLevelAlignment.LEFT,
    "center": WdListLevelAlignment.CENTER,
    "centre": WdListLevelAlignment.CENTER,
    "right": WdListLevelAlignment.RIGHT,
}

# ListFormat.ListType int -> the human string list_info() reports.
_LIST_TYPE_NAMES: dict[int, str] = {
    int(WdListType.NO_NUMBERING): "none",
    int(WdListType.LIST_NUM_ONLY): "number-only",
    int(WdListType.BULLET): "bulleted",
    int(WdListType.SIMPLE_NUMBERING): "numbered",
    int(WdListType.OUTLINE_NUMBERING): "outline",
    int(WdListType.MIXED_NUMBERING): "mixed",
}


def gallery_for(list_type: str) -> WdListGalleryType:
    """Resolve a `list_type` string to its `WdListGalleryType`.

    Raises `ValueError` for an unknown name — symmetric with how
    `_coerce_alignment` rejects bad alignment strings.
    """
    try:
        return _GALLERY_FOR[str(list_type).lower()]
    except KeyError:
        raise ValueError(
            f"unknown list type {list_type!r}; expected one of {list(_CANONICAL_TYPES)}"
        ) from None


def _read(lf: Any, attr: str, default: Any) -> Any:
    """Read a `ListFormat` attribute, tolerating non-list ranges that raise."""
    try:
        value = getattr(lf, attr)
    except Exception:
        return default
    return default if value is None else value


def apply_list_template(
    rng: Any, gallery_type: WdListGalleryType, *, continue_previous: bool
) -> None:
    """Apply gallery `gallery_type`'s first template to `rng`'s paragraphs."""
    app = rng.Application
    gallery = app.ListGalleries(int(gallery_type))
    template = gallery.ListTemplates(1)
    rng.ListFormat.ApplyListTemplate(
        ListTemplate=template,
        ContinuePreviousList=bool(continue_previous),
        ApplyTo=int(WdListApplyTo.WHOLE_LIST),
        DefaultListBehavior=int(WdDefaultListBehavior.WORD10),
    )


def restart_numbering(rng: Any) -> None:
    """Re-apply the range's current list template starting at 1.

    Raises `ValueError` if the range isn't part of a list (no template to
    restart).
    """
    lf = rng.ListFormat
    template = _read(lf, "ListTemplate", None)
    if template is None:
        raise ValueError("range is not part of a list; nothing to restart")
    lf.ApplyListTemplate(
        ListTemplate=template,
        ContinuePreviousList=False,
        ApplyTo=int(WdListApplyTo.WHOLE_LIST),
        DefaultListBehavior=int(WdDefaultListBehavior.WORD10),
    )


def _configure_level(lvl: Any, index: int, spec: dict[str, Any]) -> None:
    """Write one per-level spec onto a `ListLevel` (the unit `apply_list_format`
    fans across)."""
    kind = str(spec.get("kind", "number")).lower()
    font_name = spec.get("font")
    if kind == "bullet":
        glyph = spec.get("bullet") or spec.get("format")
        if not glyph:
            raise OpError(f"list level {index}: a bullet level needs a 'bullet' glyph")
        # A bullet is the glyph as the number format + a symbol font — NOT
        # NumberStyle=bullet, which raises on a multi-level template.
        lvl.NumberFormat = str(glyph)
        lvl.Font.Name = font_name or "Symbol"
    elif kind == "number":
        style = spec.get("style")
        if style is not None:
            try:
                lvl.NumberStyle = int(_NUMBER_STYLE_FOR[str(style).lower()])
            except KeyError:
                raise OpError(
                    f"list level {index}: unknown number style {style!r}; "
                    f"expected one of {sorted(_NUMBER_STYLE_FOR)}"
                ) from None
        lvl.NumberFormat = str(spec.get("format", f"%{index}."))
        if "start_at" in spec:
            lvl.StartAt = int(spec["start_at"])
        if font_name:
            lvl.Font.Name = str(font_name)
    else:
        raise OpError(f"list level {index}: kind must be 'number' or 'bullet', got {kind!r}")

    trailing = spec.get("trailing")
    if trailing is not None:
        try:
            lvl.TrailingCharacter = int(_TRAILING_FOR[str(trailing).lower()])
        except KeyError:
            raise OpError(
                f"list level {index}: trailing must be tab/space/none, got {trailing!r}"
            ) from None
    if spec.get("number_position") is not None:
        lvl.NumberPosition = to_points(spec["number_position"])
    if spec.get("text_position") is not None:
        lvl.TextPosition = to_points(spec["text_position"])
    if spec.get("alignment") is not None:
        try:
            lvl.Alignment = int(_LEVEL_ALIGN_FOR[str(spec["alignment"]).lower()])
        except KeyError:
            raise OpError(
                f"list level {index}: alignment must be left/center/right, "
                f"got {spec['alignment']!r}"
            ) from None
    if spec.get("bold") is not None:
        lvl.Font.Bold = bool(spec["bold"])
    if spec.get("italic") is not None:
        lvl.Font.Italic = bool(spec["italic"])
    if spec.get("color") is not None:
        lvl.Font.Color = to_bgr(spec["color"])


def apply_list_format(
    doc_com: Any, rng: Any, levels: list[dict[str, Any]], *, continue_previous: bool = False
) -> None:
    """Author a custom list template from per-level specs and apply it to `rng`.

    `levels` is a 1-based list of per-level dicts (see `Anchor.apply_list_format`).
    A multi-level `levels` mints an outline template (9 levels — extra levels keep
    Word's defaults); a single level mints a simple one. Raises `OpError` on a bad
    spec.
    """
    if not levels:
        raise OpError("apply_list_format: at least one level is required")
    outline = len(levels) > 1
    # ListTemplates.Add(OutlineNumbered) — Name omitted so Word auto-names and
    # repeated calls don't collide. Each call defines a new list in the document.
    lt = doc_com.ListTemplates.Add(outline)
    for i, spec in enumerate(levels, start=1):
        if not isinstance(spec, dict):
            raise OpError(
                f"list level {i}: each level must be an object, got {type(spec).__name__}"
            )
        _configure_level(lt.ListLevels(i), i, spec)
    rng.ListFormat.ApplyListTemplate(
        ListTemplate=lt,
        ContinuePreviousList=bool(continue_previous),
        ApplyTo=int(WdListApplyTo.WHOLE_LIST),
        DefaultListBehavior=int(WdDefaultListBehavior.WORD10),
    )


def read_list_levels(rng: Any) -> list[dict[str, Any]]:
    """The per-level format of the list `rng` sits in (its `ListTemplate`).

    Returns one dict per level — `{level, kind, format, style, trailing,
    number_position, text_position, font}` — or `[]` if the range carries no
    list. The read mirror of `apply_list_format`.
    """
    lf = rng.ListFormat
    template = _read(lf, "ListTemplate", None)
    if template is None:
        return []
    out: list[dict[str, Any]] = []
    levels = template.ListLevels
    for i in range(1, int(levels.Count) + 1):
        lvl = levels(i)
        style_int = int(_read(lvl, "NumberStyle", 0))
        fmt = str(_read(lvl, "NumberFormat", "") or "")
        # A number level's format always carries a %N placeholder; a bullet is a
        # bare glyph (and NumberStyle stays 0 when authored via glyph + font, so
        # the placeholder test is the reliable discriminator).
        is_bullet = style_int == int(WdListNumberStyle.BULLET) or (bool(fmt) and "%" not in fmt)
        out.append(
            {
                "level": i,
                "kind": "bullet" if is_bullet else "number",
                "format": fmt,
                "style": "bullet"
                if is_bullet
                else _NUMBER_STYLE_NAME.get(style_int, str(style_int)),
                "trailing": _TRAILING_NAME.get(int(_read(lvl, "TrailingCharacter", 0)), "tab"),
                "number_position": float(_read(lvl, "NumberPosition", 0.0) or 0.0),
                "text_position": float(_read(lvl, "TextPosition", 0.0) or 0.0),
                "font": str(_read(lvl.Font, "Name", "") or ""),
            }
        )
    return out


def read_list_info(rng: Any) -> dict[str, Any]:
    """Describe the list a range sits in: `{type, level, number, string}`.

    `type` is `"none"` when the range carries no list formatting. `number` is
    the value of the first paragraph's number (0 for bullets), and `string` is
    its rendered marker (`"1."`, `"a)"`, `"•"`, …).
    """
    lf = rng.ListFormat
    try:
        list_type = int(_read(lf, "ListType", 0))
    except (TypeError, ValueError):
        list_type = 0
    try:
        level = int(_read(lf, "ListLevelNumber", 0))
    except (TypeError, ValueError):
        level = 0
    try:
        number = int(_read(lf, "ListValue", 0))
    except (TypeError, ValueError):
        number = 0
    return {
        "type": _LIST_TYPE_NAMES.get(list_type, "unknown"),
        "level": level,
        "number": number,
        "string": str(_read(lf, "ListString", "") or ""),
    }


class ListCollection:
    """Read-only, iterable view over the document's lists (`doc.lists`).

    Index a list by 1-based position (`doc.lists[2]`) to get a
    [`RangeAnchor`][wordlive.RangeAnchor] over its whole range — so every list
    verb (`apply_list`, `restart_numbering`, …) is immediately available on it.
    `list()` returns a summary per list; positions match Word's own
    `Document.Lists(n)` ordering.
    """

    def __init__(self, doc: Document) -> None:
        self._doc = doc

    def __len__(self) -> int:
        with _com.translate_com_errors():
            return int(self._doc.com.Lists.Count)

    def _spans(self) -> list[tuple[int, int]]:
        with _com.translate_com_errors():
            count = int(self._doc.com.Lists.Count)
            spans: list[tuple[int, int]] = []
            for i in range(1, count + 1):
                rng = self._doc.com.Lists(i).Range
                spans.append((int(rng.Start), int(rng.End)))
        return spans

    def __getitem__(self, index: int) -> RangeAnchor:
        if isinstance(index, bool) or not isinstance(index, int):
            raise TypeError(f"list index must be int, got {type(index).__name__}")
        n = len(self)
        if not (1 <= index <= n):
            raise AnchorNotFoundError("list", str(index))
        with _com.translate_com_errors():
            rng = self._doc.com.Lists(index).Range
            start, end = int(rng.Start), int(rng.End)
        return self._doc.range(start, end)

    def __iter__(self) -> Iterator[RangeAnchor]:
        for start, end in self._spans():
            yield self._doc.range(start, end)

    def list(self) -> list[dict[str, Any]]:
        """All lists as `{index, type, count, anchor_id}` dicts."""
        out: list[dict[str, Any]] = []
        with _com.translate_com_errors():
            count = int(self._doc.com.Lists.Count)
            for i in range(1, count + 1):
                lst = self._doc.com.Lists(i)
                rng = lst.Range
                start, end = int(rng.Start), int(rng.End)
                info = read_list_info(rng)
                try:
                    n_items = int(lst.ListParagraphs.Count)
                except Exception:
                    n_items = 0
                out.append(
                    {
                        "index": i,
                        "type": info["type"],
                        "count": n_items,
                        "anchor_id": f"range:{start}-{end}",
                    }
                )
        return out
