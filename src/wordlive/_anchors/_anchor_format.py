"""Apply and read character/paragraph formatting, shading, borders, tab stops."""

from __future__ import annotations

from collections.abc import Iterator
from contextlib import contextmanager
from contextvars import ContextVar
from typing import TYPE_CHECKING, Any

from .. import _com
from .._format import to_bgr, to_points
from ..constants import (
    WdDropPosition,
    WdTabLeader,
)
from ..exceptions import OpError

if TYPE_CHECKING:
    pass

from ._helpers import (
    _DROP_POSITIONS,
    _TAB_ALIGN,
    _TAB_LEADERS,
    _apply_font,
    _apply_paragraph_format,
    _coerce_highlight,
    _coerce_named,
    _read_font,
    _read_highlight,
    _read_paragraph_format,
    apply_borders,
)

if TYPE_CHECKING:
    pass

from ._anchor_core import AnchorCore

# A style's own paragraph/font values — the baseline `format_info` diffs each range
# against to decide what is a direct override. Reading it costs ~25 COM properties
# and two object wraps, yet it depends only on the style, so a bulk reader (the
# linter, walking hundreds of same-styled paragraphs) can memoise it. A Word style
# is document-global, so the style's name is a sufficient key.
_StyleBaseline = tuple[dict[str, Any], dict[str, Any]]
_BASELINE_CACHE: ContextVar[dict[str, _StyleBaseline] | None] = ContextVar(
    "wordlive_style_baseline", default=None
)


@contextmanager
def style_baseline_cache() -> Iterator[None]:
    """Memoise style baselines for the duration of one bulk read.

    Only safe while nothing edits a *style definition*; direct formatting on a range
    is read fresh either way. Scoped to the `with` block, never to a `Document`.
    """
    token = _BASELINE_CACHE.set({})
    try:
        yield
    finally:
        _BASELINE_CACHE.reset(token)


def _style_baseline(style: Any, style_name: str) -> _StyleBaseline:
    cache = _BASELINE_CACHE.get()
    if cache is not None and style_name in cache:
        return cache[style_name]
    baseline: _StyleBaseline = (
        _read_paragraph_format(style.ParagraphFormat),
        _read_font(style.Font)[0],
    )
    if cache is not None:
        cache[style_name] = baseline
    return baseline


class AnchorFormatMixin(AnchorCore):
    """Apply and read character/paragraph formatting, shading, borders, tab stops."""

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
        use the same keywords the write verbs accept. `font.hidden` flags Word's
        hidden-text attribute. `font.highlight` is a highlight keyword (`"yellow"`,
        … or `"none"`); it lives on the range, not the style, so it's
        effective-only — `style` is always `null` and `override` just means a
        highlight is present.

        The field vocabulary is identical to the write side, so a value read here
        can be written straight back through `format_paragraph`/`format_run`.
        """
        with _com.translate_com_errors():
            rng = self._range()
            style = rng.ParagraphStyle
            style_name = str(style.NameLocal)
            eff_para = _read_paragraph_format(rng.ParagraphFormat)
            eff_font, mixed = _read_font(rng.Font)
            highlight = _read_highlight(rng.HighlightColorIndex)
            sty_para, sty_font = _style_baseline(style, style_name)

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
        # Highlight lives on the Range, not the Font, and a style never carries it
        # (see `_STYLE_RUN_FIELDS`), so it's effective-only: no style baseline, and
        # an "override" simply means a highlight is present. A mixed read (some runs
        # highlighted) surfaces via `mixed`, like the other character fields.
        if highlight is None:
            mixed.append("highlight")
        font["mixed"] = mixed
        font["highlight"] = {
            "value": highlight,
            "style": None,
            "override": highlight is not None and highlight != "none",
        }
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

        This sets per-range / per-cell borders. Page borders
        (`Section.Borders`) are out of scope; whole-table borders (the entire
        grid in one call, including interior gridlines) go through
        [`Table.set_borders`][wordlive.Table.set_borders] / the `table
        set-borders` verb. Bad input raises `OpError`. Wrap in `doc.edit(...)`.
        """
        try:
            with _com.translate_com_errors():
                apply_borders(
                    self._range().Borders,
                    sides=sides,
                    style=style,
                    weight=weight,
                    color=color,
                )
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
