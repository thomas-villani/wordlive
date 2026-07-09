"""Apply, read, and re-level list formatting on an anchor's paragraphs."""

from __future__ import annotations

from typing import TYPE_CHECKING, Any

from .. import _com, _lists
from ..constants import (
    WdNumberType,
)

if TYPE_CHECKING:
    pass


if TYPE_CHECKING:
    pass

from ._anchor_core import AnchorCore


class AnchorListsMixin(AnchorCore):
    """Apply, read, and re-level list formatting on an anchor's paragraphs."""

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

    def apply_list_format(
        self, levels: list[dict[str, Any]], *, continue_previous: bool = False
    ) -> None:
        """Author a **custom** multi-level list template and apply it here.

        The richer counterpart to `apply_list` (which only applies a gallery
        default): `levels` is a 1-based list of per-level specs that defines the
        marker, indentation, and marker font of each list level. Each spec is a
        dict; all keys are optional except a bullet level's glyph:

        - `kind` — `"number"` (default) or `"bullet"`.
        - `format` — for a number level, the marker template (`"%1."`, `"%1)"`,
          `"%1.%2"`; `%N` references level N's number), default `"%{level}."`;
          for a bullet level, the glyph (or pass `bullet`).
        - `style` — a number level's scheme: `"arabic"`, `"upper-roman"`,
          `"lower-roman"`, `"upper-letter"`, `"lower-letter"`, `"ordinal"`, … .
        - `bullet` / `font` — a bullet level's glyph and marker font (default
          `"Symbol"`); `font` also sets a number level's marker font.
        - `start_at` — a number level's first value.
        - `number_position` / `text_position` — the marker and text indents
          (points or a length string like `"0.5in"`).
        - `trailing` — what follows the marker: `"tab"` / `"space"` / `"none"`.
        - `alignment` — the marker's alignment: `"left"` / `"center"` / `"right"`.
        - `bold` / `italic` / `color` — the marker font's styling.

        More than one level mints an outline template (levels beyond those given
        keep Word's defaults). `read_list_levels()` is the read mirror. Wrap in
        `doc.edit(...)` for atomic undo; a bad spec raises `OpError`.
        """
        with _com.translate_com_errors():
            _lists.apply_list_format(
                self._doc.com, self._range(), levels, continue_previous=continue_previous
            )

    def read_list_levels(self) -> list[dict[str, Any]]:
        """The per-level format of the list this anchor sits in — a pure read.

        Returns one `{level, kind, format, number_style, style, trailing,
        number_position, text_position, font}` dict per level of the applied
        `ListTemplate`, or `[]` if the anchor carries no list (`number_style` is
        the raw `WdListNumberStyle` int). The read mirror of `apply_list_format`.
        """
        with _com.translate_com_errors():
            return _lists.read_list_levels(self._range())

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
