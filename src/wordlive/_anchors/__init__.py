"""Semantic anchors — the stable, LLM-visible addressing scheme.

`Anchor` (in `_base`) is the abstract base; each concrete anchor and its
collection live in a sibling module. Modules named `*_anchors` are named that
way to stay distinct from the same-named feature modules one level up
(`wordlive._charts`, `._images`, `._shapes`, `._equations`).
"""

from __future__ import annotations

from ._base import (
    Anchor,
)
from ._bookmarks import (
    Bookmark,
    BookmarkCollection,
    _bookmarks_including_hidden,
)
from ._chart_anchors import (
    ChartAnchor,
    ChartCollection,
)
from ._content_controls import (
    ContentControl,
    ContentControlCollection,
)
from ._equation_anchors import (
    EquationAnchor,
    EquationCollection,
)
from ._headings import (
    Heading,
    HeadingCollection,
    _IndexedHeading,
)
from ._helpers import (
    _apply_font,
    _apply_paragraph_format,
    _coerce_line_spacing,
    _coerce_named,
    _line_spacing_repr,
    _markdown_segments,
    _normalize_table_data,
    _read_font,
    _read_paragraph_format,
    _utf16_len,
    _within_table,
    apply_borders,
    paragraph_text,
    range_text,
)
from ._image_anchors import (
    ImageAnchor,
    ImageCollection,
)
from ._paragraphs import (
    Paragraph,
    ParagraphCollection,
)
from ._range import (
    EndAnchor,
    RangeAnchor,
    StartAnchor,
)
from ._refs import (
    _WL_PREFIX,
    _mint_wl_bookmark,
    _new_pin_code,
    _pin_id_for,
    _pin_name_for,
    _validate_pin_slug,
)
from ._shape_anchors import (
    ShapeAnchor,
    ShapeCollection,
    TextBoxCollection,
)

__all__ = [
    "Anchor",
    "Bookmark",
    "BookmarkCollection",
    "ChartAnchor",
    "ChartCollection",
    "ContentControl",
    "ContentControlCollection",
    "EndAnchor",
    "EquationAnchor",
    "EquationCollection",
    "Heading",
    "HeadingCollection",
    "ImageAnchor",
    "ImageCollection",
    "Paragraph",
    "ParagraphCollection",
    "RangeAnchor",
    "ShapeAnchor",
    "ShapeCollection",
    "StartAnchor",
    "TextBoxCollection",
    "_IndexedHeading",
    "_WL_PREFIX",
    "_apply_font",
    "_apply_paragraph_format",
    "_bookmarks_including_hidden",
    "_coerce_line_spacing",
    "_coerce_named",
    "_line_spacing_repr",
    "_markdown_segments",
    "_mint_wl_bookmark",
    "_new_pin_code",
    "_normalize_table_data",
    "_pin_id_for",
    "_pin_name_for",
    "_read_font",
    "_read_paragraph_format",
    "_utf16_len",
    "_validate_pin_slug",
    "_within_table",
    "apply_borders",
    "paragraph_text",
    "range_text",
]
