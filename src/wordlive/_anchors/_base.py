"""The `Anchor` abstract base: everything a concrete anchor can do to its range.

`Anchor` is assembled from feature mixins that all inherit `AnchorCore`; see
`_anchor_core.py` for the spine and the `_anchor_*.py` siblings for the rest.
"""

from __future__ import annotations

from ._anchor_format import AnchorFormatMixin
from ._anchor_insert import AnchorInsertMixin
from ._anchor_lists import AnchorListsMixin
from ._anchor_media import AnchorMediaMixin
from ._anchor_read import AnchorReadMixin
from ._anchor_references import AnchorReferencesMixin

__all__ = ["Anchor"]


class Anchor(
    AnchorInsertMixin,
    AnchorMediaMixin,
    AnchorReferencesMixin,
    AnchorFormatMixin,
    AnchorListsMixin,
    AnchorReadMixin,
):
    """Abstract base — subclasses know how to materialise their COM Range.

    Concrete subclasses must implement `_range()` and `set_text()`. Other
    operations (`text`, `insert_before`, `insert_after`, `delete`,
    `apply_style`, `format_paragraph`) are derived and inherited as-is.
    """
