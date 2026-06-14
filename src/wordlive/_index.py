"""Back-of-book index — XE entry marks plus an `INDEX` field block.

Building an index is two steps in Word, and wordlive mirrors them:

1. **Mark entries.** `Anchor.mark_index_entry(entry)` plants a hidden `XE`
   (index-entry) field on the anchor's range — `entry` is the text that appears
   in the index (use ``"main:sub"`` for a subentry). This is the per-term step;
   call it once per thing you want indexed.
2. **Build the index.** `Anchor.insert_index(...)` / `Document.add_index(...)`
   inserts the `INDEX` field that gathers every marked entry into an
   alphabetised, page-numbered list, and returns an `Index` wrapper.

Like a table of contents, the index is a *field block*, not a single
addressable range — refresh it through the returned object's `update()`, or —
alongside every other field — with `Document.update_fields()`. Page numbers
populate only after repagination.
"""

from __future__ import annotations

from typing import TYPE_CHECKING, Any

from . import _com
from ._anchors import range_text

if TYPE_CHECKING:
    from ._document import Document


class Index:
    """A back-of-book index created by `insert_index` / `add_index`."""

    def __init__(self, doc: Document, com: Any) -> None:
        self._doc = doc
        self._com = com

    @property
    def com(self) -> Any:
        """Raw COM `Index` object — the escape hatch."""
        return self._com

    @property
    def range(self) -> Any:
        """The COM `Range` the index occupies."""
        with _com.translate_com_errors():
            return self._com.Range

    @property
    def text(self) -> str:
        """The index's rendered text (entries + page numbers, once updated)."""
        with _com.translate_com_errors():
            return range_text(self._com.Range)

    def update(self) -> None:
        """Rebuild the index's entries and page numbers from the marked entries.

        Call this after adding or moving `XE` marks — or use
        [`Document.update_fields`][wordlive.Document.update_fields] to refresh the
        index together with every other field. Wrap in `doc.edit(...)` for atomic
        undo.
        """
        with _com.translate_com_errors():
            self._com.Update()

    def __repr__(self) -> str:
        return "<Index>"
