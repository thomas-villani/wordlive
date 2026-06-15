"""Table of authorities — TA citation marks plus a `TOA` field block.

The legal sibling of the back-of-book index: building a table of authorities is
two steps, and wordlive mirrors them:

1. **Mark citations.** `Anchor.mark_citation(long, short=…, category=…)` plants a
   hidden ``TA`` field on the anchor's range — `long` is the full citation as it
   appears in the table, `short` the abbreviated form used elsewhere, `category`
   the section it files under (Cases / Statutes / …). This is the per-authority
   step.
2. **Build the table.** `Anchor.insert_table_of_authorities(...)` /
   `Document.add_table_of_authorities(...)` inserts the ``TOA`` field that gathers
   the marked citations of one (or every) category into a page-numbered table, and
   returns a `TableOfAuthorities`.

Like a TOC the table is a *field block* — refresh it with `update()` or
`Document.update_fields()`. (Word's `TableOfAuthorities` has no
`UpdatePageNumbers`, unlike a TOC/table-of-figures, so neither does this wrapper.)
"""

from __future__ import annotations

from typing import TYPE_CHECKING, Any

from . import _com
from ._anchors import range_text

if TYPE_CHECKING:
    from ._document import Document

# Friendly category keyword -> Word's built-in TOA category index. 0 = all
# categories; 1-7 are Word's defaults (8-16 are user-customisable slots).
_TOA_CATEGORIES: dict[str, int] = {
    "all": 0,
    "cases": 1,
    "statutes": 2,
    "other": 3,
    "other_authorities": 3,
    "rules": 4,
    "treatises": 5,
    "regulations": 6,
    "constitutional": 7,
    "constitutional_provisions": 7,
}


class TableOfAuthorities:
    """A table of authorities created by `insert_table_of_authorities`."""

    def __init__(self, doc: Document, com: Any) -> None:
        self._doc = doc
        self._com = com

    @property
    def com(self) -> Any:
        """Raw COM `TableOfAuthorities` object — the escape hatch."""
        return self._com

    @property
    def range(self) -> Any:
        """The COM `Range` the table occupies."""
        with _com.translate_com_errors():
            return self._com.Range

    @property
    def text(self) -> str:
        """The table's rendered text (entries + page numbers, once updated)."""
        with _com.translate_com_errors():
            return range_text(self._com.Range)

    def update(self) -> None:
        """Rebuild the table from the marked ``TA`` citations.

        Call this after adding or moving citation marks — or use
        [`Document.update_fields`][wordlive.Document.update_fields] to refresh it
        with every other field. Wrap in `doc.edit(...)` for atomic undo.
        """
        with _com.translate_com_errors():
            self._com.Update()

    def __repr__(self) -> str:
        return "<TableOfAuthorities>"
