"""Table of contents — a field block built from the document's headings.

`Anchor.insert_toc(...)` / `Document.add_toc(...)` create a TOC (over Word's
`TablesOfContents`) and return a `Toc` wrapper. A TOC is a *field block*, not a
single addressable range, so it has no `anchor_id` scheme; refresh it through
the returned object's `update()` / `update_page_numbers()`, or — alongside every
other field — with `Document.update_fields()`.

Page numbers only populate after repagination: call `update()` (or take a
`snapshot`, which forces print layout) before reading them.
"""

from __future__ import annotations

from typing import TYPE_CHECKING, Any

from . import _com
from ._anchors import range_text

if TYPE_CHECKING:
    from ._document import Document


class Toc:
    """A table of contents created by `insert_toc` / `add_toc`."""

    def __init__(self, doc: Document, com: Any) -> None:
        self._doc = doc
        self._com = com

    @property
    def com(self) -> Any:
        """Raw COM `TableOfContents` object — the escape hatch."""
        return self._com

    @property
    def range(self) -> Any:
        """The COM `Range` the TOC occupies."""
        with _com.translate_com_errors():
            return self._com.Range

    @property
    def text(self) -> str:
        """The TOC's rendered text (entries + page numbers, once updated)."""
        with _com.translate_com_errors():
            return range_text(self._com.Range)

    def update(self) -> None:
        """Rebuild the TOC's entries and page numbers from the current document.

        Call this after edits that change headings or pagination — or use
        [`Document.update_fields`][wordlive.Document.update_fields] to refresh
        the TOC together with every other field. Wrap in `doc.edit(...)` for
        atomic undo.
        """
        with _com.translate_com_errors():
            self._com.Update()

    def update_page_numbers(self) -> None:
        """Refresh only the TOC's page numbers (cheaper than a full `update()`)."""
        with _com.translate_com_errors():
            self._com.UpdatePageNumbers()

    def __repr__(self) -> str:
        return "<Toc>"


class TableOfFigures:
    """A table of figures created by `insert_table_of_figures`.

    The caption-driven sibling of [`Toc`][wordlive.Toc]: it lists every caption
    of one label (``"Figure"`` / ``"Table"`` / ``"Equation"`` / a custom label)
    with its page number, built over Word's `TablesOfFigures`. Like a TOC it is a
    *field block*, not a single addressable range — refresh it with `update()` /
    `update_page_numbers()`, or with `Document.update_fields()`. Page numbers
    populate only after repagination.
    """

    def __init__(self, doc: Document, com: Any) -> None:
        self._doc = doc
        self._com = com

    @property
    def com(self) -> Any:
        """Raw COM `TableOfFigures` object — the escape hatch."""
        return self._com

    @property
    def range(self) -> Any:
        """The COM `Range` the table of figures occupies."""
        with _com.translate_com_errors():
            return self._com.Range

    @property
    def text(self) -> str:
        """The rendered text (entries + page numbers, once updated)."""
        with _com.translate_com_errors():
            return range_text(self._com.Range)

    def update(self) -> None:
        """Rebuild entries and page numbers from the current document's captions."""
        with _com.translate_com_errors():
            self._com.Update()

    def update_page_numbers(self) -> None:
        """Refresh only the page numbers (cheaper than a full `update()`)."""
        with _com.translate_com_errors():
            self._com.UpdatePageNumbers()

    def __repr__(self) -> str:
        return "<TableOfFigures>"
