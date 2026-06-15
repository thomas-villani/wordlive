"""In-text citations and the generated bibliography.

Both are Word *fields* over a source store (`doc.sources`):

- `Anchor.insert_citation(tag, ...)` inserts a ``CITATION`` field that references
  a source by tag and renders it in the document's bibliography style (e.g.
  *(Smith 2020, 15)*), returning a `Citation`.
- `Anchor.insert_bibliography()` / `Document.add_bibliography()` inserts a
  ``BIBLIOGRAPHY`` field — the reference list of every *cited* source — returning
  a `Bibliography`.

Both wrap a COM `Field`: `text` is the rendered result and `update()` re-renders
it (also done wholesale by `Document.update_fields()`). Note `wordlive.Bibliography`
(the inserted field) is distinct from COM's `doc.Bibliography`, the sources/style
manager wordlive surfaces as `doc.sources` and `doc.bibliography_style`.
"""

from __future__ import annotations

from typing import TYPE_CHECKING, Any

from . import _com
from ._anchors import range_text

if TYPE_CHECKING:
    from ._document import Document


class Citation:
    """An in-text citation field created by `insert_citation`."""

    def __init__(self, doc: Document, com: Any) -> None:
        self._doc = doc
        self._com = com

    @property
    def com(self) -> Any:
        """Raw COM `Field` object — the escape hatch."""
        return self._com

    @property
    def range(self) -> Any:
        """The COM `Range` the citation's rendered text occupies."""
        with _com.translate_com_errors():
            return self._com.Result

    @property
    def text(self) -> str:
        """The rendered citation (e.g. ``" (Smith 2020)"``)."""
        with _com.translate_com_errors():
            return range_text(self._com.Result)

    @property
    def tag(self) -> str:
        """The source tag this citation references (parsed from its field code)."""
        with _com.translate_com_errors():
            code = str(self._com.Code.Text).strip()
        parts = code.split()
        return parts[1] if len(parts) >= 2 and parts[0].upper() == "CITATION" else ""

    def update(self) -> None:
        """Re-render the citation (e.g. after changing the bibliography style)."""
        with _com.translate_com_errors():
            self._com.Update()

    def __repr__(self) -> str:
        return f"<Citation tag={self.tag!r}>"


class Bibliography:
    """A generated bibliography field created by `insert_bibliography` / `add_bibliography`."""

    def __init__(self, doc: Document, com: Any) -> None:
        self._doc = doc
        self._com = com

    @property
    def com(self) -> Any:
        """Raw COM `Field` object — the escape hatch."""
        return self._com

    @property
    def range(self) -> Any:
        """The COM `Range` the bibliography occupies."""
        with _com.translate_com_errors():
            return self._com.Result

    @property
    def text(self) -> str:
        """The rendered reference list (once updated)."""
        with _com.translate_com_errors():
            return range_text(self._com.Result)

    def update(self) -> None:
        """Rebuild the reference list from the cited sources.

        Call this after adding citations or changing
        [`Document.bibliography_style`][wordlive.Document.bibliography_style] — or
        use [`Document.update_fields`][wordlive.Document.update_fields] to refresh
        it with every other field. Wrap in `doc.edit(...)` for atomic undo.
        """
        with _com.translate_com_errors():
            self._com.Update()

    def __repr__(self) -> str:
        return "<Bibliography>"
