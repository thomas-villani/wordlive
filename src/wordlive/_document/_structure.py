"""Structural inserts: tables, TOC/index/bibliography/TOA, shape grouping."""

from __future__ import annotations

import secrets
from typing import TYPE_CHECKING, Any

from .. import _com, _shapes
from .._anchors import (
    ShapeAnchor,
)
from .._tables import Table
from ..exceptions import (
    OpError,
)

if TYPE_CHECKING:
    pass

from ._core import DocumentCore


class StructureMixin(DocumentCore):
    """Structural inserts: tables, TOC/index/bibliography/TOA, shape grouping."""

    def group_shapes(self, *anchor_ids: str) -> ShapeAnchor:
        """Group two or more floating shapes into a single group shape.

        Each `anchor_id` is a `shape:N` (the members must all be floating shapes).
        Returns the new group's [`ShapeAnchor`][wordlive.ShapeAnchor] (`shape:N`,
        `shape_type == "group"`) — move / size / delete it as one unit, or
        [`ungroup`][wordlive.ShapeAnchor.ungroup] it to get the members back. Word
        requires members to allow overlap, so this enables that first. Wrap in
        `doc.edit(...)` for atomic undo. Raises `OpError` for fewer than two
        shapes or a non-shape anchor.
        """
        if len(anchor_ids) < 2:
            raise OpError("group_shapes needs at least two shape anchors")
        with _com.translate_com_errors():
            coms: list[Any] = []
            for aid in anchor_ids:
                anchor = self.anchor_by_id(aid)
                if not isinstance(anchor, ShapeAnchor):
                    raise OpError(f"{aid!r} is not a shape; group_shapes needs shape:N anchors")
                coms.append(anchor._shape())
            group = _shapes.group_shapes(self.com, coms)
            # Locate the new group by a unique temp name — don't assume "last".
            orig_name = str(group.Name or "")
            probe_name = f"_wl_shape_{secrets.token_hex(8)}"
            group.Name = probe_name
            index = _shapes.index_of_named(self.com, probe_name)
            # Restore unconditionally so an empty original name doesn't leave the
            # `_wl_shape_*` probe lingering in list().
            group.Name = orig_name
        return ShapeAnchor(self._as_document, index)

    def add_table(
        self,
        rows: int,
        cols: int,
        *,
        style: str | None = None,
        data: list[list[Any]] | None = None,
        header: bool = False,
    ) -> Table:
        """Append a `rows` × `cols` table at the end of the document and return it.

        The "build a document from the bottom up" helper for tables — the
        counterpart to [`append_paragraph`][wordlive.Document.append_paragraph].
        Sugar for `self.end.insert_table(...)`; see
        [`Anchor.insert_table`][wordlive.Anchor.insert_table] for the full
        semantics of `style` (defaults to the built-in ``"Table Grid"``), `data`
        (row-major fill, validated up front), and `header`. To place a table
        somewhere other than the end, resolve a position anchor and call
        `insert_table` on it directly (e.g.
        `doc.headings["Pricing"].insert_table(3, 2, ...)`). Wrap in
        `doc.edit(...)` for atomic undo.
        """
        return self.end.insert_table(
            rows, cols, where="after", style=style, data=data, header=header
        )

    def add_toc(
        self,
        *,
        levels: tuple[int, int] = (1, 3),
        use_heading_styles: bool = True,
        hyperlinks: bool = True,
    ) -> Any:
        """Insert a table of contents at the very start of the document.

        The "documents want their TOC at the top" helper — sugar for
        `self.start.insert_toc(...)`. See
        [`Anchor.insert_toc`][wordlive.Anchor.insert_toc] for the full semantics
        of `levels` (a ``(upper, lower)`` heading-level pair), `use_heading_styles`,
        and `hyperlinks`, and for the page-number-repagination caveat. Returns
        the new [`Toc`][wordlive.Toc]. Wrap in `doc.edit(...)` for atomic undo.
        """
        return self.start.insert_toc(
            levels=levels, use_heading_styles=use_heading_styles, hyperlinks=hyperlinks
        )

    def add_index(
        self,
        *,
        columns: int = 2,
        run_in: bool = False,
        right_align_page_numbers: bool = False,
    ) -> Any:
        """Insert a back-of-book index at the very end of the document.

        The "indexes live at the back" helper — sugar for
        `self.end.insert_index(...)`. See
        [`Anchor.insert_index`][wordlive.Anchor.insert_index] for the full
        semantics of `columns`, `run_in`, and `right_align_page_numbers`, and for
        the page-number-repagination caveat. Mark entries first with
        [`Anchor.mark_index_entry`][wordlive.Anchor.mark_index_entry]. Returns the
        new [`Index`][wordlive.Index]. Wrap in `doc.edit(...)` for atomic undo.
        """
        return self.end.insert_index(
            columns=columns, run_in=run_in, right_align_page_numbers=right_align_page_numbers
        )

    def add_bibliography(self) -> Any:
        """Insert a bibliography at the very end of the document.

        The "references live at the back" helper — sugar for
        `self.end.insert_bibliography()`. See
        [`Anchor.insert_bibliography`][wordlive.Anchor.insert_bibliography] for the
        field-block / repagination caveat. Add sources with
        `doc.sources.add` and cite them with
        [`Anchor.insert_citation`][wordlive.Anchor.insert_citation] first. Returns
        the new [`Bibliography`][wordlive.Bibliography]. Wrap in `doc.edit(...)`.
        """
        return self.end.insert_bibliography()

    def add_table_of_authorities(
        self,
        *,
        category: str | int = "all",
        passim: bool = True,
        keep_entry_formatting: bool = True,
        entry_separator: str | None = None,
        page_range_separator: str | None = None,
    ) -> Any:
        """Insert a table of authorities at the very end of the document.

        Sugar for `self.end.insert_table_of_authorities(...)`. See
        [`Anchor.insert_table_of_authorities`][wordlive.Anchor.insert_table_of_authorities]
        for the full semantics of `category`, `passim`, `keep_entry_formatting`,
        and the separators, and for the page-number-repagination caveat. Mark
        citations first with [`Anchor.mark_citation`][wordlive.Anchor.mark_citation].
        Returns the new [`TableOfAuthorities`][wordlive.TableOfAuthorities]. Wrap in
        `doc.edit(...)` for atomic undo.
        """
        return self.end.insert_table_of_authorities(
            category=category,
            passim=passim,
            keep_entry_formatting=keep_entry_formatting,
            entry_separator=entry_separator,
            page_range_separator=page_range_separator,
        )
