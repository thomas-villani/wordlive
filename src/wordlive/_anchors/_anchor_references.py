"""Insert TOC/index/bibliography/TOA entries, captions, links, cross-references."""

from __future__ import annotations

from typing import TYPE_CHECKING, Any

from .. import _com
from ..constants import (
    WdCaptionPosition,
    WdCollapseDirection,
    WdFieldType,
    WdIndexType,
)
from ..exceptions import OpError

if TYPE_CHECKING:
    pass

from ._helpers import (
    _coerce_named,
)
from ._refs import (
    _caption_above,
    _cross_ref_kind,
    _resolve_cross_ref_target,
)

if TYPE_CHECKING:
    pass

from ._anchor_core import AnchorCore


class AnchorReferencesMixin(AnchorCore):
    """Insert TOC/index/bibliography/TOA entries, captions, links, cross-references."""

    def _caption_object_range(self) -> Any | None:
        """Return a Range selecting a caption-able *object* for `insert_caption`.

        Word's `InsertCaption` only honours its above/below `Position` when the
        range selects a real object — a whole `Table`, an `InlineShape`, or a
        floating `Shape`. A plain text anchor isn't one, so the base returns
        `None` (the caption gets its own paragraph instead); `Cell` overrides
        this to return its parent table's range so a table caption lands above /
        below the table rather than inside a cell.
        """
        return None

    def insert_toc(
        self,
        *,
        levels: tuple[int, int] = (1, 3),
        use_heading_styles: bool = True,
        hyperlinks: bool = True,
        where: str = "after",
    ) -> Any:
        """Insert a table of contents at this anchor and return it as a `Toc`.

        Builds a TOC from the document's heading paragraphs over the given
        `levels` (a ``(upper, lower)`` pair — `(1, 3)` covers Heading 1–3).
        `use_heading_styles=True` sources entries from the built-in Heading
        styles; `hyperlinks=True` makes each entry a clickable jump (and a real
        hyperlink in exported PDFs). Returns a [`Toc`][wordlive.Toc].

        A TOC's page numbers populate only after repagination — call
        `toc.update()` (or [`Document.update_fields`][wordlive.Document.update_fields],
        or take a `snapshot`, which forces print layout) before reading them.
        Most documents want the TOC at the top: `doc.add_toc(...)` is the sugar
        for `doc.start.insert_toc(...)`.

        `where` is ``"after"`` (default) or ``"before"`` this anchor's range.
        Wrap in `doc.edit(...)` for atomic undo. Bad input raises `OpError`.
        """
        from .._toc import Toc

        try:
            if where not in ("before", "after"):
                raise ValueError(f"where must be 'before' or 'after'; got {where!r}")
            try:
                upper, lower = int(levels[0]), int(levels[1])
            except (TypeError, IndexError, ValueError, KeyError) as e:
                raise ValueError(
                    f"levels must be a (upper, lower) pair of ints; got {levels!r}"
                ) from e
            if not (1 <= upper <= lower <= 9):
                raise ValueError(
                    f"levels must satisfy 1 <= upper <= lower <= 9; got {(upper, lower)}"
                )
            with _com.translate_com_errors():
                rng = self._range()
                pos = int(rng.Start) if where == "before" else int(rng.End)
                insert_rng = self._doc.com.Range(pos, pos)
                # Positional args (keyword names are dropped under pywin32 late
                # binding). Order: Range, UseHeadingStyles, UpperHeadingLevel,
                # LowerHeadingLevel, UseFields, TableID, RightAlignPageNumbers,
                # IncludePageNumbers, AddedStyles, UseHyperlinks,
                # HidePageNumbersInWeb, UseOutlineLevels.
                toc_com = self._doc.com.TablesOfContents.Add(
                    insert_rng,
                    bool(use_heading_styles),
                    upper,
                    lower,
                    False,
                    "",
                    True,
                    True,
                    "",
                    bool(hyperlinks),
                    True,
                    True,
                )
            return Toc(self._doc, toc_com)
        except (ValueError, TypeError) as e:
            raise OpError(str(e)) from e

    def link_to(
        self,
        address: str | None = None,
        *,
        bookmark: str | None = None,
        text: str | None = None,
        screen_tip: str | None = None,
    ) -> None:
        """Turn this anchor into a hyperlink (or insert new linked text).

        Pass exactly one destination: `address` for an external link (a URL,
        `mailto:`, or file path) or `bookmark` for an internal jump to a named
        bookmark in this document. With `text=None` the anchor's existing range
        becomes the clickable link; pass `text=...` to **insert** new linked text
        at the end of the anchor's range (so linking a heading or a `range:`
        phrase with `text=...` adds the link rather than overwriting the content).
        `screen_tip` is the hover tooltip.

        Pair it with [`doc.bookmarks.add(...)`][wordlive.BookmarkCollection.add]
        to build internal navigation, or a `range:START-END` id (from `find`) to
        link an existing phrase. Wrap in `doc.edit(...)` for atomic undo. Bad
        input (not exactly one destination) raises `OpError`.
        """
        try:
            if (address is None) == (bookmark is None):
                raise ValueError("link_to requires exactly one of 'address' or 'bookmark'")
            with _com.translate_com_errors():
                rng = self._range()
                if text is not None:
                    # Insert *new* linked text rather than overwriting the
                    # anchor's range: collapse to its end so a heading / phrase
                    # keeps its content and the link is added after it.
                    rng = rng.Duplicate
                    rng.Collapse(int(WdCollapseDirection.END))
                addr_arg = address or ""
                sub_arg = bookmark or ""
                tip_arg = screen_tip or ""
                # Positional args (Anchor, Address, SubAddress, ScreenTip,
                # TextToDisplay) — keep keywords out for late-binding safety.
                if text is not None:
                    self._doc.com.Hyperlinks.Add(rng, addr_arg, sub_arg, tip_arg, text)
                else:
                    self._doc.com.Hyperlinks.Add(rng, addr_arg, sub_arg, tip_arg)
        except (ValueError, TypeError) as e:
            raise OpError(str(e)) from e

    def insert_cross_reference(
        self,
        target: str,
        *,
        kind: str = "text",
        hyperlink: bool = True,
        where: str = "after",
    ) -> None:
        """Insert a cross-reference to another anchor at this anchor.

        `target` is the anchor id to point at: `bookmark:NAME`, `heading:N`,
        `footnote:N`, or `endnote:N`. `kind` selects what the reference shows:
        ``"text"`` (the heading/bookmark text — the default), ``"page"`` ("see
        page 5"), ``"number"`` (the paragraph or note number), or
        ``"above_below"`` ("above"/"below"). `hyperlink=True` makes the inserted
        reference a clickable jump.

        An unresolvable `target` raises `AnchorNotFoundError` (exit 2) before
        anything is inserted. `where` is ``"after"`` (default) or ``"before"``
        this anchor's range. Wrap in `doc.edit(...)` for atomic undo; other bad
        input raises `OpError`.
        """
        try:
            if where not in ("before", "after"):
                raise ValueError(f"where must be 'before' or 'after'; got {where!r}")
            # Resolve outside translate_com_errors so an AnchorNotFoundError for a
            # bad target propagates as exit 2 rather than being masked.
            ref_type, ref_item = _resolve_cross_ref_target(self._doc, target)
            ref_kind = _cross_ref_kind(kind, ref_type)
            with _com.translate_com_errors():
                insert_rng = self._range().Duplicate
                insert_rng.Collapse(
                    int(WdCollapseDirection.START if where == "before" else WdCollapseDirection.END)
                )
                # Positional args: IncludePositionInformation as a keyword raises
                # under pywin32 late binding, so pass only the first four.
                insert_rng.InsertCrossReference(ref_type, ref_kind, ref_item, bool(hyperlink))
        except (ValueError, TypeError) as e:
            raise OpError(str(e)) from e

    def insert_caption(
        self, label: str = "Figure", *, text: str | None = None, position: str | None = None
    ) -> None:
        """Insert a numbered caption as its **own paragraph** at this anchor.

        `label` is a caption label — built-in ``"Figure"`` / ``"Table"`` /
        ``"Equation"`` or any custom string; Word auto-numbers per label
        (Figure 1, Figure 2, …). `text` is the caption title shown after the
        label and number. Pairs with
        [`insert_cross_reference`][wordlive.Anchor.insert_cross_reference] for
        "see Figure 2".

        `position` is ``"above"`` or ``"below"`` the anchor. Left as `None` it
        follows convention: a ``"Table"`` caption goes **above**, every other
        label goes **below**. The caption always becomes its own
        `Caption`-styled paragraph — it never fuses into the target paragraph.
        On a table cell (`table:N:R:C`) the caption is placed above / below the
        **whole table**, not inside the cell.

        Wrap in `doc.edit(...)` for atomic undo. Bad input raises `OpError`.
        """
        try:
            above = _caption_above(label, position)
            title = text if text is not None else ""
            pos = int(WdCaptionPosition.ABOVE if above else WdCaptionPosition.BELOW)
            with _com.translate_com_errors():
                obj_rng = self._caption_object_range()
                if obj_rng is not None:
                    # A caption-able object (e.g. a table): let Word place the
                    # caption on its own line above/below the object natively.
                    obj_rng.InsertCaption(str(label), title, pos, False)
                else:
                    # Text/paragraph anchor: carve out a dedicated empty
                    # paragraph (before or after the anchor) and drop the
                    # caption into it, so it never fuses into the host paragraph.
                    insert_rng = self._range().Duplicate
                    insert_rng.Collapse(
                        int(WdCollapseDirection.START if above else WdCollapseDirection.END)
                    )
                    insert_rng.InsertParagraphBefore()
                    insert_rng.Collapse(int(WdCollapseDirection.START))
                    # Positional args (Label, Title, Position, ExcludeLabel) for
                    # late-binding safety; a string Label matches a built-in or
                    # defines a custom one.
                    insert_rng.InsertCaption(str(label), title, pos, False)
        except (ValueError, TypeError) as e:
            raise OpError(str(e)) from e

    def mark_index_entry(
        self,
        entry: str,
        *,
        cross_reference: str | None = None,
        bold: bool = False,
        italic: bool = False,
    ) -> None:
        """Mark this anchor's range as a back-of-book index entry (an `XE` field).

        `entry` is the text that appears in the index; use ``"main:sub"`` to file
        it as a subentry under ``main`` (Word's colon convention). `bold` /
        `italic` style the entry's *page number* in the built index.
        `cross_reference` replaces the page number with a "see …" pointer (e.g.
        ``cross_reference="Widgets"`` → *"see Widgets"*).

        This is the per-term half of indexing; once entries are marked, build the
        list with [`insert_index`][wordlive.Anchor.insert_index] /
        [`Document.add_index`][wordlive.Document.add_index]. The `XE` field is
        hidden text and doesn't disturb the visible flow. Wrap in `doc.edit(...)`
        for atomic undo. Bad input raises `OpError`.
        """
        try:
            if not str(entry).strip():
                raise ValueError("entry must be a non-empty string")
            with _com.translate_com_errors():
                rng = self._range()
                # Indexes.MarkEntry(Range, Entry, EntryAutoText, CrossReference,
                # CrossReferenceAutoText, BookmarkName, Bold, Italic) — positional
                # for late-binding safety.
                self._doc.com.Indexes.MarkEntry(
                    rng,
                    str(entry),
                    "",
                    str(cross_reference) if cross_reference is not None else "",
                    "",
                    "",
                    bool(bold),
                    bool(italic),
                )
        except (ValueError, TypeError) as e:
            raise OpError(str(e)) from e

    def insert_index(
        self,
        *,
        columns: int = 2,
        run_in: bool = False,
        right_align_page_numbers: bool = False,
        where: str = "after",
    ) -> Any:
        """Insert a back-of-book index at this anchor and return it as an `Index`.

        Gathers every `XE` entry marked with
        [`mark_index_entry`][wordlive.Anchor.mark_index_entry] into an
        alphabetised, page-numbered list. `columns` is the number of newspaper
        columns the index is laid out in (2 is the book default). `run_in=True`
        packs subentries into a single paragraph instead of one per line;
        `right_align_page_numbers=True` flushes page numbers to the right margin.

        Returns an [`Index`][wordlive.Index]; like a TOC it's a field block whose
        page numbers populate only after repagination — call `index.update()`,
        [`Document.update_fields`][wordlive.Document.update_fields], or take a
        `snapshot`. Most documents want the index at the end:
        `doc.add_index(...)` is the sugar for `doc.end.insert_index(...)`.

        `where` is ``"after"`` (default) or ``"before"`` this anchor's range.
        Wrap in `doc.edit(...)` for atomic undo. Bad input raises `OpError`.
        """
        from .._index import Index

        try:
            if where not in ("before", "after"):
                raise ValueError(f"where must be 'before' or 'after'; got {where!r}")
            cols = int(columns)
            if cols < 1:
                raise ValueError(f"columns must be >= 1; got {columns!r}")
            idx_type = int(WdIndexType.RUNIN if run_in else WdIndexType.INDENT)
            with _com.translate_com_errors():
                rng = self._range()
                pos = int(rng.Start) if where == "before" else int(rng.End)
                insert_rng = self._doc.com.Range(pos, pos)
                # Indexes.Add(Range, HeadingSeparator, RightAlignPageNumbers, Type,
                # NumberOfColumns, AccentedLetters, SortBy, IndexLanguage). Positional;
                # HeadingSeparator must be a WdHeadingSeparator value (0 = none) — an
                # empty string raises a type-mismatch on a makepy-typed Word wrapper.
                idx_com = self._doc.com.Indexes.Add(
                    insert_rng, 0, bool(right_align_page_numbers), idx_type, cols
                )
            return Index(self._doc, idx_com)
        except (ValueError, TypeError) as e:
            raise OpError(str(e)) from e

    def insert_table_of_figures(
        self,
        *,
        label: str = "Figure",
        include_label: bool = True,
        hyperlinks: bool = True,
        right_align_page_numbers: bool = True,
        where: str = "after",
    ) -> Any:
        """Insert a table of figures at this anchor and return it.

        The caption-driven sibling of [`insert_toc`][wordlive.Anchor.insert_toc]:
        it lists every caption of one `label` — ``"Figure"`` (the default),
        ``"Table"``, ``"Equation"``, or any custom label you passed to
        [`insert_caption`][wordlive.Anchor.insert_caption] — with its page number.
        `include_label=True` keeps the "Figure 1" prefix in each entry;
        `hyperlinks=True` makes entries clickable jumps;
        `right_align_page_numbers=True` flushes page numbers right.

        Returns a [`TableOfFigures`][wordlive.TableOfFigures]; like a TOC its page
        numbers populate only after repagination — call `tof.update()`,
        [`Document.update_fields`][wordlive.Document.update_fields], or take a
        `snapshot`. `where` is ``"after"`` (default) or ``"before"`` this anchor's
        range. Wrap in `doc.edit(...)` for atomic undo. Bad input raises `OpError`.
        """
        from .._toc import TableOfFigures

        try:
            if where not in ("before", "after"):
                raise ValueError(f"where must be 'before' or 'after'; got {where!r}")
            with _com.translate_com_errors():
                rng = self._range()
                pos = int(rng.Start) if where == "before" else int(rng.End)
                insert_rng = self._doc.com.Range(pos, pos)
                # Keyword args here: the optional string Variants (TableID,
                # Caption2, AddedStyles) raise a type-mismatch when passed
                # positionally as "" on a makepy-typed Word wrapper, so we name
                # only the flags we set and let the rest default (Range + Caption
                # stay positional, matching the Word signature).
                tof_com = self._doc.com.TablesOfFigures.Add(
                    insert_rng,
                    str(label),
                    IncludeLabel=bool(include_label),
                    UseHeadingStyles=False,
                    RightAlignPageNumbers=bool(right_align_page_numbers),
                    UseHyperlinks=bool(hyperlinks),
                )
            return TableOfFigures(self._doc, tof_com)
        except (ValueError, TypeError) as e:
            raise OpError(str(e)) from e

    def insert_citation(
        self,
        tag: str,
        *,
        pages: str | None = None,
        prefix: str | None = None,
        suffix: str | None = None,
        volume: str | None = None,
        suppress_author: bool = False,
        suppress_year: bool = False,
        suppress_title: bool = False,
        locale: int = 1033,
        where: str = "after",
    ) -> Any:
        """Insert an in-text citation at this anchor and return it as a `Citation`.

        References a source in the document's store (add one with
        `doc.sources.add`) by its `tag` and
        renders it in the current [`bibliography_style`][wordlive.Document.bibliography_style]
        — e.g. *(Smith 2020)*. `pages` adds a page locator (*(Smith 2020, 15)*);
        `prefix` / `suffix` wrap the citation (*"see "* / *", at 12"*); `volume`
        adds a volume. `suppress_author` / `suppress_year` / `suppress_title` drop
        those parts. `locale` is the LCID the style formats under (1033 = en-US).

        Returns a [`Citation`][wordlive.Citation]; a citation to an unknown tag
        renders *"Invalid source specified."* rather than failing. `where` is
        ``"after"`` (default) or ``"before"`` this anchor's range. Wrap in
        `doc.edit(...)` for atomic undo. Bad input raises `OpError`.
        """
        from .._citations import Citation

        try:
            if where not in ("before", "after"):
                raise ValueError(f"where must be 'before' or 'after'; got {where!r}")
            if not str(tag).strip():
                raise ValueError("tag must be a non-empty string")
            # CITATION field code (switches confirmed against live Word): \p page,
            # \v volume, \f prefix, \s suffix, \n/\y/\t suppress author/year/title.
            parts = [f"CITATION {tag} \\l {int(locale)}"]
            if pages:
                parts.append(f'\\p "{pages}"')
            if volume:
                parts.append(f"\\v {volume}")
            if prefix:
                parts.append(f'\\f "{prefix}"')
            if suffix:
                parts.append(f'\\s "{suffix}"')
            if suppress_author:
                parts.append("\\n")
            if suppress_year:
                parts.append("\\y")
            if suppress_title:
                parts.append("\\t")
            code = " ".join(parts)
            with _com.translate_com_errors():
                # Collapse a *duplicate* of the anchor's own range so the field
                # lands in the same story (header/footer-safe — see insert_field).
                insert_rng = self._range().Duplicate
                insert_rng.Collapse(
                    int(WdCollapseDirection.START if where == "before" else WdCollapseDirection.END)
                )
                # EMPTY (-1) raw-code insert — positional, the proven path Word
                # parses into a typed CITATION field (its numeric, 96, is fragile
                # to pass directly).
                field = insert_rng.Fields.Add(insert_rng, int(WdFieldType.EMPTY), code, False)
            return Citation(self._doc, field)
        except (ValueError, TypeError) as e:
            raise OpError(str(e)) from e

    def insert_bibliography(self, *, where: str = "after") -> Any:
        """Insert a bibliography at this anchor and return it as a `Bibliography`.

        Inserts a ``BIBLIOGRAPHY`` field — the reference list of every source
        *cited* in the document, formatted in the current
        [`bibliography_style`][wordlive.Document.bibliography_style]. Most
        documents want it at the end: `doc.add_bibliography()` is the sugar for
        `doc.end.insert_bibliography()`.

        Returns a [`Bibliography`][wordlive.Bibliography]; like a TOC it's a field
        block — call `bibliography.update()`,
        [`Document.update_fields`][wordlive.Document.update_fields], or take a
        `snapshot` after adding citations. `where` is ``"after"`` (default) or
        ``"before"`` this anchor's range. Wrap in `doc.edit(...)` for atomic undo.
        Bad input raises `OpError`.
        """
        from .._citations import Bibliography

        try:
            if where not in ("before", "after"):
                raise ValueError(f"where must be 'before' or 'after'; got {where!r}")
            with _com.translate_com_errors():
                insert_rng = self._range().Duplicate
                insert_rng.Collapse(
                    int(WdCollapseDirection.START if where == "before" else WdCollapseDirection.END)
                )
                field = insert_rng.Fields.Add(
                    insert_rng, int(WdFieldType.EMPTY), "BIBLIOGRAPHY", False
                )
            return Bibliography(self._doc, field)
        except (ValueError, TypeError) as e:
            raise OpError(str(e)) from e

    def mark_citation(
        self,
        long_citation: str,
        *,
        short_citation: str | None = None,
        category: str | int = "cases",
        where: str = "after",
    ) -> None:
        """Mark this anchor's range as a table-of-authorities citation (a `TA` field).

        The legal analog of [`mark_index_entry`][wordlive.Anchor.mark_index_entry]:
        `long_citation` is the full citation as it appears in the table (e.g.
        *"Smith v. Jones, 1 U.S. 1 (2020)"*), `short_citation` the abbreviated
        form Word matches elsewhere in the text (defaults to `long_citation`), and
        `category` the section it files under — ``"cases"`` (the default),
        ``"statutes"``, ``"other"``, ``"rules"``, ``"treatises"``,
        ``"regulations"``, ``"constitutional"``, or a category number (1-16).

        This is the per-authority half; build the table with
        [`insert_table_of_authorities`][wordlive.Anchor.insert_table_of_authorities]
        / [`Document.add_table_of_authorities`][wordlive.Document.add_table_of_authorities].
        The `TA` field is hidden and doesn't disturb the visible flow. Wrap in
        `doc.edit(...)` for atomic undo. Bad input raises `OpError`.
        """
        from .._toa import _TOA_CATEGORIES

        try:
            if where not in ("before", "after"):
                raise ValueError(f"where must be 'before' or 'after'; got {where!r}")
            if not str(long_citation).strip():
                raise ValueError("long_citation must be a non-empty string")
            cat = _coerce_named(category, _TOA_CATEGORIES, "TOA category")
            short = short_citation if short_citation is not None else long_citation
            code = f'TA \\l "{long_citation}" \\s "{short}" \\c {cat}'
            with _com.translate_com_errors():
                insert_rng = self._range().Duplicate
                insert_rng.Collapse(
                    int(WdCollapseDirection.START if where == "before" else WdCollapseDirection.END)
                )
                insert_rng.Fields.Add(insert_rng, int(WdFieldType.EMPTY), code, False)
        except (ValueError, TypeError) as e:
            raise OpError(str(e)) from e

    def insert_table_of_authorities(
        self,
        *,
        category: str | int = "all",
        passim: bool = True,
        keep_entry_formatting: bool = True,
        entry_separator: str | None = None,
        page_range_separator: str | None = None,
        where: str = "after",
    ) -> Any:
        """Insert a table of authorities at this anchor and return it.

        Gathers the `TA` citations marked with
        [`mark_citation`][wordlive.Anchor.mark_citation] into a page-numbered
        table. `category` selects which authorities to include — ``"all"`` (the
        default), ``"cases"``, ``"statutes"``, … or a number (1-16). `passim=True`
        replaces five-or-more page references for one authority with *"passim"*;
        `keep_entry_formatting=True` preserves each citation's character
        formatting. `entry_separator` / `page_range_separator` override the
        defaults between a citation and its first page / between page ranges.

        Returns a [`TableOfAuthorities`][wordlive.TableOfAuthorities]; like a TOC
        it's a field block whose page numbers populate only after repagination —
        call `toa.update()`, [`Document.update_fields`][wordlive.Document.update_fields],
        or take a `snapshot`. `doc.add_table_of_authorities(...)` is the sugar for
        `doc.end.insert_table_of_authorities(...)`. `where` is ``"after"``
        (default) or ``"before"`` this anchor's range. Wrap in `doc.edit(...)` for
        atomic undo. Bad input raises `OpError`.
        """
        from .._toa import _TOA_CATEGORIES, TableOfAuthorities

        try:
            if where not in ("before", "after"):
                raise ValueError(f"where must be 'before' or 'after'; got {where!r}")
            cat = _coerce_named(category, _TOA_CATEGORIES, "TOA category")
            with _com.translate_com_errors():
                rng = self._range()
                pos = int(rng.Start) if where == "before" else int(rng.End)
                insert_rng = self._doc.com.Range(pos, pos)
                # TablesOfAuthorities.Add(Range, Category, ...): Range + int
                # Category positional; the string-Variant optionals need keyword
                # form (same gotcha as TablesOfFigures).
                kwargs: dict[str, Any] = {
                    "Passim": bool(passim),
                    "KeepEntryFormatting": bool(keep_entry_formatting),
                }
                if entry_separator is not None:
                    kwargs["EntrySeparator"] = str(entry_separator)
                if page_range_separator is not None:
                    kwargs["PageRangeSeparator"] = str(page_range_separator)
                toa_com = self._doc.com.TablesOfAuthorities.Add(insert_rng, cat, **kwargs)
            return TableOfAuthorities(self._doc, toa_com)
        except (ValueError, TypeError) as e:
            raise OpError(str(e)) from e
