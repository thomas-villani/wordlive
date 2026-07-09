"""Insert images, text boxes, charts, and equations."""

from __future__ import annotations

import secrets
from typing import TYPE_CHECKING, Any

from .. import _charts, _com, _equations, _images, _shapes
from .._format import to_bgr, to_points
from ..constants import (
    MsoTextOrientation,
    MsoTriState,
    WdParagraphAlignment,
)
from ..exceptions import EquationError, ExcelNotAvailableError, OpError

if TYPE_CHECKING:
    from pathlib import Path


from ._helpers import (
    _ALIGNMENT_NAMES,
    _WRAP_NAMES,
    _WRAP_VALUES,
    _chart_index_at,
    _equation_index_at,
    _resolve_wrap,
    _utf16_len,
)

if TYPE_CHECKING:
    from ._chart_anchors import ChartAnchor
    from ._equation_anchors import EquationAnchor
    from ._shape_anchors import ShapeAnchor

from ._anchor_core import AnchorCore


class AnchorMediaMixin(AnchorCore):
    """Insert images, text boxes, charts, and equations."""

    def insert_image(
        self,
        image: str | Path | bytes,
        *,
        wrap: str,
        where: str = "after",
        block: bool = False,
        width: float | None = None,
        height: float | None = None,
        alt_text: str | None = None,
        lock_aspect: bool = True,
    ) -> ShapeAnchor | None:
        """Insert an image at this anchor (atomic-undo when inside `doc.edit()`).

        `image` is a file path, raw image bytes, or a base64 string — a `str`
        is treated as a path when it names an existing file, otherwise as
        base64. Word embeds the picture (`SaveWithDocument=True`) and
        auto-detects its natural size, so `width`/`height` (points) are optional
        overrides. `alt_text` sets the image's accessibility text.

        `wrap` is required — there is no default — so layout intent is always
        explicit:

        - ``"inline"`` keeps the image in the text flow (an `InlineShape`).
        - ``"auto"`` floats it: Square when its width is at most half the
          section's usable text width, else top-and-bottom.
        - ``"square" | "tight" | "through" | "top-bottom" | "front" | "behind"``
          floats it with that wrap type.

        `where` is ``"after"`` (default) or ``"before"`` the anchor's range.

        `block` places the image in its own new paragraph (reset to ``Normal``)
        rather than embedding it in the anchor's text run — so
        ``heading.insert_image(..., wrap="inline", where="before", block=True)``
        drops the image on its own line *above* the heading instead of joining
        the heading text. Without it, an inline image anchored at a heading lands
        mid-run and the heading text trails it on the same line.

        A floating image (any `wrap` other than ``"inline"``) leaves the inline
        text flow, so `image:N` no longer addresses it — this returns its floating
        [`ShapeAnchor`][wordlive.ShapeAnchor] (`shape:N`) for restyle
        (re-wrap / reposition / resize / `replace_image`). An ``"inline"`` image
        stays an `InlineShape` (addressed as `image:N`) and returns ``None``.

        Raises `ImageSourceError` for a missing/unreadable/invalid image and
        `ValueError` for an unknown `wrap` or `where`.
        """
        if wrap not in _WRAP_VALUES:
            raise ValueError(f"unknown wrap {wrap!r}; expected one of {sorted(_WRAP_VALUES)}")
        if where not in ("before", "after"):
            raise ValueError(f"where must be 'before' or 'after'; got {where!r}")
        # New paragraphs inherit the anchor's style — a block image above a
        # heading would otherwise become a heading-styled (and outline-polluting)
        # paragraph. Reset it to the body default, like insert_table does.
        normal_obj = self._doc.styles["Normal"] if block and "Normal" in self._doc.styles else None
        with _images.image_on_disk(image) as disk_path:
            with _com.translate_com_errors():
                doc_com = self._doc.com
                rng = self._range()
                pos = int(rng.Start) if where == "before" else int(rng.End)
                if block:
                    # Open a fresh paragraph at the insertion point and target it,
                    # so the image sits on its own line instead of in the run.
                    doc_com.Range(pos, pos).Text = "\r"
                    if normal_obj is not None:
                        doc_com.Range(pos, pos).Paragraphs(1).Range.Style = normal_obj.com
                insert_rng = doc_com.Range(pos, pos)
                ish = insert_rng.InlineShapes.AddPicture(
                    FileName=disk_path,
                    LinkToFile=False,
                    SaveWithDocument=True,
                    Range=insert_rng,
                )
                ish.LockAspectRatio = int(MsoTriState.TRUE if lock_aspect else MsoTriState.FALSE)
                if width is not None:
                    ish.Width = float(width)
                if height is not None:
                    ish.Height = float(height)
                if alt_text is not None:
                    ish.AlternativeText = alt_text
                if wrap == "inline":
                    return None
                wrap_type = _resolve_wrap(wrap, ish, insert_rng)
                shape = ish.ConvertToShape()
                shape.WrapFormat.Type = int(wrap_type)
                if alt_text is not None:
                    # AlternativeText doesn't always survive the conversion.
                    shape.AlternativeText = alt_text
                # The picture left InlineShapes (image:N no longer addresses it),
                # so hand back its floating shape:N handle for restyle. Locate by a
                # unique temp name — don't assume "last" (other floats can reorder).
                orig_name = str(shape.Name or "")
                probe_name = f"_wl_shape_{secrets.token_hex(8)}"
                shape.Name = probe_name
                index = _shapes.index_of_named(doc_com, probe_name)
                # Restore unconditionally — leaving the probe name on a shape whose
                # original name was empty would surface `_wl_shape_*` in list().
                shape.Name = orig_name
            from ._shape_anchors import ShapeAnchor  # lazy: _shape_anchors imports Anchor

            return ShapeAnchor(self._doc, index)

    def insert_text_box(
        self,
        text: str,
        *,
        width: Any = 200,
        height: Any = 100,
        wrap: str = "square",
        where: str = "after",
        font: str | None = None,
        size: Any = None,
        bold: bool | None = None,
        italic: bool | None = None,
        alignment: str | None = None,
        fill: str | None = None,
        border: str | bool | None = None,
    ) -> ShapeAnchor:
        """Insert a floating text box (a pull quote / call-out) anchored here.

        A `Shapes.AddTextbox` floating shape is anchored to this anchor's
        paragraph and seeded with `text`. `width` / `height` are points or a unit
        string (``"3in"`` / ``"8cm"``). `wrap` is how body text flows around it —
        ``"square"`` (default), ``"tight"``, ``"through"``, ``"top-bottom"``,
        ``"front"``, or ``"behind"`` (the same vocabulary as `insert_image`, minus
        ``"inline"``). `where` places the anchor ``"after"`` (default) or
        ``"before"`` this anchor's range.

        The remaining kwargs style the box and its text, each optional:
        `font` / `size` (points or unit string) / `bold` / `italic` set the
        character format; `alignment` (``"left"``/``"center"``/``"right"``/
        ``"justify"``) the paragraph; `fill` is a background colour
        (``"#eeeeff"`` / ``"navy"``) and `border` is ``False`` for no outline, a
        colour string for a coloured outline, or ``True`` for the default.

        Returns the text box's floating [`ShapeAnchor`][wordlive.ShapeAnchor]
        (`shape:N`) so it can be restyled in place afterwards (`set_text` /
        `set_wrap` / `set_position` / `set_size` / `format`); discover text boxes
        later via [`doc.text_boxes`][wordlive.Document.text_boxes]. Wrap in
        `doc.edit(...)` for atomic undo; raises `ValueError` for an unknown
        `wrap` / `where`.
        """
        if wrap not in _WRAP_NAMES:
            raise ValueError(
                f"unknown wrap {wrap!r}; expected one of {sorted(_WRAP_NAMES)} (text boxes float)"
            )
        if where not in ("before", "after"):
            raise ValueError(f"where must be 'before' or 'after'; got {where!r}")
        wd_align = _ALIGNMENT_NAMES[alignment] if alignment is not None else None
        try:
            w = to_points(width)
            h = to_points(height)
            font_size = to_points(size) if size is not None else None
            fill_bgr = to_bgr(fill) if fill is not None else None
            border_bgr = to_bgr(border) if isinstance(border, str) else None
        except (ValueError, TypeError) as e:
            raise OpError(str(e)) from e
        with _com.translate_com_errors():
            doc_com = self._doc.com
            rng = self._range()
            pos = int(rng.Start) if where == "before" else int(rng.End)
            anchor_rng = doc_com.Range(pos, pos)
            shape = doc_com.Shapes.AddTextbox(
                Orientation=int(MsoTextOrientation.HORIZONTAL),
                Left=0.0,
                Top=0.0,
                Width=w,
                Height=h,
                Anchor=anchor_rng,
            )
            text_range = shape.TextFrame.TextRange
            text_range.Text = text
            font_obj = text_range.Font
            if font is not None:
                font_obj.Name = font
            if font_size is not None:
                font_obj.Size = font_size
            if bold is not None:
                font_obj.Bold = int(MsoTriState.TRUE if bold else MsoTriState.FALSE)
            if italic is not None:
                font_obj.Italic = int(MsoTriState.TRUE if italic else MsoTriState.FALSE)
            if wd_align is not None:
                text_range.ParagraphFormat.Alignment = int(wd_align)
            shape.WrapFormat.Type = int(_WRAP_NAMES[wrap])
            if fill_bgr is not None:
                shape.Fill.Visible = int(MsoTriState.TRUE)
                shape.Fill.Solid()
                shape.Fill.ForeColor.RGB = fill_bgr
            if border is False:
                shape.Line.Visible = int(MsoTriState.FALSE)
            elif border_bgr is not None:
                shape.Line.Visible = int(MsoTriState.TRUE)
                shape.Line.ForeColor.RGB = border_bgr
            # Hand back the new text box's shape:N handle (locate by a unique temp
            # name — don't assume "last", other floats can reorder).
            orig_name = str(shape.Name or "")
            probe_name = f"_wl_shape_{secrets.token_hex(8)}"
            shape.Name = probe_name
            index = _shapes.index_of_named(doc_com, probe_name)
            # Restore unconditionally so an empty original name doesn't leave the
            # `_wl_shape_*` probe lingering in list().
            shape.Name = orig_name
        from ._shape_anchors import ShapeAnchor  # lazy: _shape_anchors imports Anchor

        return ShapeAnchor(self._doc, index)

    def insert_chart(
        self,
        kind: str,
        data: Any,
        *,
        title: str | None = None,
        where: str = "after",
    ) -> ChartAnchor:
        """Insert an Excel-backed chart at this anchor and return it.

        `kind` is one of ``"bar"`` (clustered columns), ``"pie"``, ``"line"``, or
        ``"scatter"``. `data` is either an object mapping ``{label: value}`` (for
        bar / pie / line) or an array of ``[x, y]`` pairs (for ``scatter`` — both
        axes numeric, duplicate x preserved — and ``line``). `title` sets the
        chart title and series name; ``None`` leaves it untitled. `where` places
        the chart ``"after"`` (default) or ``"before"`` this anchor's range.

        Charts are Excel-backed: this embeds a chart whose data lives in a hidden
        Excel workbook, then breaks the link so the data is **static** — no live
        workbook ships in the document and the series data can't be read back
        (deferred). Requires Excel installed: raises `ExcelNotAvailableError`
        (CLI exit 6), checked up front so the document is untouched on a missing
        Excel. Raises `OpError` for malformed `data` and `ValueError` for an
        unknown `kind` / `where`.

        Word's chart API only inserts off the live `Selection`, so this moves the
        cursor to the insertion point; wrap in `doc.edit(...)` (as the CLI / exec
        / MCP surfaces do) for atomic undo and to restore the user's selection.
        """
        if kind not in _charts.KIND_TO_XL:
            raise ValueError(
                f"unknown chart kind {kind!r}; expected one of {sorted(_charts.KIND_TO_XL)}"
            )
        if where not in ("before", "after"):
            raise ValueError(f"where must be 'before' or 'after'; got {where!r}")
        xs, ys = _charts.normalize_chart_data(kind, data)
        if not _charts.probe_excel_available():
            raise ExcelNotAvailableError()
        xl_type = int(_charts.KIND_TO_XL[kind])
        with _com.translate_com_errors():
            doc_com = self._doc.com
            rng = self._range()
            pos = int(rng.Start) if where == "before" else int(rng.End)
            # AddChart2 only works off the Selection, never an arbitrary Range
            # (a Range raises "Requested object is not available"). doc.edit()
            # restores the user's selection on exit.
            doc_com.Range(pos, pos).Select()
            shape = doc_com.Application.Selection.InlineShapes.AddChart2(-1, xl_type)
            try:
                _charts.populate_chart(shape.Chart, kind, xs, ys, title)
            except Exception:
                # Don't leave a half-built placeholder chart behind on failure.
                try:
                    shape.Delete()
                except Exception:
                    pass
                raise
            index = _chart_index_at(doc_com, int(shape.Range.Start))
        from ._chart_anchors import ChartAnchor  # lazy: _chart_anchors imports Anchor

        return ChartAnchor(self._doc, index)

    def insert_equation(
        self,
        *,
        unicodemath: str | None = None,
        latex: str | None = None,
        mathml: str | None = None,
        where: str = "after",
        display: bool = True,
    ) -> EquationAnchor:
        """Insert a mathematical equation at this anchor and return it.

        The equation is given in exactly one of three input dialects:

        - ``unicodemath=`` — Word's native **UnicodeMath** linear form, e.g.
          ``"x=(-b±√(b^2-4ac))/(2a)"`` or ``"a^2+b^2=c^2"``. Zero-dependency: the
          string is typed into a math zone and *built up* into the 2-D form by
          Word itself.
        - ``latex=`` — a **LaTeX** math string, e.g.
          ``r"\\frac{-b\\pm\\sqrt{b^2-4ac}}{2a}"``. Converted LaTeX→MathML→OMML;
          the LaTeX→MathML hop needs the optional ``latex`` extra
          (`pip install "wordlive[latex]"`) and raises `EquationError` without it.
        - ``mathml=`` — a **MathML** (``<math>…</math>``) string. Converted
          MathML→OMML through Office's own transform (no extra needed).

        The equation always lands on its **own paragraph**, and that paragraph's
        style is pinned so it never inherits the style of whatever it was
        inserted next to (an equation dropped before a `Heading 2` used to come
        out *styled* `Heading 2` and land in the outline/TOC). `display` (default
        ``True``) gives it the dedicated centred ``Equation`` paragraph style
        (created on first use, based on ``Normal`` — a stable hook for later
        equation numbering); ``display=False`` resets the paragraph to ``Normal``
        and left-aligns it (it is still its own paragraph — wordlive does not
        place math mid-sentence — but reads as body text, not centred display
        math). `where` is ``"after"`` (default) or ``"before"`` this anchor's
        range — so ``doc.headings["Derivation"].insert_equation(...)`` drops an
        equation under a heading and ``doc.end.insert_equation(...)`` appends one.

        Returns an [`EquationAnchor`][wordlive.EquationAnchor] (`equation:N`);
        read it back as MathML with `equation.mathml`, or discover every equation
        via [`doc.equations`][wordlive.Document.equations]. Wrap in
        `doc.edit(...)` for atomic undo. Raises `EquationError` for malformed
        input (none, or more than one, of the three dialects; unparseable
        MathML/LaTeX; a missing LaTeX backend) and `ValueError` for a bad `where`.
        """
        given = [
            name
            for name, value in (
                ("unicodemath", unicodemath),
                ("latex", latex),
                ("mathml", mathml),
            )
            if value is not None
        ]
        if len(given) != 1:
            raise EquationError(
                "insert_equation needs exactly one of unicodemath=, latex=, or mathml="
                + (f"; got {', '.join(given)}" if given else "")
            )
        if where not in ("before", "after"):
            raise ValueError(f"where must be 'before' or 'after'; got {where!r}")
        if unicodemath is not None:
            return self._insert_equation_native(unicodemath, where=where, display=display)
        mathml_src = _equations.latex_to_mathml(latex) if latex is not None else (mathml or "")
        omml_inner = _equations.mathml_to_omml(mathml_src)
        return self._insert_equation_omml(omml_inner, where=where, display=display)

    def _equation_paragraph_span(self, where: str) -> tuple[int, int]:
        """Return the `(start, end)` of the document paragraph the equation attaches to.

        An equation always lands on its own paragraph, so insertion targets a
        *paragraph mark*, never a mid-paragraph offset — addressing off the
        anchor's raw range would land inside a math zone (`equation:N`) or
        mid-sentence (a bookmark). We resolve the paragraph containing the
        relevant edge of the anchor: its **start** for ``"before"``, its last real
        character (``End - 1``, clamped off the terminal mark) for ``"after"``.
        """
        rng = self._range()
        doc_com = self._doc.com
        doc_end = int(doc_com.Content.End)
        if where == "before":
            probe = max(0, int(rng.Start))
        else:
            probe = min(max(int(rng.Start), int(rng.End) - 1), max(0, doc_end - 1))
        para = doc_com.Range(probe, probe).Paragraphs(1).Range
        return int(para.Start), int(para.End)

    def _insert_equation_native(
        self, unicodemath: str, *, where: str, display: bool
    ) -> EquationAnchor:
        """Native UnicodeMath path: type the linear string, wrap it, BuildUp.

        Opens a fresh paragraph at the containing paragraph's boundary, writes the
        linear string into it, wraps the run in an `OMaths.Add` zone, and asks
        Word to build it up into the 2-D form. No XML, no extra dependency.
        """
        with _com.translate_com_errors():
            doc_com = self._doc.com
            pstart, pend = self._equation_paragraph_span(where)
            doc_end = int(doc_com.Content.End)
            if where == "before":
                # Write "<text>\r" at the paragraph start: the string becomes a new
                # paragraph and pushes the anchor's paragraph down. Clean for any
                # position, including the very start of the document (prepend).
                doc_com.Range(pstart, pstart).Text = unicodemath + "\r"
                ms = pstart
            elif pend >= doc_end:
                # The anchor's paragraph is the last; there's no position past the
                # undeletable terminal mark, so split "\r<text>" in just before it.
                pos = max(0, doc_end - 1)
                doc_com.Range(pos, pos).Text = "\r" + unicodemath
                ms = pos + 1
            else:
                # Open a new paragraph after the containing one and write into it.
                doc_com.Range(pend, pend).Text = unicodemath + "\r"
                ms = pend
            me = ms + _utf16_len(unicodemath)
            zone_rng = doc_com.Range(ms, me)
            zone_rng.OMaths.Add(zone_rng)
            zone = _equations.omath_in_range(doc_com, ms)
            if zone is not None:
                zone.BuildUp()
                zone.Type = 1 if display else 0
            index = _equation_index_at(doc_com, ms)
            self._style_equation_paragraph(ms, display=display)
        from ._equation_anchors import EquationAnchor  # lazy: _equation_anchors imports Anchor

        return EquationAnchor(self._doc, index)

    def _insert_equation_omml(
        self, omml_inner: str, *, where: str, display: bool
    ) -> EquationAnchor:
        """OMML path (latex/mathml): splice into a live template and InsertXML.

        `Range.InsertXML` only accepts a full, valid WordprocessingML package, so
        we take a live `Range.WordOpenXML` at a paragraph mark as the template and
        inject one math paragraph there. ``"after"`` targets the containing
        paragraph's mark; ``"before"`` targets the *preceding* paragraph's mark.
        Prepending before the first paragraph has no preceding mark to split
        against, so we open a leading paragraph first and trim the stray empty
        paragraph afterwards.
        """
        with _com.translate_com_errors():
            doc_com = self._doc.com
            doc_end = int(doc_com.Content.End)
            pstart, pend = self._equation_paragraph_span(where)
            prepend = where == "before" and pstart <= 0
            if prepend:
                doc_com.Range(0, 0).Text = "\r"
                t = 0
            elif where == "before":
                t = pstart - 1
            else:
                t = min(pend - 1, max(0, doc_end - 1))
            package = _equations.equation_package(
                str(doc_com.Range(t, t).WordOpenXML), omml_inner, display=display
            )
            doc_com.Range(t, t).InsertXML(package)
            if prepend and str(doc_com.Content.Text).startswith("\r"):
                # Trim the leading empty paragraph opened to anchor the prepend.
                doc_com.Range(0, 1).Delete()
            eq_pos = t if prepend else t + 1
            index = _equation_index_at(doc_com, eq_pos)
            self._style_equation_paragraph(eq_pos, display=display)
        from ._equation_anchors import EquationAnchor  # lazy: _equation_anchors imports Anchor

        return EquationAnchor(self._doc, index)

    def _ensure_equation_style(self) -> Any | None:
        """Return the COM ``Equation`` paragraph style, creating it if absent.

        A centred, ``Normal``-based paragraph style dedicated to display
        equations. Applying it to every display equation means an inserted
        equation can never inherit a heading style from its insertion point
        (which would drop the equation into the navigation outline / TOC), and
        gives a stable, named hook for future equation numbering and
        cross-references. Returns ``None`` for a degenerate document with no
        ``Normal`` to base it on — the caller then falls back to ``Normal``.
        """
        styles = self._doc.styles
        if "Equation" in styles:
            return styles["Equation"].com
        if "Normal" not in styles:
            return None
        style = styles.add("Equation", based_on="Normal", next_style="Normal")
        style.format_paragraph(alignment="center")
        return style.com

    def _style_equation_paragraph(self, pos: int, *, display: bool) -> None:
        """Pin the style/alignment of the paragraph an equation just landed on.

        Without this, an equation written at a paragraph boundary inherits the
        *following* paragraph's style — so an equation inserted before a
        ``Heading 2`` came out styled ``Heading 2`` and polluted the outline/TOC.
        A **display** equation gets the dedicated centred ``Equation`` style; an
        **inline** (``display=False``) equation is reset to ``Normal`` and
        left-aligned (it still lands on its own paragraph, but reads as body
        text, not centred display math). Best-effort — a COM hiccup here must not
        sink an otherwise-successful insert.
        """
        doc_com = self._doc.com
        try:
            para = doc_com.Range(pos, pos).Paragraphs(1).Range
            if display:
                eq_style = self._ensure_equation_style()
                if eq_style is not None:
                    para.Style = eq_style
                # Centring comes from the Equation style, so a redefined style
                # still drives it — no competing direct alignment.
            else:
                if "Normal" in self._doc.styles:
                    para.Style = self._doc.styles["Normal"].com
                para.ParagraphFormat.Alignment = int(WdParagraphAlignment.LEFT)
        except Exception:  # noqa: BLE001 — styling is a finishing touch, not the insert
            pass
