"""The `Anchor` abstract base: everything a concrete anchor can do to its range."""

from __future__ import annotations

import secrets
from abc import ABC, abstractmethod
from typing import TYPE_CHECKING, Any

from .. import _charts, _com, _equations, _images, _lists, _revisions, _shapes
from .._format import to_bgr, to_points
from ..constants import (
    MsoTextOrientation,
    MsoTriState,
    WdCaptionPosition,
    WdCollapseDirection,
    WdDropPosition,
    WdFieldType,
    WdIndexType,
    WdInformation,
    WdNumberType,
    WdParagraphAlignment,
    WdTabLeader,
)
from ..exceptions import EquationError, ExcelNotAvailableError, OpError

if TYPE_CHECKING:
    from pathlib import Path

    from .._document import Document
    from .._snapshot import Snapshot

from ._helpers import (
    _ALIGNMENT_NAMES,
    _BREAK_TYPES,
    _CC_LIST_TYPES,
    _CC_TYPE_NAMES,
    _DROP_POSITIONS,
    _FIELD_TYPES,
    _TAB_ALIGN,
    _TAB_LEADERS,
    _WRAP_NAMES,
    _WRAP_VALUES,
    _apply_font,
    _apply_paragraph_format,
    _chart_index_at,
    _coerce_highlight,
    _coerce_named,
    _equation_index_at,
    _final_paragraph_empty,
    _markdown_segments,
    _normalize_cc_items,
    _normalize_table_data,
    _read_font,
    _read_highlight,
    _read_paragraph_format,
    _resolve_wrap,
    _utf16_len,
    _validate_table_data,
    _within_table,
    apply_borders,
    range_text,
)
from ._refs import (
    _caption_above,
    _cross_ref_kind,
    _resolve_cross_ref_target,
)

if TYPE_CHECKING:
    from ._chart_anchors import ChartAnchor
    from ._content_controls import ContentControl
    from ._equation_anchors import EquationAnchor
    from ._range import RangeAnchor
    from ._shape_anchors import ShapeAnchor


# ---------------------------------------------------------------------------
# Base
# ---------------------------------------------------------------------------


class Anchor(ABC):
    """Abstract base — subclasses know how to materialise their COM Range.

    Concrete subclasses must implement `_range()` and `set_text()`. Other
    operations (`text`, `insert_before`, `insert_after`, `delete`,
    `apply_style`, `format_paragraph`) are derived and inherited as-is.
    """

    kind: str = "anchor"
    name: str = ""

    def __init__(self, doc: Document, name: str) -> None:
        self._doc = doc
        self.name = name

    @property
    def com(self) -> Any:
        """Raw COM range. Subclasses override."""
        return self._range()

    @abstractmethod
    def _range(self) -> Any:
        """Return the COM Range that this anchor refers to. Must be overridden."""

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

    @property
    def text(self) -> str:
        with _com.translate_com_errors():
            return range_text(self._range())

    def _revision_runs(self, start: int, end: int) -> list[dict[str, Any]]:
        """`{change, start, end, text}` for each insert/delete revision overlapping `[start, end)`."""
        runs: list[dict[str, Any]] = []
        for row in self._doc.revisions.list():
            change = row["type"]
            if change not in ("insert", "delete"):
                continue
            r_start, r_end = int(row["start"]), int(row["end"])
            if r_end <= start or r_start >= end:
                continue
            runs.append({"change": change, "start": r_start, "end": r_end, "text": row["text"]})
        return runs

    def revision_segments(self) -> list[dict[str, Any]]:
        """The anchor's text split into tracked-change segments (revision-aware read).

        Returns `[{text, change}, …]` in document order, where `change` is
        ``"insert"``, ``"delete"``, or ``None`` (unchanged). Word's `text` read
        shows the *final* view (inserted runs present, deleted runs gone); this
        also surfaces the deleted runs, so you can see both sides of a tracked
        edit. [`text_final`][wordlive.Anchor.text_final] and
        [`text_original`][wordlive.Anchor.text_original] are the two flattened
        views. The structured, whole-document counterpart is `doc.revisions`.
        """
        with _com.translate_com_errors():
            rng = self._range()
            start, end = int(rng.Start), int(rng.End)
            final_text = range_text(rng)
        return _revisions.segment_runs(final_text, start, self._revision_runs(start, end))

    @property
    def text_final(self) -> str:
        """The anchor's text **as if every tracked change in it were accepted**.

        Inserted runs stay, deleted runs drop — the after-the-edits view. Equal to
        `text` when nothing tracked touches the range. The mirror is
        [`text_original`][wordlive.Anchor.text_original]; the per-segment breakdown
        is [`revision_segments`][wordlive.Anchor.revision_segments].
        """
        return "".join(s["text"] for s in self.revision_segments() if s["change"] != "delete")

    @property
    def text_original(self) -> str:
        """The anchor's text **as if every tracked change in it were rejected**.

        Deleted runs stay, inserted runs drop — the before-the-edits view. The
        mirror of [`text_final`][wordlive.Anchor.text_final].
        """
        return "".join(s["text"] for s in self.revision_segments() if s["change"] != "insert")

    @property
    @abstractmethod
    def anchor_id(self) -> str:
        """Stable string identifier for this anchor (e.g. `bookmark:Address`).

        Each anchor kind has its own scheme (`bookmark:`, `cc:`, `heading:`),
        so subclasses must declare theirs explicitly — no useful default
        exists at this level.
        """

    @abstractmethod
    def set_text(self, text: str) -> None:
        """Replace the anchor's text in place. Must be overridden."""

    def insert_before(self, text: str) -> None:
        with _com.translate_com_errors():
            rng = self._range()
            insert_rng = self._doc.com.Range(rng.Start, rng.Start)
            insert_rng.Text = text

    def insert_after(self, text: str) -> None:
        with _com.translate_com_errors():
            rng = self._range()
            insert_rng = self._doc.com.Range(rng.End, rng.End)
            insert_rng.Text = text

    def insert_paragraph_before(self, text: str, style: str | None = None) -> None:
        """Insert a new paragraph immediately before this anchor's range.

        If `style` is given it must name a style defined in the document;
        otherwise `StyleNotFoundError` is raised before any text is inserted.
        """
        style_obj = self._doc.styles[style] if style is not None else None
        with _com.translate_com_errors():
            doc_com = self._doc.com
            start = int(self._range().Start)
            insert_rng = doc_com.Range(start, start)
            insert_rng.Text = text + "\r"
            if style_obj is not None:
                # Word measures Range offsets in UTF-16 code units; Python's
                # len() under-counts surrogate pairs and leaves the tail unstyled.
                styled = doc_com.Range(start, start + _utf16_len(text))
                styled.Style = style_obj.com

    def insert_paragraph_after(self, text: str, style: str | None = None) -> None:
        """Insert a new paragraph immediately after this anchor's range.

        If `style` is given it must name a style defined in the document;
        otherwise `StyleNotFoundError` is raised before any text is inserted.

        When the anchor is (or ends at) the document's final paragraph there is
        no position *after* the terminal paragraph mark to write to — Word
        rejects `Range(end, end)` there with a "value out of range" COM error.
        In that case the new paragraph is split in just before the final mark
        instead, so appending to the end of a document — the common
        "build from scratch" case, where the only paragraph *is* the last one —
        just works.
        """
        style_obj = self._doc.styles[style] if style is not None else None
        with _com.translate_com_errors():
            doc_com = self._doc.com
            end = int(self._range().End)
            doc_end = int(doc_com.Content.End)
            if end >= doc_end:
                # Anchor ends at the final paragraph mark. Insert "<break><text>"
                # just before that mark: the leading break terminates the
                # anchor's paragraph and `text` becomes a new final paragraph
                # (the original final mark now closes it).
                anchor_pos = max(0, doc_end - 1)
                insert_rng = doc_com.Range(anchor_pos, anchor_pos)
                insert_rng.Text = "\r" + text
                text_start = anchor_pos + 1
            else:
                insert_rng = doc_com.Range(end, end)
                insert_rng.Text = text + "\r"
                text_start = end
            if style_obj is not None:
                # Word measures Range offsets in UTF-16 code units; Python's
                # len() under-counts surrogate pairs and leaves the tail unstyled.
                styled = doc_com.Range(text_start, text_start + _utf16_len(text))
                styled.Style = style_obj.com

    def insert_block(self, items: list[Any], *, where: str = "after") -> RangeAnchor:
        """Insert a contiguous run of styled paragraphs at this anchor, atomically.

        The multi-paragraph counterpart to `insert_paragraph_after` — drop a
        whole styled section (a feature list, a set of bullets, a heading plus
        its body) in **one** op, in natural reading order. Inserting paragraphs
        one at a time forces a reverse-order dance to dodge positional-anchor
        renumbering; this places them all at a single point so order is just the
        order of `items`.

        Each item is one paragraph, given as either a plain string or a dict:

        - ``"some text"`` — sugar for ``{"text": "some text"}``.
        - ``{"text": "**Bold lead** — rest", "style": "List Bullet"}`` — `text`
          carries the tiny inline markdown (`**bold**`, `*italic*`,
          `***both***`; escape a literal asterisk as ``\\*``), and `style` names
          the paragraph style.
        - ``{"runs": [{"text": "Bold lead", "bold": true}, {"text": " — rest"}],
          "style": "List Bullet"}`` — the structured form: each run is
          ``{text, bold?, italic?, underline?, style?}`` (a per-run character
          style). Use it when markup is ambiguous or you need a run `style`.

        Returns a [`RangeAnchor`][wordlive.RangeAnchor] spanning the inserted
        block (`range:START-END`), so a follow-up op can target the whole run —
        e.g. `apply_list` it into a bulleted section, or comment on it. `where`
        is ``"after"`` (default) or ``"before"`` this anchor's range. Resolves
        every paragraph/run style up front, so an unknown style name raises
        `StyleNotFoundError` before any text is inserted. Wrap in `doc.edit(...)`
        for atomic undo. Raises `OpError` for a malformed `items` payload.
        """
        from .._runs import normalize_block_items, runs_to_text

        if where not in ("before", "after"):
            raise ValueError(f"where must be 'before' or 'after'; got {where!r}")
        norm = normalize_block_items(items)
        # Resolve every paragraph + run style before touching the document, so a
        # bad name fails the whole block cleanly rather than leaving a partial,
        # half-styled run behind.
        para_styles = [self._doc.styles[s] if s else None for _, s in norm]
        run_styles: dict[str, Any] = {}
        for runs, _ in norm:
            for r in runs:
                if r.style and r.style not in run_styles:
                    run_styles[r.style] = self._doc.styles[r.style]
        para_texts = [runs_to_text(runs) for runs, _ in norm]
        joined = "\r".join(para_texts)
        with _com.translate_com_errors():
            doc_com = self._doc.com
            if where == "before":
                start = int(self._range().Start)
                doc_com.Range(start, start).Text = joined + "\r"
            else:
                end = int(self._range().End)
                doc_end = int(doc_com.Content.End)
                final_mark = max(0, doc_end - 1)
                if end >= final_mark:
                    # We're at the document's terminal paragraph mark — the one
                    # position you can't write *past* (`doc.end` resolves here,
                    # as does an anchor ending at the last paragraph). The old
                    # `end >= doc_end` guard missed `doc.end` (whose range ends at
                    # doc_end - 1), so the block was written *before* the final
                    # mark and merged into a non-empty last paragraph — stealing
                    # its style too. Decide by whether that final paragraph holds
                    # text, so the block neither merges nor leaves a stray empty:
                    if _final_paragraph_empty(doc_com, final_mark):
                        # Reuse the empty final paragraph as the block's last one:
                        # write without a trailing break so the existing mark
                        # closes it (no leftover empty paragraph).
                        doc_com.Range(final_mark, final_mark).Text = joined
                        start = final_mark
                    else:
                        # Open a fresh paragraph after the final one: the leading
                        # break terminates it and the block becomes new trailing
                        # paragraphs (the original final mark closes the last one).
                        doc_com.Range(final_mark, final_mark).Text = "\r" + joined
                        start = final_mark + 1
                else:
                    doc_com.Range(end, end).Text = joined + "\r"
                    start = end
            span_end = start + _utf16_len(joined)
            # Paragraph styling and run formatting both preserve text length, so
            # offsets stay valid throughout: walk the block by deterministic
            # UTF-16 offset (Word counts code units) rather than re-querying
            # Paragraphs() after each mutation.
            off = start
            for (runs, _), style_obj, ptext in zip(norm, para_styles, para_texts, strict=True):
                plen = _utf16_len(ptext)
                if style_obj is not None:
                    doc_com.Range(off, off + plen).Paragraphs(1).Range.Style = style_obj.com
                roff = off
                for run in runs:
                    rlen = _utf16_len(run.text)
                    if run.formatted():
                        sub = doc_com.Range(roff, roff + rlen)
                        if run.bold is not None:
                            sub.Bold = bool(run.bold)
                        if run.italic is not None:
                            sub.Italic = bool(run.italic)
                        if run.underline is not None:
                            sub.Underline = 1 if run.underline else 0
                        if run.style:
                            sub.Style = run_styles[run.style].com
                    roff += rlen
                off += plen + 1  # + the paragraph mark (CR)
        from ._range import RangeAnchor  # lazy: _range imports Anchor

        return RangeAnchor(self._doc, start, span_end)

    def insert_section(
        self, heading: str, body: Any, *, level: int = 1, where: str = "after"
    ) -> RangeAnchor:
        """Insert a heading plus its body in one atomic op.

        The opinionated common case over `insert_block`: a single
        ``Heading {level}`` paragraph followed by `body`, placed in reading
        order at one point. `heading` carries the same inline markdown a block
        item's `text` does (`**bold**`, `*italic*`); `body` is the `insert_block`
        items shape — a list of plain strings or ``{text|runs, style?}`` dicts
        (a bare string is sugar for a one-paragraph body). `level` is 1–9 and
        selects the built-in ``Heading {level}`` style (validated before any
        mutation; an absent style raises `StyleNotFoundError` via `insert_block`).

        Returns the section's spanning [`RangeAnchor`][wordlive.RangeAnchor]
        (`range:START-END`). Wrap in `doc.edit(...)` for atomic undo.
        """
        if not isinstance(level, int) or isinstance(level, bool) or not 1 <= level <= 9:
            raise ValueError(f"level must be an integer 1–9; got {level!r}")
        if isinstance(body, str):
            body = [body]
        if not isinstance(body, list):
            raise OpError(
                f"insert_section body must be a string or list; got {type(body).__name__}"
            )
        items = [{"text": heading, "style": f"Heading {level}"}, *body]
        return self.insert_block(items, where=where)

    def insert_markdown(self, md: str, *, where: str = "after") -> RangeAnchor:
        """Insert a constrained-Markdown block as real Word structure, atomically.

        Maps a deliberately tiny block dialect (see `_markdown`) to paragraphs,
        headings, and lists: ``#``/``##``/``###`` → `Heading 1/2/3`, ``-``/``*``
        → a bulleted list, ``1.`` → a numbered list, blank-line-separated text →
        `Normal` paragraphs, with inline ``**bold**``/``*italic*`` spans honoured.
        It is **a subset, not CommonMark** — no code fences, nested lists, block
        quotes, or tables in v1; anything unrecognised is literal paragraph text.

        The whole block is one `insert_block` (one contiguous write); each
        same-kind list run is then `apply_list`-ed over its own span, so a
        numbered list reads 1..N. `where` is ``"after"`` (default) or ``"before"``
        this anchor's range. Returns the [`RangeAnchor`][wordlive.RangeAnchor]
        spanning everything inserted. Raises `OpError` for empty markdown.
        """
        from .._markdown import parse_markdown
        from .._runs import normalize_block_items, runs_to_text

        blocks = parse_markdown(md)
        if not blocks:
            raise OpError("insert_markdown requires non-empty markdown")
        # Flatten every block into ONE insert_block (a single contiguous write —
        # chaining separate inserts would land each list before the previous
        # block's paragraph mark and merge them). Record which paragraph runs are
        # lists so we can apply_list over their spans afterwards.
        segments = _markdown_segments(blocks)
        items: list[dict[str, Any]] = []
        list_groups: list[tuple[int, int, str]] = []  # (first_para, last_para, list_type)
        for seg_items, list_type in segments:
            start_idx = len(items)
            items.extend(seg_items)
            if list_type is not None:
                list_groups.append((start_idx, len(items) - 1, list_type))
        rng = self.insert_block(items, where=where)
        if not list_groups:
            return rng
        # Recompute each paragraph's offset exactly as insert_block walks them
        # (UTF-16 text length + one CR each, from the block's start), so a list
        # group's span can be addressed without re-querying the document.
        texts = [runs_to_text(runs) for runs, _ in normalize_block_items(items)]
        offsets: list[int] = []
        off = rng.start
        for t in texts:
            offsets.append(off)
            off += _utf16_len(t) + 1
        for first, last, list_type in list_groups:
            from ._range import RangeAnchor  # lazy: _range imports Anchor

            span = RangeAnchor(self._doc, offsets[first], offsets[last] + _utf16_len(texts[last]))
            span.apply_list(list_type)
        return rng

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

    def insert_table(
        self,
        rows: int | None = None,
        cols: int | None = None,
        *,
        where: str = "after",
        style: str | None = None,
        data: list[Any] | None = None,
        header: bool = False,
    ) -> Any:
        """Create a `rows` × `cols` table at this anchor and return it.

        The structural counterpart to `insert_image` — it *creates* new
        document structure rather than editing existing structure. Returns the
        new [`Table`][wordlive.Table] wrapper so create → fill → read closes on
        one object; the table's 1-based document index is on `.index`.

        `where` is ``"after"`` (default) or ``"before"`` this anchor's range —
        so `doc.headings["Pricing"].insert_table(...)` drops a table just under
        a heading, and `doc.end.insert_table(...)` (i.e.
        [`Document.add_table`][wordlive.Document.add_table]) appends one.

        `style` names a table style defined in the document (e.g. ``"Table
        Grid"``); an unknown name raises `StyleNotFoundError` before anything is
        inserted. `style=None` applies the built-in ``"Table Grid"`` when it's
        available, so a table has visible borders by default rather than the
        invisible cell gridlines of a styleless table.

        `data` populates the cells at creation and can be given two ways:

        - a **row-major 2-D list** (``[[r1c1, r1c2], …]``); or
        - **records** — a list of dicts (``[{"Item": "Travel", "Cost": "$400"},
          …]``), where the first record's keys become a header row and each
          dict a body row (so `header` is forced on). The natural shape for
          tabular data an LLM already has as rows of objects.

        When `data` is given, `rows`/`cols` are **optional** — they're inferred
        from the data's shape — so the common case is just
        ``end.insert_table(data=…)``. Pass them explicitly to pad the grid
        larger than the data; `data` is validated against the final `rows` ×
        `cols` up front (`OpError` on overflow) and a short payload leaves the
        trailing cells empty. Filling at creation keeps the whole grid in one
        atomic undo and beats a `set_cell` storm. With no `data`, both `rows`
        and `cols` are required.

        `header=True` bolds the first row as a header (records imply it). Wrap
        in `doc.edit(...)` for atomic undo. Raises `ValueError` for an unknown
        `where` and `OpError` for a non-positive `rows`/`cols`, a missing
        dimension with no data to infer it from, or a bad `data` shape.
        """
        from .._tables import Table, index_of

        if where not in ("before", "after"):
            raise ValueError(f"where must be 'before' or 'after'; got {where!r}")
        # Normalise data first so rows/cols can be inferred from its shape.
        grid: list[list[Any]] | None = None
        if data is not None:
            grid, header_from_data = _normalize_table_data(data)
            header = header or header_from_data
            if rows is None:
                rows = len(grid)
            if cols is None:
                cols = max((len(r) for r in grid), default=0)
        if rows is None or cols is None:
            raise OpError("insert_table needs rows and cols, or a data payload to infer them from")
        if isinstance(rows, bool) or not isinstance(rows, int) or rows < 1:
            raise OpError(f"table rows must be a positive integer; got {rows!r}")
        if isinstance(cols, bool) or not isinstance(cols, int) or cols < 1:
            raise OpError(f"table cols must be a positive integer; got {cols!r}")
        if grid is not None:
            _validate_table_data(grid, rows, cols)
        # Resolve the style up-front so a bad name fails before any mutation.
        if style is not None:
            style_obj = self._doc.styles[style]  # StyleNotFoundError (exit 2) if missing
        elif "Table Grid" in self._doc.styles:
            style_obj = self._doc.styles["Table Grid"]
        else:
            style_obj = None
        # New cells inherit the *paragraph* style at the insertion point — drop a
        # table right after a `Heading 2` and Word makes every cell Heading 2,
        # which renders as large heading text and pollutes the navigation
        # outline. Reset the cells to the body default (`Normal`) so a table
        # looks like a table regardless of where it was anchored. The table
        # `style` above (borders etc.) and `header` bolding still apply on top.
        normal_obj = self._doc.styles["Normal"] if "Normal" in self._doc.styles else None
        with _com.translate_com_errors():
            doc_com = self._doc.com
            rng = self._range()
            pos = int(rng.Start) if where == "before" else int(rng.End)
            # Word's final paragraph mark is undeletable and Tables.Add needs a
            # paragraph *after* the insertion point to anchor the table; at/after
            # that mark there is none, so the add raises COM 0x80020009. Push a
            # trailing paragraph first so the table lands before it (a document
            # can't end with a table anyway — Word keeps a paragraph after one).
            doc_end = int(doc_com.Content.End)
            if pos >= doc_end - 1:
                pos = max(0, doc_end - 1)
                doc_com.Range(pos, pos).Text = "\r"
            # Word merges two tables that touch with no paragraph mark between
            # them, so a table appended at the end (or dropped next to another)
            # would silently fuse into its neighbour. Push a separator paragraph
            # onto whichever side abuts an existing table; untouched insertions
            # into ordinary text get no stray paragraph.
            if _within_table(doc_com, pos - 1, pos):
                doc_com.Range(pos, pos).Text = "\r"
                pos += 1
            if _within_table(doc_com, pos, pos + 1):
                doc_com.Range(pos, pos).Text = "\r"
            insert_rng = doc_com.Range(pos, pos)
            table_com = doc_com.Tables.Add(insert_rng, rows, cols)
            if style_obj is not None:
                table_com.Style = style_obj.com
            if normal_obj is not None:
                # Per-cell rather than table_com.Range.Style: a paragraph style
                # set on the whole table range can bleed onto the paragraph that
                # follows the table; the cell loop is contained and explicit.
                normal_com = normal_obj.com
                for r in range(1, rows + 1):
                    for c in range(1, cols + 1):
                        table_com.Cell(r, c).Range.Style = normal_com
            if grid:
                for r, row in enumerate(grid, start=1):
                    for c, val in enumerate(row, start=1):
                        table_com.Cell(r, c).Range.Text = str(val)
            if header:
                table_com.Rows(1).Range.Bold = True
            index = index_of(self._doc.com, table_com)
        return Table(self._doc, table_com, index)

    def insert_break(self, kind: str = "page", *, where: str = "after") -> None:
        """Insert a page, column, or section break at this anchor.

        The explicit one-off break — the clean alternative to appending a
        paragraph whose text is a literal form-feed. `kind` is one of:

        - ``"page"`` (default) — a manual page break (the 90% case).
        - ``"column"`` — a column break (multi-column layouts).
        - ``"section_next"`` — a section break that starts the new section on
          the next page.
        - ``"section_continuous"`` — a section break with no page break, so the
          new section flows on the same page.

        Section breaks pair with [`Document.sections`][wordlive.Document.sections]:
        each new section gets its own headers/footers and page setup. To make a
        *style* (e.g. every `Heading 1`) open a new page without a stray break
        character, prefer
        [`format_paragraph(page_break_before=True)`][wordlive.Anchor.format_paragraph]
        instead — it survives reflow.

        `where` is ``"after"`` (default) or ``"before"`` this anchor's range.
        Wrap in `doc.edit(...)` for atomic undo. Raises `ValueError` for an
        unknown `kind` or `where`.
        """
        if kind not in _BREAK_TYPES:
            raise ValueError(f"unknown break kind {kind!r}; expected one of {sorted(_BREAK_TYPES)}")
        if where not in ("before", "after"):
            raise ValueError(f"where must be 'before' or 'after'; got {where!r}")
        break_type = _BREAK_TYPES[kind]
        # A section break creates a *new* paragraph to carry the break, and that
        # paragraph inherits the anchor's style — drop one before a `Heading 1`
        # and Word makes the break paragraph a heading, leaving a spurious empty
        # entry in the navigation outline / TOC. Reset it to `Normal` so the break
        # is invisible to the outline. (Page/column breaks are an in-paragraph
        # character and create no such paragraph, so they need no reset.)
        is_section = kind in ("section_next", "section_continuous")
        normal_obj = (
            self._doc.styles["Normal"] if is_section and "Normal" in self._doc.styles else None
        )
        with _com.translate_com_errors():
            rng = self._range()
            pos = int(rng.Start) if where == "before" else int(rng.End)
            insert_rng = self._doc.com.Range(pos, pos)
            insert_rng.InsertBreak(Type=int(break_type))
            if normal_obj is not None:
                # The break now occupies the position we inserted at; the
                # paragraph containing `pos` is the break paragraph.
                break_para = self._doc.com.Range(pos, pos).Paragraphs(1)
                break_para.Range.Style = normal_obj.com

    def insert_field(self, kind: str, *, text: str | None = None, where: str = "after") -> None:
        """Insert a Word field at this anchor — a self-updating value, not literal text.

        A field shows a computed value Word keeps current: a page number, the
        page count, today's date, the file name, a document property. The named
        kinds are:

        - ``"page"`` — the current page number (`{ PAGE }`).
        - ``"numpages"`` — the total page count (`{ NUMPAGES }`); pair with
          ``"page"`` for "Page X of Y".
        - ``"date"`` / ``"time"`` — the current date / time.
        - ``"filename"`` — the document's file name.
        - ``"author"`` / ``"title"`` — document-property fields.

        For anything else, ``kind="field"`` is the escape hatch: pass the raw
        field code as `text` (e.g.
        ``insert_field("field", text="REF myBookmark \\\\h")``) and Word inserts an
        empty field carrying that code.

        Page numbers belong in a header or footer — because a `HeaderFooter`
        *is* an anchor, ``doc.sections[1].footer().insert_field("page")`` works,
        and [`HeaderFooter.insert_page_number()`][wordlive.HeaderFooter] is the
        sugar for it. Newly inserted fields render once; call
        [`Document.update_fields()`][wordlive.Document] (or take a `snapshot`,
        which repaginates) to refresh them after later edits.

        `where` is ``"after"`` (default) or ``"before"`` this anchor's range.
        Bad input raises `OpError`. Wrap in `doc.edit(...)` for atomic undo.
        """
        try:
            if where not in ("before", "after"):
                raise ValueError(f"where must be 'before' or 'after'; got {where!r}")
            wd_type = _coerce_named(kind, _FIELD_TYPES, "field kind")
            if wd_type == int(WdFieldType.EMPTY) and not text:
                raise ValueError(
                    'field kind "field" requires the raw field code via text= (e.g. text="PAGE")'
                )
            with _com.translate_com_errors():
                # Collapse a *duplicate* of the anchor's own range, so the field
                # lands in the same story — critical for header/footer anchors,
                # whose offsets are not main-document positions (a `doc.Range`
                # there would target the body instead).
                insert_rng = self._range().Duplicate
                insert_rng.Collapse(
                    int(WdCollapseDirection.START if where == "before" else WdCollapseDirection.END)
                )
                # Positional args: the Type=/Text= keywords are dropped under
                # pywin32 late binding (same gotcha as TabStops.Add / Footnotes).
                if text is not None:
                    insert_rng.Fields.Add(insert_rng, wd_type, text)
                else:
                    insert_rng.Fields.Add(insert_rng, wd_type)
        except (ValueError, TypeError) as e:
            raise OpError(str(e)) from e

    def insert_footnote(self, text: str, *, where: str = "after") -> Any:
        """Insert a footnote at this anchor and return it as a `Footnote` anchor.

        A footnote drops a reference mark in the main text and puts `text` in the
        note body at the bottom of the page; Word auto-numbers the mark. The
        returned [`Footnote`][wordlive.Footnote] is addressed `footnote:N`, so
        `note.set_text(...)` edits the body and `note.delete()` removes the mark
        and body together. Discover existing footnotes with
        [`doc.footnotes`][wordlive.Document.footnotes].

        `where` is ``"after"`` (default) or ``"before"`` this anchor's range —
        the side the reference mark lands on. Wrap in `doc.edit(...)` for atomic
        undo. Bad input raises `OpError`.
        """
        return self._insert_note("Footnotes", "footnote", text, where=where)

    def insert_endnote(self, text: str, *, where: str = "after") -> Any:
        """Insert an endnote at this anchor and return it as an `Endnote` anchor.

        The endnote mirror of [`insert_footnote`][wordlive.Anchor.insert_footnote]:
        the reference mark lands in the main text and `text` collects at the end
        of the document (or section). The returned
        [`Endnote`][wordlive.Endnote] is addressed `endnote:N`; discover existing
        endnotes with [`doc.endnotes`][wordlive.Document.endnotes].

        `where` is ``"after"`` (default) or ``"before"`` this anchor's range.
        Wrap in `doc.edit(...)` for atomic undo. Bad input raises `OpError`.
        """
        return self._insert_note("Endnotes", "endnote", text, where=where)

    def _insert_note(self, attr: str, scheme: str, text: str, *, where: str) -> Any:
        """Shared footnote/endnote insertion (`attr` is the COM collection name)."""
        from .._notes import Endnote, Footnote, index_of_note

        cls = Footnote if scheme == "footnote" else Endnote
        try:
            if where not in ("before", "after"):
                raise ValueError(f"where must be 'before' or 'after'; got {where!r}")
            with _com.translate_com_errors():
                rng = self._range()
                # A note's reference mark always lands in the main text story, so
                # a plain document Range at the anchor's edge is correct (unlike
                # insert_field, which can target a footer's own story).
                pos = int(rng.Start) if where == "before" else int(rng.End)
                insert_rng = self._doc.com.Range(pos, pos)
                coll = getattr(self._doc.com, attr)
                # Positional args: an empty Reference auto-numbers, and the
                # Reference=/Text= keywords are dropped under pywin32 late binding
                # (same gotcha as Fields.Add / TabStops.Add).
                coll.Add(insert_rng, "", text)
                index = index_of_note(coll, pos)
            return cls(self._doc, index)
        except (ValueError, TypeError) as e:
            raise OpError(str(e)) from e

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

    def insert_content_control(
        self,
        kind: str = "rich_text",
        *,
        title: str | None = None,
        tag: str | None = None,
        items: list[Any] | None = None,
        where: str = "wrap",
        lock_contents: bool = False,
        lock_control: bool = False,
    ) -> ContentControl:
        """Wrap (or insert) a content control at this anchor and return it.

        Content controls are the building blocks of form-like documents — a
        labelled region the user fills in. `kind` selects the type:
        ``"rich_text"`` (the default — formatted text), ``"text"`` (plain text),
        ``"picture"``, ``"combo_box"`` / ``"dropdown"`` (a pick list — pass
        `items`), ``"date"``, ``"checkbox"`` (Word 2013+), ``"building_block"``,
        ``"group"``, or ``"repeating_section"`` (Word 2013+).

        `where` is ``"wrap"`` (default — the control surrounds this anchor's
        existing range, so a `range:START-END` from `find` wraps that phrase) or
        ``"before"`` / ``"after"`` (insert a fresh empty control at the anchor's
        start / end). Set `title` and/or `tag` to give the control a name: a
        titled control is addressable later as `cc:TITLE` (falling back to the
        tag) and shows a label in Word's UI. `items` populates a combo box or
        dropdown — each is a string, or a `{"text": ..., "value": ...}` dict.
        `lock_contents` stops the user editing the value; `lock_control` stops
        them deleting the control.

        Returns the new [`ContentControl`][wordlive.ContentControl] (usable even
        when unnamed — it caches the live control). Wrap in `doc.edit(...)` for
        atomic undo. Bad input raises `OpError`.
        """
        try:
            if where not in ("wrap", "before", "after"):
                raise ValueError(f"where must be 'wrap', 'before', or 'after'; got {where!r}")
            try:
                cc_type = _CC_TYPE_NAMES[str(kind).lower()]
            except (KeyError, AttributeError) as e:
                raise ValueError(
                    f"unknown content control kind {kind!r}; "
                    f"expected one of {sorted(_CC_TYPE_NAMES)}"
                ) from e
            if items is not None and cc_type not in _CC_LIST_TYPES:
                raise ValueError("items is only valid for a 'combo_box' or 'dropdown' control")
            with _com.translate_com_errors():
                target = self._range().Duplicate
                if where != "wrap":
                    target.Collapse(
                        int(
                            WdCollapseDirection.START
                            if where == "before"
                            else WdCollapseDirection.END
                        )
                    )
                # Positional args (Type, Range) for late-binding safety.
                cc = self._doc.com.ContentControls.Add(int(cc_type), target)
                if title is not None:
                    cc.Title = str(title)
                if tag is not None:
                    cc.Tag = str(tag)
                if lock_contents:
                    cc.LockContents = True
                if lock_control:
                    cc.LockContentControl = True
                if items:
                    entries = cc.DropdownListEntries
                    for entry_text, value in _normalize_cc_items(items):
                        entries.Add(entry_text, value)  # positional (Text, Value)
            name = title or tag or ""
            from ._content_controls import ContentControl  # lazy: _content_controls imports Anchor

            return ContentControl(self._doc, name, com=cc)
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

    def snapshot(
        self, out: str | Path | None = None, *, dpi: int = 150, max_dim: int | None = None
    ) -> list[Snapshot]:
        """Render the page(s) this anchor sits on to PNG — let a model *see* it.

        A heading expands to its whole section; any other anchor renders the
        page(s) its range spans. Returns a list of
        [`Snapshot`][wordlive.Snapshot] (one per page); pass `out` to also write
        the image(s) to disk. `max_dim` caps each page's long edge in pixels (for
        a cheaper render). Sugar for
        [`Document.snapshot_anchor`][wordlive.Document.snapshot_anchor]; see it
        for the full semantics. Requires the `snapshot` extra (PyMuPDF).
        """
        return self._doc.snapshot_anchor(self, out, dpi=dpi, max_dim=max_dim)

    def read_image(self) -> tuple[bytes, str]:
        """Extract the image embedded in this anchor's range as `(bytes, mime_type)`.

        The read side of the image story — pull an embedded picture's original
        bytes back out (e.g. to hand to a vision model), the counterpart to
        [`insert_image`][wordlive.Anchor.insert_image]. The range must contain
        exactly one picture: an [`image:N`][wordlive.ImageAnchor] anchor (or any
        single-image text anchor) reads cleanly, while a range with no image — or
        more than one — raises `ImageSourceError`. `bytes` is the picture's raw
        encoded data (PNG/JPEG/…); `mime_type` is its content type
        (``"image/png"``, ``"image/jpeg"``, …). Discover what's there first with
        [`doc.images`][wordlive.Document.images]. Read-only — nothing is mutated.
        """
        with _com.translate_com_errors():
            return _images.read_image_from_range(self._range())

    def delete(self) -> None:
        with _com.translate_com_errors():
            self._range().Delete()

    def apply_style(self, name: str) -> None:
        """Apply the named paragraph or character style to this anchor's range.

        Word selects paragraph- vs. character-style behaviour from the style's
        own `Type`; we don't model that distinction. Raises `StyleNotFoundError`
        if the style isn't defined in the document.
        """
        style = self._doc.styles[name]  # raises StyleNotFoundError if missing
        with _com.translate_com_errors():
            self._range().Style = style.com

    def format_paragraph(
        self,
        *,
        alignment: Any = None,
        left_indent: float | None = None,
        right_indent: float | None = None,
        first_line_indent: float | None = None,
        space_before: float | None = None,
        space_after: float | None = None,
        line_spacing: Any = None,
        page_break_before: bool | None = None,
        keep_together: bool | None = None,
        keep_with_next: bool | None = None,
        widow_control: bool | None = None,
    ) -> None:
        """Set paragraph-formatting properties on this anchor's range.

        All kwargs are optional; only the ones explicitly passed are written.
        Indent and spacing values are in points (Word's native unit for
        `ParagraphFormat.LeftIndent` etc.). `alignment` accepts a
        `WdParagraphAlignment` enum, its int value, or a string
        (`"left"`/`"center"`/`"right"`/`"justify"`).

        `line_spacing` sets the leading between lines *within* the paragraph
        (distinct from `space_before`/`space_after`, which space paragraphs
        apart). It accepts a **number** — a multiple of single spacing (`1`
        single, `1.5`, `2` double) — one of the keywords `"single"`/`"1.5"`/
        `"double"`, or an **exact length string** (`"14pt"`, `"1.5cm"`) for a
        fixed line height.

        `page_break_before=True` forces the paragraph to begin on a new page —
        the *clean* way to page-break (e.g. apply it to every `Heading 1`): it's
        a paragraph property that survives reflow and leaves no stray break
        character, unlike [`insert_break`][wordlive.Anchor.insert_break].
        `False` clears the property. Indents/spacing accept a number (points) or
        a unit string (`"0.5in"`).

        The remaining flags are Word's *pagination* controls (all tri-state —
        `True`/`False` set, `None` leaves untouched), for clean multi-page
        layout: `keep_together` keeps every line of the paragraph on one page;
        `keep_with_next` keeps it on the same page as the following paragraph
        (e.g. a heading with its first body line); `widow_control` prevents a
        lone first/last line stranded at the bottom/top of a page (on by default
        in Word).
        """
        try:
            with _com.translate_com_errors():
                _apply_paragraph_format(
                    self._range().ParagraphFormat,
                    alignment=alignment,
                    left_indent=left_indent,
                    right_indent=right_indent,
                    first_line_indent=first_line_indent,
                    space_before=space_before,
                    space_after=space_after,
                    line_spacing=line_spacing,
                    page_break_before=page_break_before,
                    keep_together=keep_together,
                    keep_with_next=keep_with_next,
                    widow_control=widow_control,
                )
        except (ValueError, TypeError) as e:
            raise OpError(str(e)) from e

    def drop_cap(
        self,
        lines: int = 3,
        *,
        position: str = "dropped",
        distance: Any = 0.0,
        font: str | None = None,
    ) -> None:
        """Turn the first letter of this anchor's paragraph into a drop cap.

        The editorial oversized initial — a real Word `DropCap`, not a faked
        big-font run, so it reflows and re-wraps the body text around it
        natively. Applies to the **first paragraph** of the anchor's range.

        `position` is ``"dropped"`` (the default — the letter sits *into* the
        text, the common magazine style), ``"margin"`` (it hangs out in the left
        margin), or ``"none"`` (remove an existing drop cap; `lines`/`distance`/
        `font` are then ignored). `lines` is how many lines tall the letter is
        (Word's default is 3). `distance` is the gap between the letter and the
        body text, in points (or a unit string like ``"2pt"``). `font` optionally
        sets the dropped letter's font family.

        Word rejects a drop cap on an **empty** paragraph (there's no letter to
        drop) — that surfaces as a `ComError`. Wrap in `doc.edit(...)` for atomic
        undo. Raises `OpError` for an unknown `position` or a bad `distance`.
        """
        try:
            pos = _coerce_named(position, _DROP_POSITIONS, "drop-cap position")
            dist = to_points(distance)
            if not isinstance(lines, int) or isinstance(lines, bool) or lines < 1:
                raise ValueError(f"lines must be a positive integer; got {lines!r}")
            with _com.translate_com_errors():
                dc = self._range().Paragraphs(1).DropCap
                # Enable the cap first: Word resets LinesToDrop/DistanceFromText/
                # FontName to its defaults when Position changes, so the geometry
                # must be written *after* the position or it's silently dropped.
                dc.Position = pos
                if pos == int(WdDropPosition.NONE):
                    return
                dc.LinesToDrop = lines
                dc.DistanceFromText = dist
                if font is not None:
                    dc.FontName = str(font)
        except (ValueError, TypeError) as e:
            raise OpError(str(e)) from e

    def format_run(
        self,
        *,
        bold: bool | None = None,
        italic: bool | None = None,
        underline: bool | None = None,
        strikethrough: bool | None = None,
        font: str | None = None,
        size: Any = None,
        color: Any = None,
        highlight: Any = None,
        subscript: bool | None = None,
        superscript: bool | None = None,
        small_caps: bool | None = None,
        all_caps: bool | None = None,
        spacing: Any = None,
    ) -> None:
        """Set character-formatting (run-level) properties on this anchor's range.

        Direct formatting — the *bold this phrase* layer, distinct from
        [`apply_style`][wordlive.Anchor.apply_style] (named styles) and
        [`format_paragraph`][wordlive.Anchor.format_paragraph] (paragraph-scope).
        Pairs naturally with `range:START-END` to style a sub-paragraph span.

        All kwargs are optional and tri-state; only the ones explicitly passed
        are written (`None` leaves the property untouched). `bold`/`italic`/
        `underline`/`strikethrough`/`subscript`/`superscript`/`small_caps`/
        `all_caps` are booleans. `font` is a family name; `size` and `spacing`
        accept a number (points) or a unit string (`"12pt"`, `"1.5mm"`).
        `color` accepts a named colour, hex (`"#FF0000"`), or `(r, g, b)`.
        `highlight` is a named text-highlight colour (`"yellow"`, `"green"`, …,
        or `"none"`/`"auto"` to clear it) — a palette index, *not* an RGB.

        Bad colour/length/highlight input raises `OpError` (bad-input). Wrap in
        `doc.edit(...)` for atomic undo.
        """
        try:
            with _com.translate_com_errors():
                rng = self._range()
                _apply_font(
                    rng.Font,
                    bold=bold,
                    italic=italic,
                    underline=underline,
                    strikethrough=strikethrough,
                    font_name=font,
                    size=size,
                    color=color,
                    subscript=subscript,
                    superscript=superscript,
                    small_caps=small_caps,
                    all_caps=all_caps,
                    spacing=spacing,
                )
                if highlight is not None:
                    rng.HighlightColorIndex = _coerce_highlight(highlight)
        except (ValueError, TypeError) as e:
            raise OpError(str(e)) from e

    def format_info(self) -> dict[str, Any]:
        """The effective paragraph + character formatting on this anchor — the
        read mirror of [`format_paragraph`][wordlive.Anchor.format_paragraph] and
        [`format_run`][wordlive.Anchor.format_run]. Pure read.

        Returns `{anchor_id, style, paragraph, font}`. `style` is the applied
        paragraph style's name. `paragraph` and `font` carry one entry per field,
        each `{value, style, override}`:

        - `value` — the *effective* value (what's actually rendered);
        - `style` — the value the applied **style** would give on its own;
        - `override` — `True` when `value != style`, i.e. a **direct override**
          sits on top of the style (the signal the consistency linter rules act
          on). A mixed field (`value is None`) is never flagged as an override.

        `font.mixed` lists the character fields that read `wdUndefined` because
        they vary across the range's runs (e.g. a heading with one bold word) —
        those carry `value: null` rather than a bogus number. Lengths are in
        points; `color` is `#RRGGBB` (or `"auto"`); `alignment`/`line_spacing`
        use the same keywords the write verbs accept. `font.hidden` flags Word's
        hidden-text attribute. `font.highlight` is a highlight keyword (`"yellow"`,
        … or `"none"`); it lives on the range, not the style, so it's
        effective-only — `style` is always `null` and `override` just means a
        highlight is present.

        The field vocabulary is identical to the write side, so a value read here
        can be written straight back through `format_paragraph`/`format_run`.
        """
        with _com.translate_com_errors():
            rng = self._range()
            style = rng.ParagraphStyle
            eff_para = _read_paragraph_format(rng.ParagraphFormat)
            sty_para = _read_paragraph_format(style.ParagraphFormat)
            eff_font, mixed = _read_font(rng.Font)
            sty_font, _ = _read_font(style.Font)
            highlight = _read_highlight(rng.HighlightColorIndex)
            style_name = str(style.NameLocal)

        def _annotate(eff: dict[str, Any], sty: dict[str, Any]) -> dict[str, Any]:
            return {
                key: {
                    "value": eff[key],
                    "style": sty[key],
                    "override": eff[key] is not None and eff[key] != sty[key],
                }
                for key in eff
            }

        font = _annotate(eff_font, sty_font)
        # Highlight lives on the Range, not the Font, and a style never carries it
        # (see `_STYLE_RUN_FIELDS`), so it's effective-only: no style baseline, and
        # an "override" simply means a highlight is present. A mixed read (some runs
        # highlighted) surfaces via `mixed`, like the other character fields.
        if highlight is None:
            mixed.append("highlight")
        font["mixed"] = mixed
        font["highlight"] = {
            "value": highlight,
            "style": None,
            "override": highlight is not None and highlight != "none",
        }
        return {
            "anchor_id": self.anchor_id,
            "style": style_name,
            "paragraph": _annotate(eff_para, sty_para),
            "font": font,
        }

    def set_shading(self, *, fill: Any = None, pattern: Any = None) -> None:
        """Set the background (fill) shading of this anchor's range.

        `fill` is a named colour, hex (`"#FFFF00"`), or `(r, g, b)` — applied to
        `Range.Shading.BackgroundPatternColor`. Because a `Cell` is an `Anchor`,
        this is also how you shade a table cell. `pattern` (a shading pattern/
        texture) is accepted for forward-compatibility but not yet applied —
        deferred. Bad colour input raises `OpError`. Wrap in `doc.edit(...)`.
        """
        try:
            with _com.translate_com_errors():
                if fill is not None:
                    self._range().Shading.BackgroundPatternColor = to_bgr(fill)
        except (ValueError, TypeError) as e:
            raise OpError(str(e)) from e

    def set_borders(
        self,
        *,
        sides: Any = "all",
        style: Any = "single",
        weight: Any = 0.5,
        color: Any = None,
    ) -> None:
        """Draw borders on this anchor's range (or cell — a `Cell` is an `Anchor`).

        `sides` is `"all"`/`"box"` (the default — four outer edges), a single
        edge (`"top"`/`"bottom"`/`"left"`/`"right"`), an interior gridline
        (`"horizontal"`/`"vertical"`, for multi-cell ranges), or a list of those.
        `style` is a line style (`"single"`, `"double"`, `"dot"`, `"dash"`, …, or
        `"none"` to remove). `weight` is the line width in points, snapped to
        Word's discrete set (0.25/0.5/0.75/1/1.5/2.25/3 pt). `color` is an
        optional border colour (name/hex/RGB).

        This sets per-range / per-cell borders. Page borders
        (`Section.Borders`) are out of scope; whole-table borders (the entire
        grid in one call, including interior gridlines) go through
        [`Table.set_borders`][wordlive.Table.set_borders] / the `table
        set-borders` verb. Bad input raises `OpError`. Wrap in `doc.edit(...)`.
        """
        try:
            with _com.translate_com_errors():
                apply_borders(
                    self._range().Borders,
                    sides=sides,
                    style=style,
                    weight=weight,
                    color=color,
                )
        except (ValueError, TypeError) as e:
            raise OpError(str(e)) from e

    def add_tab_stop(self, position: Any, *, align: Any = "left", leader: Any = None) -> None:
        """Add a tab stop to this anchor's paragraph(s).

        `position` is the distance from the left margin in points (or a unit
        string like `"3in"`). `align` is `"left"`/`"center"`/`"right"`/
        `"decimal"`/`"bar"`. `leader` is an optional fill drawn up to the stop —
        `"dots"` (price lists / tables of contents), `"dashes"`, `"lines"`, … —
        defaulting to none. Maps to `ParagraphFormat.TabStops.Add`. Bad input
        raises `OpError`. Wrap in `doc.edit(...)`.
        """
        try:
            pos = to_points(position)
            al = _coerce_named(align, _TAB_ALIGN, "tab alignment")
            ld = (
                _coerce_named(leader, _TAB_LEADERS, "tab leader")
                if leader is not None
                else int(WdTabLeader.SPACES)
            )
            with _com.translate_com_errors():
                # Positional args: the `Leader=` keyword is dropped under pywin32
                # late binding, so pass Position, Alignment, Leader positionally.
                self._range().ParagraphFormat.TabStops.Add(pos, al, ld)
        except (ValueError, TypeError) as e:
            raise OpError(str(e)) from e

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

    def location(self) -> dict[str, Any]:
        """Where this anchor sits in the laid-out document — a pure read.

        Returns `{page, end_page, line, column, in_table}`:

        - `page` / `end_page` — the 1-based pages the anchor's **first** and
          **last** characters fall on (equal for a collapsed/single-line anchor);
          the pair is the anchor's *page span*, so a section/table/image that
          straddles a page boundary reports both. `page` is what answers "what
          page is this on"; scan `paragraphs` and watch `page` step up to find
          "which paragraph starts page 2".
        - `line` / `column` — the first character's 1-based line and column in
          the page's text grid (`Range.Information`).
        - `in_table` — whether the anchor sits inside a table.

        Page/line numbers are only meaningful in print layout, so the document
        is **repaginated first** (content-neutral — it touches neither the
        user's selection, scroll, nor view), mirroring the guarantee a
        `snapshot` gives. No politeness concern: this mutates nothing — the
        document's `Saved` state is snapshotted and restored around the
        repaginate, which would otherwise flip Word's dirty bit.
        """
        with _com.translate_com_errors(), _com.preserve_saved(self._doc.com):
            rng = self._range()
            self._doc.com.Repaginate()
            start, end = int(rng.Start), int(rng.End)
            doc_com = self._doc.com
            head = doc_com.Range(start, start)
            tail = doc_com.Range(end, end)
            return {
                "page": int(head.Information(int(WdInformation.ACTIVE_END_PAGE_NUMBER))),
                "end_page": int(tail.Information(int(WdInformation.ACTIVE_END_PAGE_NUMBER))),
                "line": int(head.Information(int(WdInformation.FIRST_CHARACTER_LINE_NUMBER))),
                "column": int(head.Information(int(WdInformation.FIRST_CHARACTER_COLUMN_NUMBER))),
                "in_table": bool(rng.Information(int(WdInformation.WITH_IN_TABLE))),
            }

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

    def __repr__(self) -> str:
        return f"<{type(self).__name__} {self.name!r}>"
