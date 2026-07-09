"""Save/export, pins, watermark, track-changes, bibliography style."""

from __future__ import annotations

from collections.abc import Iterator
from contextlib import contextmanager
from pathlib import Path
from typing import TYPE_CHECKING, Any

from .. import _com, _snapshot
from .._anchors import (
    _WL_PREFIX,
    _bookmarks_including_hidden,
    _mint_wl_bookmark,
    _new_pin_code,
    _pin_id_for,
    _pin_name_for,
    _validate_pin_slug,
)
from .._format import to_bgr
from ..constants import (
    MsoPresetTextEffect,
    MsoTriState,
    WdHeaderFooterIndex,
    WdRelativeHorizontalPosition,
    WdRelativeVerticalPosition,
    WdSaveFormat,
    WdShapePosition,
    WdWrapSideType,
    WdWrapType,
)
from ..exceptions import (
    OpError,
)

if TYPE_CHECKING:
    from .._anchors import Anchor

from ._core import DocumentCore, WatermarkInfo, _resolve_level_band


class PersistenceMixin(DocumentCore):
    """Save/export, pins, watermark, track-changes, bibliography style."""

    def save(self) -> str:
        """Save the document to its existing file, returning the absolute path.

        Raises [`OpError`][wordlive.OpError] if the document has never
        been saved (it has no path yet) — call [`save_as`][wordlive.Document.save_as]
        first. This is the **ungated** Python-API surface: it writes wherever the
        document already lives. The CLI / MCP `save` verb additionally checks that
        path against the configured save-directory whitelist before calling this.
        """
        with _com.translate_com_errors():
            folder = str(self._doc.Path)
            if not folder:
                raise OpError("document has never been saved; use save_as(path) first")
            self._doc.Save()
            return str(self._doc.FullName)

    def save_as(self, path: str | Path, *, fmt: str = "docx", overwrite: bool = False) -> str:
        """Save the document to `path`, returning the absolute path written.

        `fmt` is `"docx"` (the modern Open XML format). For PDF, use
        [`export_pdf`][wordlive.Document.export_pdf] (it goes through a different
        COM call and takes a page range). By default refuses to clobber an
        existing file — pass `overwrite=True` to allow it. **Ungated** like
        [`save`][wordlive.Document.save]; the CLI / MCP surface whitelists the
        target first.
        """
        target = Path(path).expanduser()
        fmt_norm = str(fmt).lower().lstrip(".")
        if fmt_norm == "pdf":
            raise OpError("save_as does not write PDF; use export_pdf(path) instead")
        if fmt_norm not in ("docx",):
            raise OpError(f"unsupported save format {fmt!r}; supported: docx (PDF via export_pdf)")
        if not overwrite and target.exists():
            raise OpError(
                f"refusing to overwrite existing file {str(target)!r}; pass overwrite=True"
            )
        abspath = str(target.resolve())
        with _com.translate_com_errors():
            self._doc.SaveAs2(FileName=abspath, FileFormat=int(WdSaveFormat.DOCUMENT_DEFAULT))
        return abspath

    def export_pdf(
        self, path: str | Path, *, from_page: int | None = None, to_page: int | None = None
    ) -> str:
        """Export the document (or a page span) to a PDF at `path`; return the path.

        `from_page` / `to_page` are 1-based and inclusive; omit both to export the
        whole document, or give `from_page` alone to export a single page. Goes
        through `Document.ExportAsFixedFormat` (the same engine
        [`snapshot`][wordlive.Document.snapshot] uses), so the PDF is a
        pixel-faithful render — the recommended "hand back a deliverable" path.
        Overwrites an existing file. **Ungated** like [`save`][wordlive.Document.save].
        """
        abspath = str(Path(path).expanduser().resolve())
        _snapshot._export_pdf(self._doc, abspath, from_page=from_page, to_page=to_page)
        return abspath

    def pin(self, anchor: Anchor | str, name: str | None = None) -> dict[str, Any]:
        """Plant a durable handle on `anchor`'s range and return its `pin:` id.

        The fix for fragile positional ids: `pin("para:7")` mints a hidden
        bookmark over that paragraph's range and hands back a `pin:<code>` anchor
        id that keeps pointing at the same content across later inserts / deletes
        / edits (Word maintains the association natively — that's the durability).
        Resolve it like any anchor — `doc.anchor_by_id("pin:a3f9c2")` — or feed it
        straight into another op. If the pinned content is later deleted the handle
        correctly vanishes (resolving it raises `AnchorNotFoundError`).

        `anchor` is an [`Anchor`][wordlive.Anchor] or an anchor id string. `name`
        optionally gives a readable slug (``budget-intro`` -> ``pin:budget-intro``;
        lowercase words joined by single hyphens); omit it for a random code.
        Re-using a slug moves the handle to the new range (Word's `Bookmarks.Add`
        semantics). Editing *through* the pin (`set_text`) keeps it; rewriting the
        same span via a different anchor's `Range.Text` drops it.

        Returns `{"anchor_id": "pin:…", "pin": "pin:…", "target": <resolved id>}`.
        `stamp` is an alias. Wrap in `doc.edit(...)` for atomic undo — but do not
        call it inside an already-open edit scope (custom undo records don't nest;
        the `exec` batch already owns one). The CLI verb is
        `wordlive pin ANCHOR_ID [--name SLUG]`; the exec op is `pin`.
        """
        resolved = self.anchor_by_id(anchor) if isinstance(anchor, str) else anchor
        target = anchor if isinstance(anchor, str) else resolved.anchor_id
        if name is not None:
            code = _validate_pin_slug(name)
            with _com.translate_com_errors():
                _mint_wl_bookmark(self._doc, resolved.com, code)
        else:
            with _com.translate_com_errors():
                code = _new_pin_code()
                while self._doc.Bookmarks.Exists(_pin_name_for(code)):
                    code = _new_pin_code()
                _mint_wl_bookmark(self._doc, resolved.com, code)
        return {"anchor_id": f"pin:{code}", "pin": f"pin:{code}", "target": target}

    # `stamp` reads better for "stamp a handle on this"; same operation.
    stamp = pin

    def _existing_pin_starts(self) -> dict[int, str]:
        """Map each existing `_wl_` bookmark's range start -> its pin code.

        Backs `pin_outline` idempotency: a heading whose range start already
        carries a wordlive handle reuses it instead of minting a duplicate.
        """
        out: dict[int, str] = {}
        with _com.translate_com_errors():
            for bm in _bookmarks_including_hidden(self._doc):
                nm = str(bm.Name)
                if nm.startswith(_WL_PREFIX):
                    out[int(bm.Range.Start)] = _pin_id_for(nm)
        return out

    def pin_outline(self, *, levels: int | tuple[int, int] | None = None) -> dict[str, str]:
        """Pin every heading at once and return the `{heading_id: pin_id}` map.

        A durable navigation scaffold up front: stamp a handle on each heading so
        an agent can address sections by `pin:` ids that survive the inserts /
        deletes it is about to make, instead of re-reading `outline` after every
        edit. Idempotent — a heading already carrying a wordlive handle reuses it,
        so calling this twice returns the same map (run it once on a stable
        document; the reuse keys on each heading's range start).

        `levels` filters which headings get pinned: `None` (default) pins every
        heading, an `int` n pins levels ``1..n``, and a ``(lo, hi)`` tuple pins
        the inclusive band. Returns an ordered ``{"heading:3": "pin:a3f9c2", …}``.
        Wrap in `doc.edit(...)` for atomic undo. See
        [`pin`][wordlive.Document.pin] for the single-anchor form.
        """
        lo, hi = _resolve_level_band(levels)
        existing = self._existing_pin_starts()
        out: dict[str, str] = {}
        with _com.translate_com_errors():
            for idx, para in enumerate(self._doc.Paragraphs, start=1):
                try:
                    level = int(para.OutlineLevel)
                except Exception:
                    continue
                if level >= 10 or not (lo <= level <= hi):
                    continue
                rng = para.Range
                start = int(rng.Start)
                code = existing.get(start)
                if code is None:
                    code = _new_pin_code()
                    while self._doc.Bookmarks.Exists(_pin_name_for(code)):
                        code = _new_pin_code()
                    _mint_wl_bookmark(self._doc, rng, code)
                    existing[start] = code
                out[f"heading:{idx}"] = f"pin:{code}"
        return out

    @property
    def bibliography_style(self) -> str:
        """The citation/bibliography style (e.g. ``"APA"``, ``"MLA"``, ``"Chicago"``).

        Read/write. Setting it changes how every citation and the bibliography
        render (refresh them with [`update_fields`][wordlive.Document.update_fields]).
        Word accepts a build-dependent set of identifiers; an unsupported value
        raises [`OpError`][wordlive.OpError].
        """
        with _com.translate_com_errors():
            return str(self._doc.Bibliography.BibliographyStyle)

    @bibliography_style.setter
    def bibliography_style(self, style: str) -> None:
        if not str(style).strip():
            raise OpError("bibliography_style must be a non-empty string")
        with _com.translate_com_errors():
            self._doc.Bibliography.BibliographyStyle = str(style)

    @property
    def track_changes(self) -> bool:
        """Whether Word's Track Changes is currently on for this document."""
        with _com.translate_com_errors():
            return bool(self._doc.TrackRevisions)

    @track_changes.setter
    def track_changes(self, value: bool) -> None:
        with _com.translate_com_errors():
            self._doc.TrackRevisions = bool(value)

    @contextmanager
    def tracked_changes(self) -> Iterator[None]:
        """Turn on Track Changes for the duration of the block, then restore it.

        Every mutation made inside the scope is recorded as a tracked revision
        the user can accept or reject — "make this edit *visibly*." The prior
        `TrackRevisions` setting is restored on exit, so the scope stays polite
        even when the user had tracking off.

        Pairs with `edit()` for an atomic, visibly-tracked batch:

            with doc.tracked_changes(), doc.edit("Suggest rewordings"):
                doc.find_replace("utilise", "use", all=True)
        """
        with _com.translate_com_errors():
            previous = bool(self._doc.TrackRevisions)
            self._doc.TrackRevisions = True
        try:
            yield
        finally:
            with _com.translate_com_errors():
                self._doc.TrackRevisions = previous

    # Word's own text-watermark feature names its WordArt shapes with this prefix
    # (e.g. "PowerPlusWaterMarkObject357921"); reusing it means set_watermark
    # replaces a watermark the user added through the ribbon, and remove_watermark
    # finds it — the established convention, not a wordlive marker.
    _WATERMARK_NAME_PREFIX = "PowerPlusWaterMarkObject"

    def set_watermark(
        self,
        text: str,
        *,
        font: str = "Calibri",
        color: str = "#C0C0C0",
        layout: str = "diagonal",
        semitransparent: bool = True,
    ) -> int:
        """Stamp a text watermark (DRAFT / CONFIDENTIAL / …) behind every page.

        Adds a WordArt shape to each section's primary header story — the same
        mechanism (and shape name) as Word's *Design → Watermark → Custom*, so it
        shows behind the body text on every page and replaces any existing text
        watermark. `layout` is ``"diagonal"`` (default, rotated 45°) or
        ``"horizontal"``; `color` is the fill colour (``"#C0C0C0"`` / ``"red"``);
        `semitransparent` washes it out (50% transparency) so body text stays
        readable. Returns the number of sections stamped.

        Any prior watermark is cleared first, so calling it twice doesn't stack.
        Remove one with [`remove_watermark`][wordlive.Document.remove_watermark].
        Wrap in `doc.edit(...)` for atomic undo. Raises `OpError` for a bad
        `layout` or `color`.
        """
        if layout not in ("diagonal", "horizontal"):
            raise OpError(f"watermark layout must be 'diagonal' or 'horizontal'; got {layout!r}")
        try:
            fill_bgr = to_bgr(color)
        except (ValueError, TypeError) as e:
            raise OpError(str(e)) from e
        rotation = 315.0 if layout == "diagonal" else 0.0
        self.remove_watermark()
        with _com.translate_com_errors():
            sections = self._doc.Sections
            count = int(sections.Count)
            for s in range(1, count + 1):
                section = sections(s)
                header = section.Headers(int(WdHeaderFooterIndex.PRIMARY))
                ps = section.PageSetup
                usable = float(ps.PageWidth) - float(ps.LeftMargin) - float(ps.RightMargin)
                width = max(72.0, usable)
                # Shapes live on the HeaderFooter itself, not its Range (a Range
                # has no .Shapes) — this is the header story Word's own watermark
                # feature draws into.
                shape = header.Shapes.AddTextEffect(
                    PresetTextEffect=int(MsoPresetTextEffect.TEXT_EFFECT1),
                    Text=text,
                    FontName=font,
                    FontSize=1.0,  # WordArt scales to the box; explicit size below
                    FontBold=int(MsoTriState.FALSE),
                    FontItalic=int(MsoTriState.FALSE),
                    Left=0.0,
                    Top=0.0,
                )
                shape.Name = f"{self._WATERMARK_NAME_PREFIX}{s}"
                shape.TextEffect.NormalizedHeight = False
                shape.Line.Visible = int(MsoTriState.FALSE)
                shape.Fill.Visible = int(MsoTriState.TRUE)
                shape.Fill.Solid()
                shape.Fill.ForeColor.RGB = fill_bgr
                shape.Fill.Transparency = 0.5 if semitransparent else 0.0
                shape.Rotation = rotation
                shape.LockAspectRatio = int(MsoTriState.TRUE)
                shape.Width = width
                shape.Height = width / 5.0
                shape.WrapFormat.AllowOverlap = True
                shape.WrapFormat.Side = int(WdWrapSideType.BOTH)
                shape.WrapFormat.Type = int(WdWrapType.BEHIND)
                shape.RelativeHorizontalPosition = int(WdRelativeHorizontalPosition.MARGIN)
                shape.RelativeVerticalPosition = int(WdRelativeVerticalPosition.MARGIN)
                shape.Left = float(WdShapePosition.CENTER)
                shape.Top = float(WdShapePosition.CENTER)
        return count

    def remove_watermark(self) -> int:
        """Remove any text watermark added by `set_watermark` (or Word's ribbon).

        Deletes every WordArt shape named like Word's watermark object across all
        sections' header stories. Returns the number of shapes removed (0 if there
        was no watermark). Wrap in `doc.edit(...)` for atomic undo.
        """
        removed = 0
        with _com.translate_com_errors():
            sections = self._doc.Sections
            for s in range(1, int(sections.Count) + 1):
                header = sections(s).Headers(int(WdHeaderFooterIndex.PRIMARY))
                shapes = header.Shapes
                # Delete back-to-front: removing a shape renumbers those after it.
                for i in range(int(shapes.Count), 0, -1):
                    shape = shapes(i)
                    if str(shape.Name or "").startswith(self._WATERMARK_NAME_PREFIX):
                        shape.Delete()
                        removed += 1
        return removed

    def watermark(self) -> WatermarkInfo | None:
        """The text watermark stamped behind the pages, or `None` if there is none.

        The read side of [`set_watermark`][wordlive.Document.set_watermark] /
        [`remove_watermark`][wordlive.Document.remove_watermark] — it walks each
        section's primary header story for the WordArt shape Word's watermark
        feature draws (named like `PowerPlusWaterMarkObject…`) and reads its text.
        Returns a [`WatermarkInfo`][wordlive.WatermarkInfo] (`text` + the 1-based
        `sections` carrying it), or `None` when the document has no text
        watermark. Pure read — selection, scroll, and `Saved` are untouched.

        Only text watermarks (the `set_watermark` / *Design → Watermark* kind) are
        reported; a picture watermark or an ordinary floating shape is not.
        """
        found: dict[int, str] = {}
        with _com.translate_com_errors():
            sections = self._doc.Sections
            for s in range(1, int(sections.Count) + 1):
                header = sections(s).Headers(int(WdHeaderFooterIndex.PRIMARY))
                shapes = header.Shapes
                for i in range(1, int(shapes.Count) + 1):
                    shape = shapes(i)
                    if not str(shape.Name or "").startswith(self._WATERMARK_NAME_PREFIX):
                        continue
                    try:
                        text = str(shape.TextEffect.Text or "")
                    except Exception:
                        # A non-text (picture) watermark shape has no TextEffect text.
                        text = ""
                    found[s] = text
        if not found:
            return None
        # Word stamps the same text into every section; surface the common value
        # (the first non-empty one), falling back to "" if every read came back blank.
        non_empty = [t for t in found.values() if t]
        return WatermarkInfo(text=non_empty[0] if non_empty else "", sections=sorted(found))
