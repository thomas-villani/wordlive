"""Reads/queries/exports: outline, digests, stats, lint, checkpoint, snapshot."""

from __future__ import annotations

import difflib
from pathlib import Path
from typing import TYPE_CHECKING, Any

from .. import _checkpoint, _com, _export, _findreplace, _linting, _proofing, _snapshot
from .._anchors import (
    Heading,
    RangeAnchor,
    paragraph_text,
)
from .._checkpoint import Checkpoint
from .._snapshot import Snapshot
from ..constants import (
    WdInformation,
    WdStatistic,
)
from ..exceptions import (
    OpError,
)

if TYPE_CHECKING:
    from .._anchors import Anchor

from ._core import DocumentCore, _markup_flag


class ReadingMixin(DocumentCore):
    """Reads/queries/exports: outline, digests, stats, lint, checkpoint, snapshot."""

    def outline(self, *, pin: bool = False) -> list[dict[str, Any]]:
        """Return all heading paragraphs as `[{level, text, anchor_id}, ...]`.

        With `pin=True` each row also carries a durable `pin` id and the headings
        are pinned as a side effect (idempotent — see
        [`pin_outline`][wordlive.Document.pin_outline]). This **mutates** the
        document, so it is a Python-API-only convenience; the read surfaces
        (`wordlive read outline`, MCP `word_read outline`) stay pure — pin in bulk
        via `pin_outline` / the `pin_outline` exec op instead.
        """
        out: list[dict[str, Any]] = []
        with _com.translate_com_errors():
            for idx, para in enumerate(self._doc.Paragraphs, start=1):
                try:
                    level = int(para.OutlineLevel)
                except Exception:
                    continue
                if level >= 10:
                    continue
                out.append(
                    {
                        "level": level,
                        "text": paragraph_text(para),
                        "anchor_id": f"heading:{idx}",
                    }
                )
        if pin:
            pinmap = self._as_document.pin_outline()
            for row in out:
                handle = pinmap.get(row["anchor_id"])
                if handle:
                    row["pin"] = handle
        return out

    def between(
        self,
        start: str | Anchor,
        end: str | Anchor,
        *,
        inclusive: bool = False,
    ) -> RangeAnchor:
        """Return a `RangeAnchor` spanning the gap between two anchors.

        The "give me the block between these two headings" read. `start` and
        `end` are anchor ids (e.g. ``"heading:1"`` / ``"heading:3"``) or
        `Anchor` objects; the headline use is a pair of `heading:N` ids, but any
        anchors work (bookmarks, paragraphs, ranges).

        With ``inclusive=False`` (default) the span runs from the **end** of
        `start`'s range to the **start** of `end`'s range — the content strictly
        between them, excluding both bounding paragraphs (so two headings yield
        just the body in between). With ``inclusive=True`` it runs from the
        start of `start` to the end of `end`, covering both bounding paragraphs.

        Read ``.text`` on the result for the spanned text, or feed its
        `range:START-END` id into any range-taking op. A pure read (the returned
        offsets are live — use them before further edits shift the document).
        Raises `OpError` if `end` begins before `start`.
        """
        with _com.translate_com_errors():
            s_anchor = self.anchor_by_id(start) if isinstance(start, str) else start
            e_anchor = self.anchor_by_id(end) if isinstance(end, str) else end
            s_rng, e_rng = s_anchor.com, e_anchor.com
            s_start, s_end = int(s_rng.Start), int(s_rng.End)
            e_start, e_end = int(e_rng.Start), int(e_rng.End)
        if e_start < s_start:
            raise OpError(
                f"'between' end anchor ({e_anchor.anchor_id}) begins before start "
                f"anchor ({s_anchor.anchor_id})"
            )
        if inclusive:
            lo, hi = min(s_start, e_start), max(s_end, e_end)
        else:
            # Strictly between: end of `start` to start of `end`. When the anchors
            # abut with no gap, clamp to an empty span at the boundary.
            lo, hi = s_end, max(s_end, e_start)
        return self.range(lo, hi)

    def nearest_heading(
        self,
        where: str | Anchor | int,
        *,
        direction: str = "before",
    ) -> dict[str, Any] | None:
        """The heading nearest to a position, scanning ``before`` or ``after`` it.

        `where` is an anchor id (``"para:12"``), an `Anchor`, or a raw character
        offset (int). `direction` is ``"before"`` (the nearest heading at or
        above the position — i.e. the section the position sits in) or
        ``"after"`` (the next heading past it). Returns an `outline()`-shaped
        row ``{level, text, anchor_id}`` (``anchor_id`` is ``heading:N``), or
        ``None`` if there is no heading in that direction. A pure read.
        """
        if direction not in ("before", "after"):
            raise OpError(f"direction must be 'before' or 'after', got {direction!r}")
        with _com.translate_com_errors():
            if isinstance(where, str):
                offset = int(self.anchor_by_id(where).com.Start)
            elif isinstance(where, int):  # raw character offset
                offset = int(where)
            else:
                offset = int(where.com.Start)
            best: dict[str, Any] | None = None
            for idx, para in enumerate(self._doc.Paragraphs, start=1):
                try:
                    level = int(para.OutlineLevel)
                except Exception:
                    continue
                if level >= 10:  # body text, not a heading
                    continue
                h_start = int(para.Range.Start)
                row = {"level": level, "text": paragraph_text(para), "anchor_id": f"heading:{idx}"}
                if direction == "before":
                    if h_start <= offset:
                        best = row  # paragraphs are in order; keep the last one at/above
                    else:
                        break
                elif h_start > offset:  # "after": first heading strictly past the offset
                    best = row
                    break
        return best

    def find_paragraphs(
        self,
        text: str,
        *,
        limit: int = 5,
        min_score: float = 0.6,
    ) -> list[dict[str, Any]]:
        """Fuzzy-rank paragraphs by similarity to `text` (typo/paraphrase tolerant).

        Unlike `find()` (exact substring on normalized text), this scores
        **every paragraph** against `text` with `difflib.SequenceMatcher` over
        the same normalized form (NFKC, smart quotes, dashes, whitespace) — so
        an approximately-remembered paragraph still locates its `para:N`.
        Returns up to `limit` rows, sorted by descending `score`, keeping only
        those with ``score >= min_score``:
        ``[{anchor_id, index, score, text, level, is_heading}, ...]``. An empty
        or whitespace-only query returns ``[]``. A pure read.
        """
        if limit < 1:
            raise OpError(f"limit must be >= 1, got {limit}")
        if not 0.0 <= min_score <= 1.0:
            raise OpError(f"min_score must be in [0, 1], got {min_score}")
        needle = _findreplace._normalize(text).text
        if not needle:
            return []
        scored: list[dict[str, Any]] = []
        with _com.translate_com_errors():
            for idx, para in enumerate(self._doc.Paragraphs, start=1):
                raw = paragraph_text(para)
                hay = _findreplace._normalize(raw).text
                if not hay:
                    continue
                score = difflib.SequenceMatcher(None, needle, hay).ratio()
                if score < min_score:
                    continue
                try:
                    level = int(para.OutlineLevel)
                except Exception:
                    level = 10
                scored.append(
                    {
                        "anchor_id": f"para:{idx}",
                        "index": idx,
                        "score": round(score, 4),
                        "text": raw,
                        "level": level,
                        "is_heading": level < 10,
                    }
                )
        scored.sort(key=lambda r: r["score"], reverse=True)
        return scored[:limit]

    def stats(self) -> dict[str, Any]:
        """A one-call summary of the document — the "what am I looking at" read.

        Returns `{pages, words, characters, paragraphs, lines, sections,
        headings, tables, images, equations, charts, comments, revisions, saved}`. The five text
        counts come from Word's own `ComputeStatistics`; the structural counts
        come from wordlive's discovery collections (so they agree with
        `doc.tables` / `doc.images` / `outline` etc.); `saved` is `doc.saved`.

        `pages`/`lines` are print-layout truth, so the document is
        **repaginated first** (content-neutral — selection, scroll, and view are
        untouched), the same guarantee a `snapshot` gives. A pure read; nothing
        is mutated — `Repaginate` flips Word's dirty bit, so the document's
        `Saved` state is snapshotted and restored around it.
        """
        with _com.translate_com_errors(), _com.preserve_saved(self._doc):
            self._doc.Repaginate()
            text_counts = {
                "pages": int(self._doc.ComputeStatistics(int(WdStatistic.PAGES))),
                "words": int(self._doc.ComputeStatistics(int(WdStatistic.WORDS))),
                "characters": int(self._doc.ComputeStatistics(int(WdStatistic.CHARACTERS))),
                "paragraphs": int(self._doc.ComputeStatistics(int(WdStatistic.PARAGRAPHS))),
                "lines": int(self._doc.ComputeStatistics(int(WdStatistic.LINES))),
            }
        return {
            **text_counts,
            "sections": len(self.sections),
            "headings": len(self.outline()),
            "tables": len(self.tables),
            "images": len(self.images),
            "equations": len(self.equations),
            "charts": len(self.charts),
            "comments": len(self.comments),
            "revisions": len(self.revisions),
            "saved": self.saved,
        }

    def to_markdown(self, *, within: str | Anchor | None = None) -> str:
        """Serialise the document (or one anchor's range) to clean Markdown.

        The read mirror of [`insert_markdown`][wordlive.Anchor.insert_markdown]:
        headings (``#``–``######``), bullet / numbered lists (with nesting),
        ``**bold**`` / ``*italic*`` / ``***both***``, GFM pipe tables, inline
        images as ``![alt](image:N)``, and hyperlinks as ``[text](url)``. The
        constrained subset import speaks round-trips; the rest is a richer read
        (export is **lossy by design** — underline, colours, and merged table
        cells do not survive).

        `within` scopes the output to an anchor's **literal range** — pass a
        `range:START-END` (e.g. from `find`) or any anchor id / `Anchor`. A
        `heading:N` covers only the heading line, not its section body — use
        `doc.between(...)` or a range for "the section under X". ``None`` (the
        default) serialises the whole document. A pure read; nothing is mutated.
        """
        with _com.translate_com_errors():
            blocks = _export.walk_blocks(self._as_document, within)
        return _export.render_markdown(blocks)

    def to_html(self, *, within: str | Anchor | None = None) -> str:
        """Serialise the document (or one anchor's range) to an HTML fragment.

        The HTML counterpart of [`to_markdown`][wordlive.Document.to_markdown],
        rendered from the same document walk so the two agree on structure:
        headings (``<h1>``–``<h6>``), ``<ul>``/``<ol>`` lists (nested),
        ``<strong>``/``<em>``/``<u>``, ``<table>``, ``<img>``, and ``<a>``. Unlike
        the Markdown dialect, HTML keeps underline. Returns a fragment (no
        ``<html>``/``<body>`` wrapper). `within` scopes to an anchor's literal
        range (see `to_markdown`); ``None`` is the whole document. A pure read.
        """
        with _com.translate_com_errors():
            blocks = _export.walk_blocks(self._as_document, within)
        return _export.render_html(blocks)

    def read(self, *, budget: int = 6000, depth: int | None = None) -> str:
        """A token-budgeted, structure-aware digest of the **whole** document.

        Loads a large document into context cheaply while keeping **every anchor
        addressable**: headings are emitted verbatim (each tagged with its
        `<!-- heading:N -->` anchor — the navigation spine), tables become one-line
        shape stubs (`> table:N — R rows × C cols: …`), and body text is sampled to
        fit `budget` (an approximate token count, ~4 chars/token), weighted so
        shallower sections keep more than deep ones. Overflow is elided to markers
        that still name the `para:` range and word count, so an agent can drill in
        with [`to_markdown(within=…)`][wordlive.Document.to_markdown]. `depth` caps
        how deep a section keeps any body (deeper sections collapse to a marker).

        Returns annotated Markdown. A pure read; the eliding heuristic's knobs live
        in `_export` for tuning. For the full text of any region, use `to_markdown`.
        """
        with _com.translate_com_errors():
            blocks = _export.walk_blocks(self._as_document, None)
        return _export.build_digest(blocks, budget=budget, depth=depth)

    def proofing(self) -> dict[str, Any]:
        """Run Word's proofing tools and report spelling, grammar, and readability.

        Returns `{spelling, grammar, readability}`. `spelling` and `grammar` are
        each `{count, errors}` — the exact error count plus a (capped) list of
        `{text, anchor_id, para}` for the flagged runs, so a `range:START-END`
        can be fed back into `read` or `comments.add`. `readability` is Word's
        readability statistics (Flesch Reading Ease, Flesch-Kincaid Grade Level,
        passive-sentence %, word/sentence averages), snake_cased.

        Heavier than [`stats`][wordlive.Document.stats]: it asks Word to (re)check
        the document. Still a pure read — nothing is mutated. If proofing is
        disabled or the document is protected, the affected section reports a
        `None` count / empty readability rather than failing.
        """
        return _proofing.read_proofing(self._as_document)

    def lint(
        self,
        *,
        rules: Any = None,
        within: str | Anchor | None = None,
        profile: Any = None,
    ) -> list[dict[str, Any]]:
        """Audit the document for publishing-quality defects — a pure read.

        Returns a severity-ranked list of findings, each
        `{rule, kind, severity, anchor_id, message, fixable, fix, observed,
        expected}`. `kind` is `consistency` (a direct override fighting the
        applied style — a `Heading 1` at 15pt), `structural` (an objective layout
        defect — a heading that may dangle at a page foot, a multi-page table with
        no repeating header, a numbered list Word split into independent "1."
        runs), or `policy` (a house-style target — off unless a `profile` enables
        it). A `fixable` finding carries an op-shaped `fix` describing exactly what
        [`regularize`][wordlive.Document.regularize] would change.

        `rules` selects which rules run: `None` is the default set (all
        consistency + structural); a list of rule ids / tags
        (`["headings", "lists"]`) includes only those; `{"exclude": [...]}` runs
        the default set minus the listed ids/tags. `within=anchor` scopes the
        audit to an anchor's range (a heading's section, a `range:`, a table).

        `profile` is a house-style config (a path to a `wordlive.lint.json` file,
        an inline dict, or `None`) that opts **policy** rules in, supplies their
        targets (`body-line-spacing`'s spacing, `table-numeric-right-align`'s
        threshold), and can override a rule's severity or disable a default rule —
        `spec-linter.md` §6.

        Read-only — selection, scroll, and `Saved` are untouched (the layout
        rules repaginate content-neutrally, like [`stats`][wordlive.Document.stats]).
        """
        return [
            f.to_dict()
            for f in _linting.run_lint(
                self._as_document, rules=rules, within=within, profile=profile
            )
        ]

    def regularize(
        self,
        *,
        rules: Any = None,
        within: str | Anchor | None = None,
        profile: Any = None,
        dry_run: bool = False,
        allow_content: bool = False,
    ) -> dict[str, Any]:
        """Apply the fixable [`lint`][wordlive.Document.lint] findings in one
        atomic-undo step. Returns `{applied, skipped, deferred, findings}`.

        Each fixable finding's `fix` op(s) run through the batch op loop inside a
        single `doc.edit("Regularize formatting")`, so one Ctrl-Z reverts the
        whole pass and the user's selection/scroll are preserved. The default
        fixes are **targeted and idempotent** — they bring a drifted direct
        override back to its style's value, so running `regularize` twice applies
        nothing the second time. `rules` / `within` / `profile` are as for `lint`
        (a `profile` also lets its policy fixes — justify, line-spacing,
        numeric-column alignment — participate). `dry_run=True` plans the fixes
        (returning them in `findings`) without writing.

        **Formatting/structure fixes apply by default; content-changing fixes are
        opt-in.** A fix that adds or destroys content (inserting a caption/notice,
        deleting a stray paragraph, stripping a watermark) is flagged
        `adds_content` and **withheld** unless `allow_content=True`. Withheld fixes
        are listed in `deferred` so you can see what an opt-in would apply. If
        Track Changes is on, the edits are tracked like any other for review.
        """
        return _linting.regularize(
            self._as_document,
            rules=rules,
            within=within,
            profile=profile,
            dry_run=dry_run,
            allow_content=allow_content,
        )

    def checkpoint(
        self,
        *,
        include: str = "text+style",
        within: str | Anchor | None = None,
    ) -> Checkpoint:
        """Fingerprint the document's structure right now — a pure read.

        Returns an opaque, serialisable [`Checkpoint`][wordlive.Checkpoint] (call
        `.to_json()` to store it). Later, feed it to
        [`changes_since`][wordlive.Document.changes_since] (checkpoint → now) or
        [`diff`][wordlive.Document.diff] (two stored checkpoints) for a structured,
        content-aligned change list — the only reliable way to answer "what
        changed in session" (Word emits no content-change event), and the way an
        agent verifies its own edits landed without re-reading the whole document.

        `include` sets the fingerprint depth: ``"text"`` (cheapest — a restyle is
        invisible), ``"text+style"`` (default — folds the applied paragraph-style
        name in, so a restyle surfaces), or ``"text+format"`` (also hashes each
        paragraph's `format_info`, so a pure direct-formatting edit surfaces as a
        `reformat`). `within=anchor` fingerprints just one section/range.

        Read-only — walks paragraphs like [`outline`][wordlive.Document.outline],
        touching no selection/scroll and leaving `Saved` untouched.
        """
        return _checkpoint.build_checkpoint(self._as_document, include=include, within=within)

    def changes_since(self, cp: Checkpoint | str | dict[str, Any]) -> list[dict[str, Any]]:
        """Diff a stored checkpoint against the document **now** — a pure read.

        `cp` is a [`Checkpoint`][wordlive.Checkpoint] (or its `to_json()` string /
        parsed dict, so a token round-tripped through a file works directly).
        Returns the change list described in [`diff`][wordlive.Document.diff]; the
        checkpoint's `include` depth and `within` scope are re-derived so the two
        fingerprints are comparable. An unchanged document returns ``[]`` via the
        `doc_hash` fast-path.
        """
        return _checkpoint.changes_since(self._as_document, cp)

    def diff(
        self,
        cp_a: Checkpoint | str | dict[str, Any],
        cp_b: Checkpoint | str | dict[str, Any],
    ) -> list[dict[str, Any]]:
        """Diff two stored checkpoints → a structured, content-aligned change list.

        Each change is one of: ``replace`` (text edit), ``insert``, ``delete``,
        ``restyle`` (same text, paragraph style changed), or ``reformat`` (same
        text+style, direct formatting changed — only with ``include="text+format"``).
        Records carry ``{op, anchor_id, index_before, index_after, text_before,
        text_after, style_before, style_after}`` as applicable; inserts/replaces/
        restyles carry the **current** ``para:N`` (`anchor_id`) so the caller can
        act on the change immediately, while a delete references only the old
        index/text (its anchor is gone).

        Alignment is by paragraph **content**, not index (`para:N` renumbers under
        inserts/deletes). Both checkpoints must share the same `include` depth.
        Move detection is deferred — a cut-paste surfaces as delete+insert. A pure
        read (the tokens carry the data; Word is not touched).
        """
        return _checkpoint.diff_checkpoints(cp_a, cp_b)

    def _page_of(self, position: int) -> int:
        """1-based page number that document offset `position` falls on."""
        with _com.translate_com_errors():
            rng = self._doc.Range(int(position), int(position))
            return int(rng.Information(int(WdInformation.ACTIVE_END_PAGE_NUMBER)))

    @staticmethod
    def _resolve_page_arg(pages: int | tuple[int, int] | None) -> tuple[int | None, int | None]:
        """Normalise a `pages` argument into a `(from, to)` 1-based span (or all)."""
        if pages is None:
            return None, None
        if isinstance(pages, bool):  # bool is an int subclass — reject before the int branch
            raise ValueError(f"pages must be an int or (start, end) tuple, not {pages!r}")
        if isinstance(pages, int):
            if pages < 1:
                raise ValueError(f"page number must be >= 1, got {pages}")
            return pages, pages
        if isinstance(pages, (tuple, list)) and len(pages) == 2:
            start, end = int(pages[0]), int(pages[1])
            if start < 1 or end < start:
                raise ValueError(f"invalid page span: {pages!r}")
            return start, end
        raise ValueError(f"pages must be an int or (start, end) tuple, got {pages!r}")

    def _anchor_page_span(self, anchor: Anchor) -> tuple[int, int]:
        """Page span an anchor occupies. Headings expand to their whole section.

        Mirrors `_scope_range`'s heading-means-its-body rule, so a snapshot of a
        `heading:` anchor shows the section a model is editing, not just the
        heading line.
        """
        with _com.translate_com_errors():
            if isinstance(anchor, Heading):
                head = anchor.com
                start, end = int(head.Start), int(anchor.section_range().End)
            else:
                rng = anchor.com
                start, end = int(rng.Start), int(rng.End)
        from_page = self._page_of(start)
        to_page = max(from_page, self._page_of(max(start, end)))
        return from_page, to_page

    def snapshot(
        self,
        out: str | Path | None = None,
        *,
        pages: int | tuple[int, int] | None = None,
        dpi: int = 150,
        max_dim: int | None = None,
        markup: str = "none",
    ) -> list[Snapshot]:
        """Render document page(s) to PNG so a vision model can *see* the layout.

        Word exports a pixel-faithful PDF of the live document and wordlive
        rasterises the requested pages — a true WYSIWYG image (real fonts,
        spacing, page geometry), ideal for iterating on style and formatting.

        `pages` selects what to render: `None` (default) renders every page,
        an `int` a single 1-based page, and a `(start, end)` tuple an inclusive
        span. Returns one [`Snapshot`][wordlive.Snapshot] per page (so a single
        page is a one-element list); read `.png` for the bytes.

        If `out` is given the image is also written there: a single page to `out`
        itself, multiple pages alongside it as `<stem>-p<N><suffix>`.

        `markup` is `"none"` (default — render the final document) or `"all"`
        (render tracked changes and comments as visible revision marks and
        balloons). The marks come from the export, not a view change, so the
        user's on-screen markup setting is left untouched. The structured
        counterpart is [`revisions`][wordlive.Document.revisions].

        `dpi` controls resolution; ~150 reads well for a vision model without
        bloating the image. `max_dim` caps each page's **long edge** in pixels,
        only ever lowering the resolution — the lever for a cheap *whole-document*
        layout check (a vision model is billed on pixel area, so a long-edge cap
        gives a predictable per-page token budget regardless of paper size; ~1000
        keeps a page legible for "did my styling land" at a fraction of the
        tokens). `dpi=72` is a coarser alternative. Read-only — the document and
        the user's cursor are untouched. Requires the `snapshot` extra (PyMuPDF),
        else [`SnapshotError`][wordlive.SnapshotError].
        """
        if max_dim is not None and (isinstance(max_dim, bool) or int(max_dim) < 1):
            raise OpError(f"max_dim must be a positive integer (pixels); got {max_dim!r}")
        from_page, to_page = self._resolve_page_arg(pages)
        rendered = _snapshot.render(
            self._doc,
            from_page=from_page,
            to_page=to_page,
            dpi=dpi,
            max_dim=max_dim,
            markup=_markup_flag(markup),
        )
        return _snapshot.build_snapshots(rendered, out)

    def snapshot_anchor(
        self,
        anchor: Anchor,
        out: str | Path | None = None,
        *,
        dpi: int = 150,
        max_dim: int | None = None,
        markup: str = "none",
    ) -> list[Snapshot]:
        """Render the page(s) an anchor sits on. Backs [`Anchor.snapshot`][wordlive.Anchor.snapshot].

        A `heading:` anchor expands to its whole section (the heading plus the
        body beneath it, up to the next same-or-higher heading); any other
        anchor renders the page(s) its range spans. See
        [`snapshot`][wordlive.Document.snapshot] for `out`/`dpi`/`max_dim`/`markup`
        semantics and the return shape.
        """
        if max_dim is not None and (isinstance(max_dim, bool) or int(max_dim) < 1):
            raise OpError(f"max_dim must be a positive integer (pixels); got {max_dim!r}")
        from_page, to_page = self._anchor_page_span(anchor)
        rendered = _snapshot.render(
            self._doc,
            from_page=from_page,
            to_page=to_page,
            dpi=dpi,
            max_dim=max_dim,
            markup=_markup_flag(markup),
        )
        return _snapshot.build_snapshots(rendered, out)
