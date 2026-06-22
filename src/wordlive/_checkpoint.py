"""Checkpoint + diff — `doc.checkpoint()` / `doc.changes_since()` / `doc.diff()`.

Fingerprint the document's structure at one moment (`checkpoint`, a pure read),
then produce a structured, content-aligned change list against a later moment
(`diff` / `changes_since`). This is the only reliable way to answer "what
changed in session" — Word emits no content-change event (see
`feature-plan.md` Priority 7), so a fingerprint-and-diff is the mechanism, not a
convenience. It also lets a multi-step agent verify its own edits landed without
re-reading the whole document.

**No new COM surface.** Pure composition over shipped reads:
`_findreplace._normalize` (the shared text normalisation `find`/`find_paragraphs`
use — so the diff agrees and ignores cosmetic churn), `paragraph_text` (paragraph
text with `[image]` tokens), `format_info` (the `text+format` fingerprint), and
`difflib.SequenceMatcher` (the alignment engine `find_paragraphs` already uses).

Design: `spec-checkpoint-diff.md`. Mirrors `_linting.py`'s module shape (free
functions + a `Document` method that delegates here). Build-order steps 1-4 ship
here; `track=True` (pin-backed exact identity), `moves=True`, and per-cell table
diffing are documented follow-ups.
"""

from __future__ import annotations

import difflib
import hashlib
import json
from dataclasses import asdict, dataclass, field
from typing import TYPE_CHECKING, Any

from . import _com, _findreplace
from ._anchors import _read_font, _read_paragraph_format, range_text
from ._tables import _strip_cell_text
from .exceptions import OpError

if TYPE_CHECKING:
    from ._anchors import Anchor
    from ._document import Document

# Checkpoint format version — carried in the token so checkpoints stay
# comparable across wordlive versions (a future format bump can refuse to diff a
# v1 token).
VERSION = 1

# The three fingerprint depths, cheapest first. `text` ignores style (a restyle
# is invisible); `text+style` (default) folds the applied paragraph-style name
# into the change hash (so a restyle surfaces); `text+format` additionally hashes
# the paragraph's `format_info()` so a pure direct-formatting edit surfaces as a
# `reformat`.
INCLUDE_LEVELS = ("text", "text+style", "text+format")

# A (start, end) character span used by `within=` to clip the walk; None = whole
# document (mirrors `_linting.Span`).
Span = tuple[int, int]


@dataclass(frozen=True)
class Checkpoint:
    """An opaque, serialisable structural fingerprint of the document at one
    moment. Build with [`Document.checkpoint`][wordlive.Document.checkpoint]; the
    caller holds the token and feeds it back to `changes_since` / `diff`.

    `paragraphs` is one dict per paragraph in document order
    (`{i, text, style, level, list, fmt, key, hash}`): `key` is the alignment
    identity (normalised text **only** — deliberately not style/level, so a
    restyled paragraph still aligns as `equal` and surfaces as a `restyle` rather
    than a delete+insert), and `hash` is the change key (normalised text plus, per
    `include`, style and the format fingerprint). `tables` fingerprints each table
    coarsely (`{index, shape, cells_hash}` — detects *a* cell changed); `doc_hash`
    is the whole-fingerprint fast-path (equal ⇒ no changes).

    Because `key` is text-only, paragraphs with identical normalised text (blank
    lines, repeated boilerplate) share a key and are interchangeable to the
    aligner, so an edit amid many identical lines can mis-pair them (usually a
    spurious blank-line insert/delete, not a misclassified real change). Exact
    per-paragraph identity is the deferred `track=True` (pin-backed) mode.
    """

    version: int
    include: str
    scope: str | None
    paragraphs: list[dict[str, Any]]
    tables: list[dict[str, Any]] = field(default_factory=list)
    doc_hash: str = ""

    def to_json(self) -> str:
        """Serialise to a JSON string — the token the caller stores."""
        return json.dumps(asdict(self), ensure_ascii=False)

    @classmethod
    def from_json(cls, data: str | dict[str, Any]) -> Checkpoint:
        """Rebuild a `Checkpoint` from `to_json()` output (a JSON string or the
        already-parsed dict)."""
        obj = json.loads(data) if isinstance(data, str) else dict(data)
        return cls(
            version=int(obj.get("version", VERSION)),
            include=str(obj.get("include", "text+style")),
            scope=obj.get("scope"),
            paragraphs=list(obj.get("paragraphs", [])),
            tables=list(obj.get("tables", [])),
            doc_hash=str(obj.get("doc_hash", "")),
        )


def _coerce(cp: Checkpoint | str | dict[str, Any]) -> Checkpoint:
    """Accept a `Checkpoint`, its JSON string, or the parsed dict — so the CLI/MCP
    file token round-trips without the caller reconstructing the dataclass."""
    if isinstance(cp, Checkpoint):
        return cp
    return Checkpoint.from_json(cp)


def _sha1(s: str) -> str:
    """A fixed, deterministic content hash (spec §7) — stable across runs."""
    return hashlib.sha1(s.encode("utf-8")).hexdigest()


def _overlaps(span: Span | None, lo: int, hi: int) -> bool:
    """Does the paragraph range `[lo, hi)` genuinely overlap the scope `span`
    (None = whole doc)? Word paragraph ranges are half-open — the body paragraph
    that *starts* exactly at a heading's end is a different paragraph, not part of
    the heading's scope — so a boundary touch does **not** count (unlike
    `_linting._overlaps`, which scopes heading-line rules where this never bites)."""
    if span is None:
        return True
    return lo < span[1] and hi > span[0]


def _format_fingerprint(rng: Any) -> str:
    """A stable hash of a paragraph range's effective paragraph + character
    formatting, for `include="text+format"`. Reads the *effective* values straight
    off the already-iterated `rng` (same `_read_paragraph_format`/`_read_font`
    helpers `format_info` uses) and folds them into a sorted, deterministic string.

    Takes the live range rather than re-resolving `para:{idx}` via `anchor_by_id`
    — resolving a `Paragraph` anchor re-enumerates `doc.Paragraphs` from the top,
    so probing every paragraph that way made the `text+format` checkpoint O(n²) in
    paragraph count. This stays O(n)."""
    eff_para = _read_paragraph_format(rng.ParagraphFormat)
    eff_font, _mixed = _read_font(rng.Font)
    parts: list[str] = []
    for section, block in (("paragraph", eff_para), ("font", eff_font)):
        for key in sorted(block):
            parts.append(f"{section}.{key}={block[key]}")
    return _sha1("\x1f".join(parts))


def _stable_paragraph_text(rng: Any) -> str:
    """A paragraph's text read with a fixed field-code / hidden-text policy.

    `Range.Text` otherwise reflects the live ``ShowFieldCodes`` / ``ShowHidden``
    view state, so the *same* document could fingerprint differently depending on
    how the user happens to be viewing it (a phantom `replace`), or a `{ DATE }`
    shown as a code vs its result would flip the hash. Pinning the retrieval mode
    makes the read view-independent. (An auto-updating field whose *result*
    changes — e.g. `PAGEREF` — can still surface as a change; excluding field
    results entirely is a deeper follow-up.)"""
    try:
        mode = rng.TextRetrievalMode
        mode.IncludeFieldCodes = False
        mode.IncludeHiddenText = False
    except Exception:
        pass
    return range_text(rng).rstrip("\r\n\x07")


def _change_hash(include: str, norm: str, style: str, fmt: str | None) -> str:
    """The per-paragraph change key — what `doc_hash` and the equal-block restyle
    check compare. Depth scales with `include`: text only · +style · +format."""
    if include == "text":
        payload = norm
    elif include == "text+style":
        payload = f"{norm}\x00{style}"
    else:  # text+format
        payload = f"{norm}\x00{style}\x00{fmt or ''}"
    return _sha1(payload)


def _table_fingerprints(doc_com: Any) -> list[dict[str, Any]]:
    """Coarse per-table fingerprint: `{index, shape, cells_hash}`. v1 detects that
    *a* cell changed (per-cell diffing is deferred). Defensive against irregular
    tables — a cell that raises (merged cells in real Word) is skipped."""
    out: list[dict[str, Any]] = []
    for idx, tbl in enumerate(doc_com.Tables, start=1):
        try:
            rows = int(tbl.Rows.Count)
        except Exception:
            rows = 0
        try:
            cols = int(tbl.Columns.Count)
        except Exception:
            cols = 0
        texts: list[str] = []
        for r in range(1, rows + 1):
            for c in range(1, cols + 1):
                try:
                    raw = _strip_cell_text(range_text(tbl.Cell(r, c).Range))
                except Exception:
                    continue
                texts.append(_findreplace._normalize(raw).text)
        out.append(
            {
                "index": idx,
                "shape": [rows, cols],
                "cells_hash": _sha1("\x07".join(texts)),
            }
        )
    return out


def build_checkpoint(
    doc: Document,
    *,
    include: str = "text+style",
    within: str | Anchor | None = None,
) -> Checkpoint:
    """Walk the paragraphs once and build the fingerprint. Pure read — touches no
    selection/scroll and leaves `Saved` untouched (it only reads)."""
    if include not in INCLUDE_LEVELS:
        raise OpError(f"include must be one of {INCLUDE_LEVELS}, got {include!r}")

    scope_id: str | None = None
    span: Span | None = None
    with _com.translate_com_errors():
        if within is not None:
            anchor = doc.anchor_by_id(within) if isinstance(within, str) else within
            scope_id = anchor.anchor_id
            rng = anchor.com
            span = (int(rng.Start), int(rng.End))

        doc_com = doc._doc
        paragraphs: list[dict[str, Any]] = []
        for idx, para in enumerate(doc_com.Paragraphs, start=1):
            rng = para.Range
            lo, hi = int(rng.Start), int(rng.End)
            if not _overlaps(span, lo, hi):
                continue
            raw = _stable_paragraph_text(rng)
            norm = _findreplace._normalize(raw).text
            try:
                level = int(para.OutlineLevel)
            except Exception:
                level = 10
            try:
                style = str(rng.ParagraphStyle.NameLocal)
            except Exception:
                style = ""
            fmt = _format_fingerprint(rng) if include == "text+format" else None
            paragraphs.append(
                {
                    "i": idx - 1,
                    "text": raw,
                    "style": style,
                    "level": level,
                    "list": None,  # reserved; list-context detection deferred
                    "fmt": fmt,
                    "key": _sha1(norm),  # alignment identity (text only)
                    "hash": _change_hash(include, norm, style, fmt),
                }
            )

        # Tables are fingerprinted only for a whole-document checkpoint; a scoped
        # checkpoint is paragraph-range only (v1).
        tables = _table_fingerprints(doc_com) if span is None else []

    doc_hash = _sha1(
        "\n".join(p["hash"] for p in paragraphs)
        + "\n\x1etables\x1e\n"
        + "\n".join(t["cells_hash"] for t in tables)
    )
    return Checkpoint(
        version=VERSION,
        include=include,
        scope=scope_id,
        paragraphs=paragraphs,
        tables=tables,
        doc_hash=doc_hash,
    )


def _record(op: str, a: dict[str, Any] | None, b: dict[str, Any] | None) -> dict[str, Any]:
    """One change record (spec §4). Insert/replace/restyle carry the *current*
    `para:N` (from `b`) so the caller can act on the change immediately; a delete
    references only the old index/text (its anchor is gone)."""
    rec: dict[str, Any] = {"op": op}
    if a is not None:
        rec["index_before"] = a["i"]
    if b is not None:
        rec["index_after"] = b["i"]
        rec["anchor_id"] = f"para:{b['i'] + 1}"
    if op == "replace":
        rec["text_before"] = a["text"]  # type: ignore[index]
        rec["text_after"] = b["text"]  # type: ignore[index]
    elif op == "insert":
        rec["text_after"] = b["text"]  # type: ignore[index]
    elif op == "delete":
        rec["text_before"] = a["text"]  # type: ignore[index]
    elif op in ("restyle", "reformat"):
        rec["text_before"] = a["text"]  # type: ignore[index]
        rec["text_after"] = b["text"]  # type: ignore[index]
        rec["style_before"] = a["style"]  # type: ignore[index]
        rec["style_after"] = b["style"]  # type: ignore[index]
    return rec


# Within a `replace` opcode of size na×nb, pairing old→new paragraphs by text
# similarity (instead of positionally) makes "an edit next to an insert" classify
# the way a human reads it — the closest pair is the `replace`, the leftover is the
# `insert`/`delete`. The matching is O(na·nb·len); above this cap (a degenerate
# whole-block rewrite) fall back to cheap positional pairing.
_SIM_PAIR_CAP = 2500


def _pair_replace(
    a_items: list[dict[str, Any]], b_items: list[dict[str, Any]]
) -> list[tuple[dict[str, Any] | None, dict[str, Any] | None]]:
    """Pair the paragraphs in one `replace` block. Returns `(a, b)` tuples — both
    present ⇒ replace, only `b` ⇒ insert, only `a` ⇒ delete. Greedy
    highest-similarity matching yields `min(na, nb)` replace pairs; the surplus
    side becomes inserts/deletes."""
    na, nb = len(a_items), len(b_items)
    if na == 0 or nb == 0 or na * nb > _SIM_PAIR_CAP:
        common = min(na, nb)
        pairs: list[tuple[dict[str, Any] | None, dict[str, Any] | None]] = [
            (a_items[k], b_items[k]) for k in range(common)
        ]
        pairs += [(a_items[k], None) for k in range(common, na)]
        pairs += [(None, b_items[k]) for k in range(common, nb)]
        return pairs

    scored: list[tuple[float, int, int]] = []
    for i, a in enumerate(a_items):
        for j, b in enumerate(b_items):
            ratio = difflib.SequenceMatcher(None, a["text"], b["text"]).ratio()
            scored.append((ratio, i, j))
    scored.sort(key=lambda t: t[0], reverse=True)

    used_a: set[int] = set()
    used_b: set[int] = set()
    matched: list[tuple[int, int]] = []
    for _ratio, i, j in scored:
        if i in used_a or j in used_b:
            continue
        used_a.add(i)
        used_b.add(j)
        matched.append((i, j))
    matched.sort()  # stable, ordered by old-document position

    pairs = [(a_items[i], b_items[j]) for i, j in matched]
    pairs += [(a_items[i], None) for i in range(na) if i not in used_a]
    pairs += [(None, b_items[j]) for j in range(nb) if j not in used_b]
    return pairs


def diff_checkpoints(
    cp_a: Checkpoint | str | dict[str, Any],
    cp_b: Checkpoint | str | dict[str, Any],
) -> list[dict[str, Any]]:
    """Content-aligned diff of two checkpoints → a structured change list.

    Aligns the two paragraph sequences by `key` (normalised **text**, not index —
    `para:N` renumbers under inserts/deletes) with `difflib.SequenceMatcher`, then
    classifies each opcode: `replace` (text edit) / `insert` / `delete`, and —
    within an `equal` block whose change-`hash` still differs — `restyle` (style
    changed) or `reformat` (only the format fingerprint changed, `text+format`
    only). Tables are then compared coarsely by index, emitting
    `table_change` / `table_insert` / `table_delete`. A matching `doc_hash`
    short-circuits to `[]` (the cheap "nothing changed" path). Move detection is
    deferred — a cut-paste surfaces as delete+insert.
    """
    a = _coerce(cp_a)
    b = _coerce(cp_b)
    if a.include != b.include:
        raise OpError(
            f"cannot diff checkpoints with different include depths "
            f"({a.include!r} vs {b.include!r})"
        )
    if a.doc_hash == b.doc_hash:
        return []

    pa, pb = a.paragraphs, b.paragraphs
    keys_a = [p["key"] for p in pa]
    keys_b = [p["key"] for p in pb]
    sm = difflib.SequenceMatcher(None, keys_a, keys_b, autojunk=False)

    changes: list[dict[str, Any]] = []
    for tag, i1, i2, j1, j2 in sm.get_opcodes():
        if tag == "equal":
            # Texts align; a differing change-hash means a same-text restyle/reformat.
            for k in range(i2 - i1):
                pa_k, pb_k = pa[i1 + k], pb[j1 + k]
                if pa_k["hash"] == pb_k["hash"]:
                    continue
                # `reformat` is only a meaningful verdict in `text+format` mode and
                # only when the format fingerprint actually moved; otherwise the
                # same-text hash difference is a style change. (In `text`/
                # `text+style` mode no `fmt` is captured, so a hash diff here is
                # always a restyle — never report a `reformat` with no format data.)
                if (
                    pa_k["style"] == pb_k["style"]
                    and a.include == "text+format"
                    and pa_k.get("fmt") != pb_k.get("fmt")
                ):
                    op = "reformat"
                else:
                    op = "restyle"
                changes.append(_record(op, pa_k, pb_k))
        elif tag == "replace":
            # Pair old→new by text similarity so an edit beside an insert/delete
            # classifies intuitively (closest pair = replace, leftover = insert/delete).
            for a_item, b_item in _pair_replace(pa[i1:i2], pb[j1:j2]):
                if a_item is not None and b_item is not None:
                    changes.append(_record("replace", a_item, b_item))
                elif b_item is not None:
                    changes.append(_record("insert", None, b_item))
                else:
                    changes.append(_record("delete", a_item, None))
        elif tag == "delete":
            for k in range(i1, i2):
                changes.append(_record("delete", pa[k], None))
        elif tag == "insert":
            for k in range(j1, j2):
                changes.append(_record("insert", None, pb[k]))

    # Tables are folded into `doc_hash` (so a table-only edit breaks the fast
    # path) but the paragraph alignment above can't see them — without this pass
    # a table-only change would slip through and `diff` would wrongly return [].
    # Compare coarsely by 1-based index (per-cell diffing is deferred): a shape
    # or cell-content change surfaces as `table_change`; an added/removed table
    # as `table_insert`/`table_delete`.
    ta = {t["index"]: t for t in a.tables}
    tb = {t["index"]: t for t in b.tables}
    for index in sorted(set(ta) | set(tb)):
        before, after = ta.get(index), tb.get(index)
        if before == after:
            continue
        if before is None:
            changes.append(
                {
                    "op": "table_insert",
                    "index_after": index - 1,
                    "anchor_id": f"table:{index}",
                    "shape_after": after["shape"],  # type: ignore[index]
                }
            )
        elif after is None:
            changes.append(
                {
                    "op": "table_delete",
                    "index_before": index - 1,
                    "shape_before": before["shape"],
                }
            )
        else:
            changes.append(
                {
                    "op": "table_change",
                    "index_after": index - 1,
                    "anchor_id": f"table:{index}",
                    "shape_before": before["shape"],
                    "shape_after": after["shape"],
                }
            )
    return changes


def changes_since(doc: Document, cp: Checkpoint | str | dict[str, Any]) -> list[dict[str, Any]]:
    """Diff a stored checkpoint against the document *now*. Re-derives the
    checkpoint's `include` depth and `scope` so the two fingerprints are
    comparable."""
    base = _coerce(cp)
    # A `range:START-END` scope is pinned to character offsets; re-resolving it
    # against the *current* document (after edits that shifted offsets — exactly
    # what we're diffing for) would clip a different window than the original
    # checkpoint covered, silently comparing mismatched regions. Stable semantic
    # anchors (`heading:N`, `bookmark:`, `cc:`, `start`/`end`) re-resolve cleanly.
    if base.scope is not None and base.scope.startswith("range:"):
        raise OpError(
            f"changes_since cannot re-derive an offset-based scope ({base.scope!r}) — "
            "offsets shift under edits. Scope the checkpoint with a stable anchor "
            "(heading:N / bookmark: / cc:), or diff two stored checkpoints with diff()."
        )
    now = build_checkpoint(doc, include=base.include, within=base.scope)
    return diff_checkpoints(base, now)
