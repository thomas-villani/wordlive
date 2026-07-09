"""Bookmark and pin minting/validation, cross-reference target resolution, and
caption placement — the naming and addressing layer under the anchor classes."""

from __future__ import annotations

import re
import secrets
from typing import TYPE_CHECKING, Any

from ..constants import (
    WdReferenceKind,
    WdReferenceType,
)
from ..exceptions import AnchorNotFoundError, OpError

if TYPE_CHECKING:
    from .._document import Document

from ._helpers import (
    paragraph_text,
)

# Bookmark names: must start with a letter, then letters/digits/underscores, and
# Word caps them at 40 characters. (Leading-underscore names are Word's hidden
# internal bookmarks, so user-created ones must lead with a letter.)
_BOOKMARK_NAME_RE = re.compile(r"^[A-Za-z][A-Za-z0-9_]*$")


def _validate_bookmark_name(name: Any) -> str:
    """Validate a new bookmark name against Word's rules. Raises `OpError`."""
    if not isinstance(name, str) or not name:
        raise OpError("bookmark name must be a non-empty string")
    if len(name) > 40:
        raise OpError(f"bookmark name must be at most 40 characters; got {len(name)}")
    if not _BOOKMARK_NAME_RE.match(name):
        raise OpError(
            f"invalid bookmark name {name!r}: must start with a letter and contain only "
            "letters, digits, and underscores (no spaces)"
        )
    return name


# Durable handles ("pins") are minted as Word-hidden bookmarks named `_wl_<code>`.
# The leading underscore makes Word treat them as internal (so they stay out of
# `bookmarks.list()`), and Word maintains the bookmark<->range association across
# inserts / deletes / edits natively — that native behaviour is what makes a
# `pin:` anchor durable where a positional `para:N` is not.
_WL_PREFIX = "_wl_"


# A user-supplied pin slug: lowercase alphanumeric words joined by single
# hyphens. No underscores — storage maps `-` -> `_`, so a slug carrying `_`
# would round-trip ambiguously.
_PIN_SLUG_RE = re.compile(r"^[a-z0-9]+(?:-[a-z0-9]+)*$")


def _new_pin_code() -> str:
    """A fresh random pin code — six lowercase hex chars (e.g. ``a3f9c2``).

    Random rather than sequential so codes don't leak document structure and
    two pins never collide by construction. Tests monkeypatch this for
    deterministic ids.
    """
    return secrets.token_hex(3)


def _pin_name_for(code: str) -> str:
    """Bookmark storage name for a pin display code (`pin:<code>` -> `_wl_…`)."""
    return _WL_PREFIX + code.replace("-", "_")


def _pin_id_for(name: str) -> str:
    """Pin display code for a `_wl_` bookmark name (inverse of `_pin_name_for`)."""
    return name[len(_WL_PREFIX) :].replace("_", "-")


def _validate_pin_slug(slug: Any) -> str:
    """Validate a user-supplied pin slug and return it. Raises `OpError`.

    A slug is lowercase alphanumeric words joined by single hyphens
    (``budget-intro``), and its stored bookmark name (`_wl_` + slug with hyphens
    mapped to underscores) must fit Word's 40-character bookmark-name cap.
    """
    if not isinstance(slug, str) or not slug:
        raise OpError("pin name must be a non-empty string")
    if not _PIN_SLUG_RE.match(slug):
        raise OpError(
            f"invalid pin name {slug!r}: use lowercase letters, digits, and single "
            "hyphens (e.g. 'budget-intro')"
        )
    storage = _pin_name_for(slug)
    if len(storage) > 40:
        raise OpError(
            f"pin name {slug!r} is too long: the stored bookmark name {storage!r} "
            "exceeds Word's 40-character limit"
        )
    return slug


def _mint_wl_bookmark(doc_com: Any, rng: Any, code: str) -> str:
    """Plant a hidden `_wl_<code>` bookmark over `rng`; return the bookmark name.

    Bypasses `_validate_bookmark_name` (which forbids the leading underscore) on
    purpose — these are wordlive-internal names, not user-chosen ones. Word still
    enforces its own 40-char cap, which `_validate_pin_slug` checks up front.
    """
    name = _pin_name_for(code)
    doc_com.Bookmarks.Add(Name=name, Range=rng)
    return name


def _stale_anchor_hint(doc_com: Any, kind: str, index: int) -> str | None:
    """A recovery hint for a positional `para:N` / `heading:N` that missed.

    Positional ids renumber under inserts / deletes, so a stale one is the
    common failure. One cheap pass over `Paragraphs` reports whether the index
    is out of range (renumbered / vanished) or — for `heading:N` — points at
    body text, and names the nearest heading. Always recommends pinning a
    durable handle. Returns `None` if the document can't be read (best-effort —
    the hint must never mask the original miss).
    """
    try:
        paras = list(doc_com.Paragraphs)
    except Exception:
        return None
    count = len(paras)
    pin_tip = "Pin a durable handle with `pin` to survive renumbering."

    def _level(p: Any) -> int:
        try:
            return int(p.OutlineLevel)
        except Exception:
            return 10

    def _nearest_heading() -> tuple[int, str] | None:
        best: tuple[int, int, str] | None = None  # (distance, idx, text)
        for i, p in enumerate(paras, start=1):
            if i == index:
                continue
            if _level(p) < 10:
                dist = abs(i - index)
                if best is None or dist < best[0]:
                    best = (dist, i, paragraph_text(p))
        return (best[1], best[2]) if best is not None else None

    nh = _nearest_heading()
    if index < 1 or index > count:
        base = f"{kind}:{index} out of range; document has {count} paragraph(s)"
        base += (
            f' (nearest heading is heading:{nh[0]} "{nh[1]}")'
            if nh
            else " (likely renumbered by an edit)"
        )
        return f"{base}. {pin_tip}"
    if kind == "heading":
        # In range but not a heading (body text at this index).
        base = f"heading:{index} is not a heading (body text)"
        if nh:
            base += f'; nearest heading is heading:{nh[0]} "{nh[1]}"'
        return f"{base}. Use a pin for a stable handle."
    return pin_tip


def _resolve_cross_ref_target(doc: Document, target: str) -> tuple[int, int | str]:
    """Map a cross-reference `target` anchor id to `(ReferenceType, ReferenceItem)`.

    `ReferenceItem` is what Word's `InsertCrossReference` expects, which differs
    by type: for `bookmark:NAME` it is the **bookmark name string** (Word looks
    it up by name, not position); for `heading:N` it is the 1-based ordinal among
    heading paragraphs; for `footnote:N` / `endnote:N` it is the 1-based index.
    Raises `AnchorNotFoundError` (exit 2) for an unknown/unresolvable target.
    """
    kind, _, value = str(target).partition(":")
    if kind == "bookmark":
        try:
            items = [
                str(it).strip()
                for it in doc.com.GetCrossReferenceItems(int(WdReferenceType.BOOKMARK))
            ]
        except Exception as e:
            raise AnchorNotFoundError("bookmark", target) from e
        if value not in items:
            raise AnchorNotFoundError("bookmark", target)
        # For bookmarks the ReferenceItem is the *name*, not a position index.
        return int(WdReferenceType.BOOKMARK), value
    if kind == "heading":
        try:
            n = int(value)
        except ValueError as e:
            raise AnchorNotFoundError("heading", target) from e
        ordinal = 0
        for idx, para in enumerate(doc.com.Paragraphs, start=1):
            try:
                level = int(para.OutlineLevel)
            except Exception:
                level = 10
            if level < 10:
                ordinal += 1
            if idx == n:
                if level >= 10:
                    raise AnchorNotFoundError("heading", target)
                return int(WdReferenceType.HEADING), ordinal
        raise AnchorNotFoundError("heading", target)
    if kind in ("footnote", "endnote"):
        try:
            n = int(value)
        except ValueError as e:
            raise AnchorNotFoundError(kind, target) from e
        ref_type = WdReferenceType.FOOTNOTE if kind == "footnote" else WdReferenceType.ENDNOTE
        coll = doc.footnotes if kind == "footnote" else doc.endnotes
        if not (1 <= n <= len(coll)):
            raise AnchorNotFoundError(kind, target)
        return int(ref_type), n
    raise AnchorNotFoundError(
        "cross-reference target",
        target,
        hint="target must be a bookmark:, heading:, footnote:, or endnote: anchor id",
    )


def _cross_ref_kind(kind: str, ref_type: int) -> int:
    """Map an `insert_cross_reference(kind=...)` string to a `WdReferenceKind`.

    Type-dependent in two places: a note has no "text" content to reference, so
    `"text"`/`"number"` both resolve to its number; for headings/bookmarks
    `"number"` is the paragraph number. Raises `ValueError` on an unknown kind.
    """
    is_note = ref_type in (int(WdReferenceType.FOOTNOTE), int(WdReferenceType.ENDNOTE))
    note_number = (
        int(WdReferenceKind.FOOTNOTE_NUMBER)
        if ref_type == int(WdReferenceType.FOOTNOTE)
        else int(WdReferenceKind.ENDNOTE_NUMBER)
    )
    if kind == "text":
        # wdContentText is invalid for footnotes/endnotes (a mark has no text);
        # fall back to the note number, which is the meaningful reference.
        return note_number if is_note else int(WdReferenceKind.CONTENT_TEXT)
    if kind == "page":
        return int(WdReferenceKind.PAGE_NUMBER)
    if kind == "above_below":
        return int(WdReferenceKind.POSITION)
    if kind == "number":
        return note_number if is_note else int(WdReferenceKind.NUMBER_NO_CONTEXT)
    raise ValueError(
        f"unknown cross-reference kind {kind!r}; expected text/page/number/above_below"
    )


def _caption_above(label: str, position: str | None) -> bool:
    """Resolve a caption's placement to a boolean (`True` = above the anchor).

    `position` is the user override (``"above"``/``"below"``, with
    ``"before"``/``"after"`` accepted as aliases). When it's `None` the
    *convention* applies: a ``"Table"`` caption goes **above**, every other
    label (Figure, Equation, …) goes **below**. Raises `ValueError` on a bad
    string.
    """
    if position is None:
        return str(label).strip().casefold() == "table"
    p = str(position).strip().casefold()
    if p in ("above", "before", "top"):
        return True
    if p in ("below", "after", "bottom"):
        return False
    raise ValueError(f"position must be 'above' or 'below'; got {position!r}")
