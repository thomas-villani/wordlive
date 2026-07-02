"""Hyperlink rules — the §I "are the links sound and print-ready?" cluster.

`spec-linter.md` §5b·I: three rules that walk `doc.hyperlinks` (the read wrapper
already returns `{text, address, sub_address, anchor_id, para}` per link, so this
batch adds **no new read surface** — it is pure rule engine, the Batch-3 shape):

- `hyperlink-broken-internal` — an in-document jump (`SubAddress` set, `Address`
  empty — a ``HYPERLINK \\l bookmark``) whose target bookmark no longer exists. An
  unambiguous dead link, so **on by default** like the v1 structural set.
- `hyperlink-bare-for-print` — an external link whose visible text doesn't contain
  its URL, so the destination is invisible on paper. A **print/sharing** policy,
  **off by default** (tags `hyperlinks` / `print`).
- `hyperlink-display-is-raw-url` — a link whose visible text is itself a bare URL
  where a human-readable label is wanted. **Off by default** (tags `hyperlinks` /
  `print`; `--rule print` selects just these two print/sharing rules).

All three are **report-only** (`fixable=False`): repairing a broken jump needs a
human to pick the intended bookmark, and the print/label fixes add content
(append `(url)`, invent a label) — opt-in, deferred with the `adds_content` gate.
The write verbs those fixes would use already exist (`Hyperlink.update`), so this
batch adds no new COM write surface.

Detection reuses the shipped `doc.hyperlinks` / `doc.bookmarks` read wrappers.
`name in doc.bookmarks` resolves through Word's `Bookmarks.Exists`, which also sees
the hidden `_Toc…` / `_Ref…` bookmarks a real internal link usually targets — so a
live cross-reference jump is correctly *not* flagged as broken. Imported by
`_linting` for its side effect of registering the rules.
"""

from __future__ import annotations

import re
from collections.abc import Iterator
from typing import TYPE_CHECKING, Any

from ._linting import Finding, Rule, Span, _overlaps, _register_rule
from .exceptions import ComError

if TYPE_CHECKING:
    from ._document import Document
    from ._lint_profile import Profile

# The visible text of a link that is *itself* a bare URL token (no surrounding
# prose) — the signal that a raw URL is on show where a label was wanted. Anchored
# so "see https://x for more" (a URL embedded in prose) doesn't match; that is a
# different concern than a link whose whole label is the URL.
_BARE_URL = re.compile(r"^(?:https?://|ftp://|www\.)\S+$", re.IGNORECASE)


def _range_of(anchor_id: Any) -> tuple[int, int] | None:
    """Parse a `range:START-END` anchor id back to `(start, end)`, else `None`."""
    if isinstance(anchor_id, str) and anchor_id.startswith("range:"):
        try:
            lo, hi = anchor_id[len("range:") :].split("-", 1)
            return int(lo), int(hi)
        except ValueError:
            return None
    return None


def _hyperlinks(doc: Document) -> list[dict[str, Any]]:
    """`doc.hyperlinks.list()`, or `[]` on a transient COM hiccup (skip, don't fail)."""
    try:
        return doc.hyperlinks.list()
    except ComError:
        return []


def _check_hyperlink_broken_internal(
    doc: Document, span: Span | None, profile: Profile
) -> Iterator[Finding]:
    """An internal jump (SubAddress set, Address empty) whose target bookmark is
    gone — a dead in-document link. Report-only: the repair needs a human to choose
    the intended target (the same stance as `broken-cross-reference`)."""
    for link in _hyperlinks(doc):
        sub = str(link.get("sub_address") or "").strip()
        address = str(link.get("address") or "").strip()
        if not sub or address:
            continue  # external link (or a fragment on a remote URL) — not an in-doc jump
        rng = _range_of(link.get("anchor_id"))
        if rng is not None and not _overlaps(span, *rng):
            continue
        try:
            if sub in doc.bookmarks:
                continue  # target exists (Exists sees hidden _Toc/_Ref bookmarks too)
        except ComError:
            continue
        yield Finding(
            rule="hyperlink-broken-internal",
            kind="structural",
            severity="warning",
            anchor_id=link.get("anchor_id") or "start",
            message=(
                f"Internal hyperlink {str(link.get('text') or '').strip()!r} points at "
                f"bookmark {sub!r}, which doesn't exist; the link is dead."
            ),
            fixable=False,
            observed=f"sub_address={sub!r} (no such bookmark)",
            expected="a bookmark that exists",
        )


def _check_hyperlink_bare_for_print(
    doc: Document, span: Span | None, profile: Profile
) -> Iterator[Finding]:
    """An external link whose visible text doesn't contain its URL — the destination
    is invisible when the document is printed. Off by default (a print/sharing
    concern). Report-only: appending `(url)` adds content (an opt-in fix, deferred)."""
    for link in _hyperlinks(doc):
        address = str(link.get("address") or "").strip()
        text = str(link.get("text") or "").strip()
        if not address:
            continue  # internal jump — no URL to lose on paper
        if address.casefold() in text.casefold():
            continue  # the URL is already visible in the label
        rng = _range_of(link.get("anchor_id"))
        if rng is not None and not _overlaps(span, *rng):
            continue
        yield Finding(
            rule="hyperlink-bare-for-print",
            kind="policy",
            severity="info",
            anchor_id=link.get("anchor_id") or "start",
            message=(
                f"Hyperlink {text!r} hides its destination ({address!r}); the URL is "
                "invisible in print."
            ),
            fixable=False,
            observed=f"text={text!r}, address={address!r}",
            expected="the URL shown alongside the label",
        )


def _check_hyperlink_display_is_raw_url(
    doc: Document, span: Span | None, profile: Profile
) -> Iterator[Finding]:
    """A link whose whole visible text is a bare URL where a readable label was
    wanted. Off by default (stylistic). Report-only: choosing a label is a human
    call. Distinct from `hyperlink-bare-for-print` (which fires when the URL is
    *hidden*); this fires when the URL is the label."""
    for link in _hyperlinks(doc):
        text = str(link.get("text") or "").strip()
        if not _BARE_URL.match(text):
            continue
        rng = _range_of(link.get("anchor_id"))
        if rng is not None and not _overlaps(span, *rng):
            continue
        yield Finding(
            rule="hyperlink-display-is-raw-url",
            kind="consistency",
            severity="info",
            anchor_id=link.get("anchor_id") or "start",
            message=(
                f"Hyperlink shows a raw URL ({text!r}) as its text; consider a descriptive label."
            ),
            fixable=False,
            observed=f"text={text!r}",
            expected="a descriptive link label",
        )


for _rule in (
    Rule(
        id="hyperlink-broken-internal",
        kind="structural",
        severity="warning",
        tags=("hyperlinks",),
        check=_check_hyperlink_broken_internal,
        default_on=True,
    ),
    Rule(
        id="hyperlink-bare-for-print",
        kind="policy",
        severity="info",
        tags=("hyperlinks", "print"),
        check=_check_hyperlink_bare_for_print,
        default_on=False,
    ),
    Rule(
        id="hyperlink-display-is-raw-url",
        kind="consistency",
        severity="info",
        tags=("hyperlinks", "print"),
        check=_check_hyperlink_display_is_raw_url,
        default_on=False,
    ),
):
    _register_rule(_rule)
