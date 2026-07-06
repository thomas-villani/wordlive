"""Page-layout & document-level rules — the §H "is the document furniture right?"
cluster.

`spec-linter.md` §5b·H (the **P4** primitive: a section / header-footer walk plus
a couple of document-level probes). Five rules ship here, all **off by default**
(a §H issue is rarely a defect mid-authoring — like the finalization cluster,
these are an opt-in "getting it ready to hand off" check):

- `document-properties-filled` — a required built-in property (default `Title` /
  `Author`, overridable via the profile's `required` list) is empty or unset.
- `confidentiality-notice` — a profile-supplied notice string is not present in any
  header/footer or the body (the notice is missing where policy expects it).
- `copyright-notice` — like the above for a copyright line (profile `text`, default
  `"©"`).
- `header-footer-consistent` — the primary header (or footer) **text** disagrees
  across the document's own (non-linked) sections where uniformity is expected.
- `draft-watermark-present` — a text watermark (a leftover *DRAFT* / *CONFIDENTIAL*
  stamp) is still on the document.

All five are **report-only** (`fixable=False`): filling a property, inserting a
notice, or stripping a watermark each add or destroy content — opt-in fixes,
deferred with the `adds_content` gate (the write verbs already exist —
`doc.properties.set`, `doc.remove_watermark` — so this batch adds no new COM write
surface). Detection reuses the shipped `doc.properties` / `doc.sections` read
wrappers plus the new `doc.watermark()` read (the mirror of
`set_watermark`/`remove_watermark` this batch added).

The two notice rules and `document-properties-filled` are `policy` (off until a
profile opts them in and — for the notices — supplies the text); the other two are
off via `default_on=False`. Every rule is document-global, so `within` does not
scope it. Imported by `_linting` for its side effect of registering the rules.
"""

from __future__ import annotations

from collections.abc import Iterator
from typing import TYPE_CHECKING, Any

from ._linting import Finding, Rule, Span, _register_rule
from .exceptions import ComError

if TYPE_CHECKING:
    from ._document import Document
    from ._lint_profile import Profile

# The header/footer stories to scan per section (primary, first-page, even-pages).
_HF_WHICH = ("primary", "first", "even")

# The built-in properties `document-properties-filled` requires by default — the two
# a hand-off almost always wants set. A profile's `required` list overrides this.
_DEFAULT_REQUIRED_PROPS = ("Title", "Author")


def _builtin_properties(doc: Document) -> dict[str, Any]:
    """`doc.properties.builtin()`, or `{}` on a transient COM hiccup (skip, don't fail)."""
    try:
        return doc.properties.builtin()
    except ComError:
        return {}


def _header_footer_texts(doc: Document) -> list[str]:
    """Every header/footer story's text across all sections (for a notice presence scan).

    Includes linked-to-previous stories too — a repeated notice still counts as
    present. Returns raw texts; the caller casefolds for matching."""
    texts: list[str] = []
    try:
        for section in doc.sections:
            for which in _HF_WHICH:
                for hf in (section.header(which), section.footer(which)):
                    try:
                        if hf.exists and hf.text:
                            texts.append(hf.text)
                    except ComError:
                        continue
    except ComError:
        return texts
    return texts


def _body_text(doc: Document) -> str:
    """The document body's text (via the `.com` escape hatch), or `""` on failure."""
    try:
        return str(doc.com.Content.Text or "")
    except (ComError, AttributeError):
        return ""


def _notice_present(doc: Document, text: str) -> bool:
    """Whether `text` appears (case-insensitively) in any header/footer or the body."""
    needle = text.casefold()
    if needle in _body_text(doc).casefold():
        return True
    return any(needle in hf.casefold() for hf in _header_footer_texts(doc))


def _check_document_properties_filled(
    doc: Document, span: Span | None, profile: Profile
) -> Iterator[Finding]:
    """A required built-in property (Title / Author by default) is empty or unset.
    Policy (off until a profile enables it); the profile may override the checked
    set via `required`. Report-only — filling a property needs a value the linter
    doesn't have. Document-global, so `within` doesn't scope it."""
    configured = profile.config_for("document-properties-filled").get("required")
    required = (
        [str(name) for name in configured]
        if isinstance(configured, list) and configured
        else list(_DEFAULT_REQUIRED_PROPS)
    )
    builtin = _builtin_properties(doc)
    for name in required:
        if str(builtin.get(name, "") or "").strip():
            continue
        yield Finding(
            rule="document-properties-filled",
            kind="policy",
            severity="info",
            anchor_id="start",
            message=f"Document property {name!r} is empty; set it before hand-off.",
            fixable=False,
            observed=f"{name}=<empty>",
            expected=f"a non-empty {name}",
        )


def _check_confidentiality_notice(
    doc: Document, span: Span | None, profile: Profile
) -> Iterator[Finding]:
    """A profile-supplied confidentiality notice is not present anywhere in the
    document (header/footer or body). Policy — off until the profile supplies
    `text` (with no configured text there is nothing to police). Report-only
    (inserting the notice adds content). Document-global."""
    text = str(profile.config_for("confidentiality-notice").get("text") or "").strip()
    if not text:
        return  # no configured notice → the rule polices nothing
    if _notice_present(doc, text):
        return
    yield Finding(
        rule="confidentiality-notice",
        kind="policy",
        severity="warning",
        anchor_id="start",
        message=f"Confidentiality notice {text!r} not found in any header/footer or the body.",
        fixable=False,
        observed="notice absent",
        expected=f"the notice {text!r} present",
    )


def _check_copyright_notice(
    doc: Document, span: Span | None, profile: Profile
) -> Iterator[Finding]:
    """A copyright notice is not present anywhere in the document. Policy — off until
    the profile enables it; the expected string is the profile's `text`, defaulting
    to the `©` symbol. Report-only. Document-global."""
    text = str(profile.config_for("copyright-notice").get("text") or "©").strip()
    if not text:
        return
    if _notice_present(doc, text):
        return
    yield Finding(
        rule="copyright-notice",
        kind="policy",
        severity="warning",
        anchor_id="start",
        message=f"Copyright notice {text!r} not found in any header/footer or the body.",
        fixable=False,
        observed="notice absent",
        expected=f"the notice {text!r} present",
    )


def _own_story_texts(doc: Document, *, is_footer: bool) -> list[tuple[int, str]]:
    """`(section_index, text)` for each section that defines its **own** primary
    header/footer (exists and not linked-to-previous) with non-empty text."""
    out: list[tuple[int, str]] = []
    for section in doc.sections:
        hf = section.footer("primary") if is_footer else section.header("primary")
        try:
            if hf.exists and not hf.linked_to_previous and hf.text.strip():
                out.append((section.index, hf.text.strip()))
        except ComError:
            continue
    return out


def _check_header_footer_consistent(
    doc: Document, span: Span | None, profile: Profile
) -> Iterator[Finding]:
    """The primary header (or footer) text disagrees across the document's own
    (non-linked) sections, where a single running header/footer is usually intended.
    Consistency, off by default (a deliberate per-section header is legitimate).
    Report-only — which text is canonical is a human call. Document-global."""
    try:
        section_count = len(doc.sections)
    except ComError:
        return
    if section_count < 2:
        return  # a single-section document has nothing to compare
    for is_footer, kind in ((False, "header"), (True, "footer")):
        try:
            texts = _own_story_texts(doc, is_footer=is_footer)
        except ComError:
            continue
        if len(texts) < 2:
            continue
        baseline_idx, baseline = texts[0]
        deviating = [(idx, t) for idx, t in texts[1:] if t != baseline]
        if not deviating:
            continue
        first_idx, first_text = deviating[0]
        yield Finding(
            rule="header-footer-consistent",
            kind="consistency",
            severity="info",
            anchor_id=f"{kind}:{first_idx}:primary",
            message=(
                f"Primary {kind} text differs across sections: "
                f"{baseline!r} (section {baseline_idx}) vs {first_text!r} (section {first_idx})."
            ),
            fixable=False,
            observed=f"section {first_idx} {kind}={first_text!r}",
            expected=f"the same {kind} as section {baseline_idx} ({baseline!r})",
        )


def _check_draft_watermark_present(
    doc: Document, span: Span | None, profile: Profile
) -> Iterator[Finding]:
    """A text watermark (a leftover DRAFT / CONFIDENTIAL stamp) is still on the
    document. Off by default — a watermark is only a defect once you're finalising —
    so it also carries the `finalization` tag. Report-only (removing it is content
    loss; `remove_watermark` is the opt-in fix). Document-global."""
    try:
        mark = doc.watermark()
    except ComError:
        return
    if mark is None:
        return
    shown = mark.text.strip() or "(blank)"
    yield Finding(
        rule="draft-watermark-present",
        kind="structural",
        severity="warning",
        anchor_id="start",
        message=f"A text watermark ({shown!r}) is still on the document.",
        fixable=False,
        observed=f"watermark={shown!r}",
        expected="no watermark on a final document",
    )


for _rule in (
    Rule(
        id="document-properties-filled",
        kind="policy",
        severity="info",
        tags=("layout",),
        check=_check_document_properties_filled,
        default_on=False,
    ),
    Rule(
        id="confidentiality-notice",
        kind="policy",
        severity="warning",
        tags=("layout", "notices"),
        check=_check_confidentiality_notice,
        default_on=False,
    ),
    Rule(
        id="copyright-notice",
        kind="policy",
        severity="warning",
        tags=("layout", "notices"),
        check=_check_copyright_notice,
        default_on=False,
    ),
    Rule(
        id="header-footer-consistent",
        kind="consistency",
        severity="info",
        tags=("layout",),
        check=_check_header_footer_consistent,
        default_on=False,
    ),
    Rule(
        id="draft-watermark-present",
        kind="structural",
        severity="warning",
        tags=("layout", "finalization"),
        check=_check_draft_watermark_present,
        default_on=False,
    ),
):
    _register_rule(_rule)
