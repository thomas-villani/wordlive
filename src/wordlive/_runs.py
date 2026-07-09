"""Inline runs — rich text within a paragraph, shared by `insert_paragraph`
(structured `runs`) and `insert_block` (per-item `text`/`runs`).

A *run* is a contiguous span of one paragraph carrying its own character
formatting (`bold`/`italic`/`underline`/`style`). Two ways to express runs reach
this module, and both normalise to a `list[Run]`:

- **Structured** — `runs: [{"text": "Fast", "bold": true}, {"text": " — quick"}]`.
  Unambiguous, supports a per-run character `style`. The canonical form.
- **Markdown sugar** — a plain `text` string with a deliberately tiny markup:
  `**bold**`, `*italic*`, `***bold italic***`, `` `code` ``, with a backslash
  escape before any character `to_markdown` escapes. The LLM-native path for the
  common "bold lead-in" bullet, and for the identifiers and paths an agent
  backticks by reflex; it desugars to the same `Run` list. Capped at
  bold/italic/code on purpose — anything richer wants `runs`.

This module is COM-free and pure so the parser and validation are unit-testable
off-Windows; the actual character formatting is applied in `_anchors.insert_block`.
"""

from __future__ import annotations

import re
from dataclasses import dataclass
from typing import Any

from .exceptions import OpError

# Run fields an LLM/CLI payload may set. `text` is mandatory; the rest are
# tri-state (None = "leave as inherited", not False) so a run only touches the
# attributes it names.
_RUN_KEYS = ("text", "bold", "italic", "underline", "code", "style")

# The font a `` `code` `` span is rendered in. Direct character formatting, not a
# character style — the same choice `**bold**` makes (`Font.Bold`, not `Strong`),
# so `insert_markdown` never mutates the document's style gallery.
CODE_FONT = "Consolas"


@dataclass(frozen=True)
class Run:
    """One inline span of a paragraph plus its character formatting.

    `bold`/`italic`/`underline`/`code` are tri-state: ``None`` leaves the
    attribute as inherited, ``True``/``False`` set it explicitly. `code` renders
    the span in `CODE_FONT`. `style` names a (character) style applied to the
    span.
    """

    text: str
    bold: bool | None = None
    italic: bool | None = None
    underline: bool | None = None
    style: str | None = None
    code: bool | None = None

    def formatted(self) -> bool:
        """Whether this run carries any formatting worth a COM round-trip."""
        return any(
            v is not None for v in (self.bold, self.italic, self.underline, self.style, self.code)
        )


# The characters `_export._escape_inline` backslash-escapes on the way out, and
# so exactly the ones this parser must unescape on the way back in. The two sets
# must stay in lockstep: that is what makes `to_markdown` -> `insert_markdown` a
# fixed point. While they were out of step (only `*` and `\` were unescaped) each
# round-trip accreted one backslash per special character, without bound.
_ESCAPABLE = "\\`*{}[]#_"

# Escaped characters are lifted to private-use sentinels *before* the delimiter
# scan, so `\*` never reads as emphasis and is restored to a literal `*` after.
_ESC_SENTINEL = {ch: chr(0xE000 + i) for i, ch in enumerate(_ESCAPABLE)}
_ESC_RESTORE = {v: k for k, v in _ESC_SENTINEL.items()}

# One left-to-right pass, so a `\\` consumes its own backslash and leaves a
# following delimiter live: `\\*x*` is a literal backslash then an italic `x`.
_ESC_RX = re.compile(r"\\([" + re.escape(_ESCAPABLE) + r"])")

# Matched-pair patterns, longest delimiter first so `***x***` wins over `**`/`*`
# at the same position. `.+?` is non-greedy and crosses no newline by default,
# which is irrelevant here (run text is single-paragraph) but keeps matches tight.
_EMPHASIS = (
    (re.compile(r"\*\*\*(.+?)\*\*\*"), True, True),
    (re.compile(r"\*\*(.+?)\*\*"), True, False),
    (re.compile(r"\*(.+?)\*"), False, True),
)

# A code span: a backtick fence, the shortest content that reaches a matching
# fence of the same length, and the closing fence. The variable fence (`` ` ``
# vs ```` `` ````) is what lets a span contain literal backticks. Code binds
# tighter than emphasis (as in CommonMark), so the scan splits on these first
# and only emphasis-parses what falls between — otherwise `*a `b` c*` would let
# the italic swallow the span and strand its backticks as literal text.
_CODE = re.compile(r"(`+)(.+?)\1")


def _restore(s: str) -> str:
    return "".join(_ESC_RESTORE.get(ch, ch) for ch in s)


def _strip_code_pad(content: str) -> str:
    """Undo the one-space pad a fence uses to hold backtick-edged content.

    CommonMark: if a code span's content both begins and ends with a space and is
    not all spaces, one space is stripped from each end. That is what lets
    ``` `` `x` `` ``` mean ``` `x` ```. The inverse of `_export._code_span`'s pad.
    """
    if len(content) >= 2 and content[0] == " " and content[-1] == " " and content.strip():
        return content[1:-1]
    return content


def _parse_emphasis(work: str) -> list[Run]:
    """Scan one code-free segment for `*` emphasis. `work` has escapes lifted."""
    runs: list[Run] = []

    def emit(span: str, bold: bool, italic: bool) -> None:
        if span:
            runs.append(Run(text=_restore(span), bold=bold or None, italic=italic or None))

    pos = 0
    while pos < len(work):
        best: tuple[re.Match[str], bool, bool] | None = None
        for rx, bold, italic in _EMPHASIS:
            m = rx.search(work, pos)
            # Earliest start wins; ties keep the first (longest-delimiter) pattern
            # because the replacement test is strict `<`.
            if m and (best is None or m.start() < best[0].start()):
                best = (m, bold, italic)
        if best is None:
            emit(work[pos:], False, False)
            break
        m, bold, italic = best
        if m.start() > pos:
            emit(work[pos : m.start()], False, False)
        emit(m.group(1), bold, italic)
        pos = m.end()
    return runs


def parse_markup(text: str) -> list[Run]:
    """Desugar the tiny inline markdown into a `list[Run]`.

    Supports `**bold**`, `*italic*`, `***bold italic***`, `` `code` `` (rendered
    in `CODE_FONT`), and a backslash escape before any of ``\\`` ` ``*{}[]#_`` —
    the exact set `to_markdown` emits, so a read/write round-trip is a fixed
    point.

    Code spans bind tighter than emphasis and their content is literal, so
    ``` `b*c` ``` is one code run reading ``b*c``, not an italic. Use a longer
    fence to embed a backtick (```` ``a `b` c`` ````). Because runs are flat,
    emphasis does **not** reach across a code span: in ``` *a `b` c* ``` the
    asterisks stay literal (they land in different segments). Wrap the emphasis
    inside the code-free part, or use structured `runs`, if you need both.

    Unmatched delimiters stay literal — `"a * b"` is one plain run, not a
    dangling emphasis. Plain text with no markup returns a single unformatted
    run, so routing every insert through this parser is behaviour-preserving for
    ordinary text.
    """
    # Lift escapes out of the way so the scan never sees an escaped delimiter.
    # An escaped backtick is now a sentinel, so it can't open a code span either.
    work = _ESC_RX.sub(lambda m: _ESC_SENTINEL[m.group(1)], text)
    runs: list[Run] = []
    pos = 0
    for m in _CODE.finditer(work):
        if m.start() > pos:
            runs.extend(_parse_emphasis(work[pos : m.start()]))
        # Code content is literal — no emphasis scan, only escape restoration.
        runs.append(Run(text=_restore(_strip_code_pad(m.group(2))), code=True))
        pos = m.end()
    if pos < len(work):
        runs.extend(_parse_emphasis(work[pos:]))
    if not runs:
        # All-empty input (or "") — emit one empty-text run so callers always get
        # a paragraph, matching insert_paragraph("") semantics.
        runs.append(Run(text=_restore(work)))
    return runs


def _run_from_dict(raw: Any, where: str) -> Run:
    """Validate one structured-run dict into a `Run` (`OpError` on bad shape)."""
    if not isinstance(raw, dict):
        raise OpError(f"{where}: each run must be an object; got {type(raw).__name__}")
    stray = [k for k in raw if k not in _RUN_KEYS]
    if stray:
        raise OpError(
            f"{where}: run has unknown field(s): {', '.join(repr(k) for k in stray)}; "
            f"allowed: {', '.join(_RUN_KEYS)}"
        )
    if "text" not in raw or not isinstance(raw["text"], str):
        raise OpError(f"{where}: each run requires a string 'text'")
    for flag in ("bold", "italic", "underline", "code"):
        if flag in raw and not isinstance(raw[flag], bool):
            raise OpError(f"{where}: run '{flag}' must be a boolean")
    if "style" in raw and raw["style"] is not None and not isinstance(raw["style"], str):
        raise OpError(f"{where}: run 'style' must be a string")
    return Run(
        text=raw["text"],
        bold=raw.get("bold"),
        italic=raw.get("italic"),
        underline=raw.get("underline"),
        code=raw.get("code"),
        style=raw.get("style"),
    )


def runs_from_payload(
    *, text: str | None = None, runs: Any = None, where: str = "item"
) -> list[Run]:
    """Normalise one paragraph's content (``text`` xor ``runs``) to `list[Run]`.

    ``text`` is parsed for markdown sugar; ``runs`` is a list of structured run
    dicts. Exactly one must be supplied — both or neither raises `OpError`,
    mirroring `insert_image`'s exactly-one-of contract.
    """
    if (text is None) == (runs is None):
        raise OpError(f"{where} requires exactly one of 'text' or 'runs'")
    if runs is not None:
        if not isinstance(runs, list) or not runs:
            raise OpError(f"{where}: 'runs' must be a non-empty list")
        return [_run_from_dict(r, where) for r in runs]
    assert text is not None
    return parse_markup(text)


def normalize_block_items(items: Any) -> list[tuple[list[Run], str | None]]:
    """Validate `insert_block` items into `[(runs, paragraph_style), …]`.

    Each item is either a plain string (sugar for ``{"text": …}``) or a dict
    with ``text`` xor ``runs`` plus an optional paragraph ``style``. Raises
    `OpError` on a malformed item so a bad block fails before any mutation.
    """
    if not isinstance(items, list) or not items:
        raise OpError("insert_block requires a non-empty 'items' list")
    out: list[tuple[list[Run], str | None]] = []
    for i, item in enumerate(items, start=1):
        where = f"items[{i}]"
        if isinstance(item, str):
            out.append((parse_markup(item), None))
            continue
        if not isinstance(item, dict):
            raise OpError(f"{where} must be a string or object; got {type(item).__name__}")
        stray = [k for k in item if k not in ("text", "runs", "style")]
        if stray:
            raise OpError(
                f"{where} has unknown field(s): {', '.join(repr(k) for k in stray)}; "
                "allowed: text, runs, style"
            )
        style = item.get("style")
        if style is not None and not isinstance(style, str):
            raise OpError(f"{where}: 'style' must be a string")
        runs = runs_from_payload(text=item.get("text"), runs=item.get("runs"), where=where)
        out.append((runs, style))
    return out


def runs_to_text(runs: list[Run]) -> str:
    """The plain concatenated text of a run list — the bulk-insert payload."""
    return "".join(r.text for r in runs)
