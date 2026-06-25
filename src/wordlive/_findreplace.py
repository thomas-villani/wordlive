"""LLM-friendly fuzzy plain-text find/replace.

The matching is forgiving of cosmetic differences that show up when LLMs
re-emit text from a Word doc — smart quotes, dashes, NBSP, NFKC variants,
whitespace runs — but produces *original* character offsets so the actual
Word range can be edited without disturbing surrounding formatting.

It is *not* markdown-aware: replacing text inside a Range inherits the
formatting of the first character of the match, which is what Word's native
Find/Replace does too. If you need structured edits ("rewrite this paragraph
but keep these runs italic"), use the raw `.com` Range API instead.
"""

from __future__ import annotations

import re
import unicodedata
from dataclasses import dataclass

_QUOTE_FOLDS = {
    "‘": "'",
    "’": "'",
    "‚": "'",
    "‛": "'",  # single quotes
    "“": '"',
    "”": '"',
    "„": '"',
    "‟": '"',  # double quotes
    "′": "'",
    "″": '"',  # primes
    "«": '"',
    "»": '"',  # guillemets
}

_DASH_FOLDS = {
    "‐": "-",
    "‑": "-",
    "‒": "-",
    "–": "-",
    "—": "-",
    "―": "-",
    "−": "-",
}

_SPACE_FOLDS = {
    " ": " ",  # NBSP
    " ": " ",
    " ": " ",
    " ": " ",
    " ": " ",
    "\t": " ",
    "\v": " ",
    "\f": " ",
    "\r": "\n",
    "": "",  # Word's cell/end-of-row marker
}


def _fold_char(ch: str) -> str:
    """Map one character to its fuzzy-match equivalent."""
    if ch in _QUOTE_FOLDS:
        return _QUOTE_FOLDS[ch]
    if ch in _DASH_FOLDS:
        return _DASH_FOLDS[ch]
    if ch in _SPACE_FOLDS:
        return _SPACE_FOLDS[ch]
    return ch


@dataclass(frozen=True)
class _Normalized:
    """A normalized string and a mapping back to the original offsets.

    `text[i]` is the normalized character; `offsets[i]` is the *first* original
    offset that contributed to it. `offsets[len(text)]` is the original offset
    immediately after the last contributing character — i.e. you can use
    `offsets[match_start]` and `offsets[match_end]` directly as a Word range.
    """

    text: str
    offsets: list[int]


def _normalize(s: str, *, collapse_whitespace: bool = True) -> _Normalized:
    """NFKC + character folds + (optional) whitespace collapse.

    Tracks original offsets so a match in the normalized string maps cleanly
    back to a Word Range over the original text.
    """
    out_chars: list[str] = []
    out_offsets: list[int] = []
    prev_space = True  # collapse leading whitespace, like text.strip() lite

    for i, raw_ch in enumerate(s):
        # NFKC may expand one char into several (e.g. ligatures). Each output
        # char shares the same source offset.
        decomposed = unicodedata.normalize("NFKC", raw_ch)
        for ch in decomposed:
            folded = _fold_char(ch)
            if folded == "":
                continue
            for fch in folded:
                is_space = fch in (" ", "\n")
                if collapse_whitespace and is_space:
                    if prev_space:
                        continue
                    out_chars.append(" ")
                    out_offsets.append(i)
                    prev_space = True
                else:
                    out_chars.append(fch)
                    out_offsets.append(i)
                    prev_space = False

    # Trailing space is harmless but ugly; strip it.
    while collapse_whitespace and out_chars and out_chars[-1] == " ":
        out_chars.pop()
        out_offsets.pop()

    # Sentinel: the original offset one *past the last contributing character*,
    # so a match that runs to the end maps to a half-open range that stops at the
    # last real character — not `len(s)`, which would swallow any trailing chars
    # that folded away to nothing (a paragraph mark `\r`, a cell mark `\x07`, a
    # stripped trailing space). Eating such a mark at a segment boundary fuses
    # the paragraph into whatever follows (e.g. the first cell of an adjacent
    # table). `out_offsets[-1]` is the source index of the last retained
    # character; +1 is the position immediately after it (== `len(s)` whenever
    # the final source character was itself retained, so the common case is
    # unchanged). Empty result → fall back to `len(s)` (unused: no matches).
    out_offsets.append(out_offsets[-1] + 1 if out_offsets else len(s))
    return _Normalized(text="".join(out_chars), offsets=out_offsets)


def normalized_equal(a: str, b: str) -> bool:
    """Whether `a` and `b` are equal under find/replace normalization.

    Used by `find_replace` to verify a resolved Word range matches the located
    text before overwriting it (the fuzzy folds mean a smart-quote vs straight
    round-trip still counts as equal).
    """
    return _normalize(a).text == _normalize(b).text


@dataclass(frozen=True)
class Match:
    """A located occurrence of a `find` string inside a scope's text.

    `start` / `end` are offsets *into the scope range* (not absolute document
    offsets), measured against the original (un-normalized) text. `text` is the
    actual original substring at those offsets — useful both for round-tripping
    formatting and for showing the user what was matched.

    `replacement` is the per-match replacement text, set only in `regex` mode
    where backreferences (`\\1`) expand differently for each hit; `None` for
    `fuzzy` / `literal` (where the caller's single `replace` string applies to
    every match) and for plain `find()` (no replacement at all).
    """

    start: int
    end: int
    text: str
    replacement: str | None = None


# `find`/`find_replace` matching modes. `fuzzy` is the historical default
# (Unicode/whitespace-tolerant); `literal` and `regex` are exact — they skip
# `_normalize` entirely, which is what lets the typography linter express a real
# whitespace/punctuation edit (a fuzzy find of "  " matches every single space).
_MODES = ("fuzzy", "literal", "regex")


def find_matches(
    haystack: str,
    needle: str,
    *,
    mode: str = "fuzzy",
    replacement: str | None = None,
) -> list[Match]:
    """Locate every occurrence of `needle` inside `haystack`.

    - `fuzzy` (default): both sides normalized identically (NFKC, smart-quote /
      dash / NBSP folds, whitespace collapse). Forgiving of cosmetic drift.
    - `literal`: exact substring match, no folding/collapse.
    - `regex`: `needle` is a Python regular expression; `replacement` (if given)
      is expanded per match (`\\1` backreferences), so a single call can collapse
      runs or reorder groups. Zero-width matches are skipped.

    Returns matches in original-offset coordinates of `haystack`. Empty `needle`
    returns no matches. `replacement` is only consulted in `regex` mode.
    """
    if not needle:
        return []
    if mode == "literal":
        return _literal_matches(haystack, needle)
    if mode == "regex":
        return _regex_matches(haystack, needle, replacement)
    if mode != "fuzzy":
        raise ValueError(f"unknown find/replace mode {mode!r}; expected one of {_MODES}")

    norm_h = _normalize(haystack)
    norm_n = _normalize(needle)
    if not norm_n.text:
        return []

    matches: list[Match] = []
    i = 0
    nlen = len(norm_n.text)
    while True:
        j = norm_h.text.find(norm_n.text, i)
        if j == -1:
            break
        start = norm_h.offsets[j]
        end = norm_h.offsets[j + nlen]
        matches.append(Match(start=start, end=end, text=haystack[start:end]))
        i = j + nlen
    return matches


def _literal_matches(haystack: str, needle: str) -> list[Match]:
    """Exact, non-overlapping substring matches (no normalization)."""
    matches: list[Match] = []
    i = 0
    nlen = len(needle)
    while True:
        j = haystack.find(needle, i)
        if j == -1:
            break
        matches.append(Match(start=j, end=j + nlen, text=haystack[j : j + nlen]))
        i = j + nlen
    return matches


def _regex_matches(haystack: str, pattern: str, replacement: str | None) -> list[Match]:
    """Python-regex matches, each carrying its expanded `replacement` if given."""
    try:
        rx = re.compile(pattern)
    except re.error as exc:
        raise ValueError(f"invalid regex {pattern!r}: {exc}") from exc
    matches: list[Match] = []
    for m in rx.finditer(haystack):
        if m.start() == m.end():
            continue  # zero-width: nothing to replace, and would spin the writer
        repl = m.expand(replacement) if replacement is not None else None
        matches.append(Match(start=m.start(), end=m.end(), text=m.group(0), replacement=repl))
    return matches
