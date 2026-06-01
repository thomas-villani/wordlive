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

    # Sentinel: one-past-the-last so callers can use offsets[end] as a half-open
    # right boundary even when the match runs to the end of the string.
    out_offsets.append(len(s))
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
    """

    start: int
    end: int
    text: str


def find_matches(haystack: str, needle: str) -> list[Match]:
    """Locate every fuzzy occurrence of `needle` inside `haystack`.

    Both sides are normalized identically (NFKC, smart-quote / dash / NBSP
    folds, whitespace collapse). Returns matches in original-offset
    coordinates of `haystack`. Empty `needle` returns no matches.
    """
    if not needle:
        return []
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
