"""Nearest-name suggestions for unknown command / op / action names.

The dispatch surfaces (`word_read`, `word_write`, `exec` ops) have large flat
vocabularies — 45 read commands, 93 write commands, ~150 ops. Naming the whole
set back at a caller that missed by one token is complete but useless. These
helpers turn a miss into `did you mean 'format_info'?`.

Plain `difflib` is not good enough here, because command names share
low-information prefixes: it ranks `read_format` closest to `read_image` (the
shared `read_`) when the caller plainly wanted `format_info`. So a name's score
blends two signals in equal measure:

* **character similarity** — `difflib`'s ratio, which catches typos and elisions
  (`paragraph` -> `paragraphs`);
* **IDF-weighted token overlap** — split on `_`, then weight each shared token by
  how *rare* it is across the vocabulary. `read` appears in five read commands
  and says little; `format` appears in one and says a lot.

A candidate that wholly contains the miss (or vice versa) short-circuits to a
high score, so `markdown` finds `to_markdown` and `apply_list` finds `list`.
Names are compared case- and separator-insensitively, so `FIND` and `list info`
resolve like `find` and `list_info`.

Suggestions are deliberately capped and thresholded: three confident-looking
wrong answers are worse than none, so a miss with no near neighbour
(`xyzzy`) yields an empty list and the caller falls back to a pointer at the
full vocabulary.
"""

from __future__ import annotations

import math
import re
from collections.abc import Iterable, Sequence
from difflib import SequenceMatcher

__all__ = ["did_you_mean", "or_list", "unknown_value_message"]

_SEPARATORS = re.compile(r"[^a-z0-9]+")

# Below this blended score a candidate is noise, not a suggestion.
_CUTOFF = 0.35
# Character similarity and token overlap are weighted equally; see the module
# docstring for why neither alone is sufficient.
_RATIO_WEIGHT = 0.5
_TOKEN_WEIGHT = 0.5
# A containment match (`markdown` in `to_markdown`) outranks any blended score,
# scaled by how much of the candidate the miss actually covers — otherwise
# `paragraph` would rank `find_paragraphs` alongside the `paragraphs` it meant.
_CONTAINMENT_BASE = 0.75
_CONTAINMENT_SPAN = 0.15


def _normalise(value: str) -> str:
    """Casefold and collapse every non-alphanumeric run to a single `_`."""
    return _SEPARATORS.sub("_", value.strip().casefold()).strip("_")


def _tokens(normalised: str) -> set[str]:
    return {tok for tok in normalised.split("_") if tok}


def _inverse_frequencies(candidates: Sequence[str]) -> dict[str, float]:
    """Map each token in the vocabulary to its inverse document frequency."""
    frequency: dict[str, int] = {}
    for candidate in candidates:
        for token in _tokens(_normalise(candidate)):
            frequency[token] = frequency.get(token, 0) + 1
    total = len(candidates)
    return {tok: math.log(1 + total / (1 + count)) for tok, count in frequency.items()}


def did_you_mean(value: str, candidates: Iterable[str], *, limit: int = 3) -> list[str]:
    """Return up to `limit` candidates plausibly meant by `value`, best first.

    Empty when nothing scores above the noise floor. An exact match modulo case
    and separators (`Read-Text` for `read_text`) short-circuits to that single
    name, since there is nothing to disambiguate.
    """
    pool = list(candidates)
    target = _normalise(value)
    if not target or not pool:
        return []

    exact = [c for c in pool if _normalise(c) == target]
    if exact:
        return exact[:limit]

    idf = _inverse_frequencies(pool)
    target_tokens = _tokens(target)
    scored: list[tuple[float, str]] = []
    for candidate in pool:
        normalised = _normalise(candidate)
        if target in normalised or normalised in target:
            covered = min(len(target), len(normalised)) / max(len(target), len(normalised))
            scored.append((_CONTAINMENT_BASE + _CONTAINMENT_SPAN * covered, candidate))
            continue
        candidate_tokens = _tokens(normalised)
        union = target_tokens | candidate_tokens
        weight = sum(idf.get(tok, 0.0) for tok in union)
        overlap = (
            sum(idf.get(tok, 0.0) for tok in target_tokens & candidate_tokens) / weight
            if weight
            else 0.0
        )
        ratio = SequenceMatcher(None, target, normalised).ratio()
        score = _RATIO_WEIGHT * ratio + _TOKEN_WEIGHT * overlap
        if score >= _CUTOFF:
            scored.append((score, candidate))

    # Sort by score, then name, so equal scores rank deterministically.
    scored.sort(key=lambda pair: (-pair[0], pair[1]))
    return [name for _, name in scored[:limit]]


def or_list(names: Sequence[str]) -> str:
    """Render names as a quoted English list: `'a', 'b', or 'c'`."""
    quoted = [repr(str(n)) for n in names]
    if len(quoted) <= 1:
        return "".join(quoted)
    if len(quoted) == 2:
        return f"{quoted[0]} or {quoted[1]}"
    return f"{', '.join(quoted[:-1])}, or {quoted[-1]}"


def unknown_value_message(
    label: str,
    value: str,
    candidates: Sequence[str],
    *,
    fallback: str | None = None,
) -> str:
    """Build the standard `unknown <label> 'x'; did you mean …?` error text.

    `fallback` names where the full vocabulary lives, and is appended whenever
    the suggestions are partial (or absent) so the caller is never left with a
    dead end. Vocabularies of five or fewer names are simply listed in full —
    a "did you mean" is pointless when the caller can see every option.
    """
    message = f"unknown {label} {value!r}"
    if len(candidates) <= 5:
        return f"{message}; expected {or_list(candidates)}" if candidates else message

    suggestions = did_you_mean(value, candidates)
    if not suggestions:
        pointer = f" — {fallback}" if fallback else ""
        return f"{message}; {len(candidates)} valid {label}s exist{pointer}"
    message += f"; did you mean {or_list(suggestions)}?"
    # A lone confident suggestion needs no escape hatch; a shortlist does.
    if fallback and len(suggestions) > 1:
        message += f" ({len(candidates)} valid — {fallback})"
    return message
