"""Proofing — spelling / grammar errors and readability statistics.

`doc.proofing()` runs Word's own proofing tools over the document and reports
what they find: the spelling errors (count plus each misspelled run with a
`range:START-END` id), the grammar errors, and the readability statistics
(Flesch Reading Ease, Flesch-Kincaid Grade Level, passive-sentence percentage,
averages). It's a read — nothing is changed — but it does ask Word to (re)check
the document, so it's heavier than a plain `stats()`.

Each error carries an anchor id so an agent can act on it: feed the
`range:START-END` into `read` to see context or `comments.add` to flag it.
"""

from __future__ import annotations

import re
from typing import TYPE_CHECKING, Any

from . import _com

if TYPE_CHECKING:
    from ._document import Document

# Cap how many individual errors are listed (the full tally is always in
# `count`). A document with hundreds of misspellings shouldn't return a payload
# dominated by them — the count tells the agent the scale; the list is a sample.
_MAX_ERRORS = 100


def _slug(name: str) -> str:
    """A snake_case key for a readability-stat name (`"Flesch-Kincaid Grade Level"`)."""
    return re.sub(r"_+", "_", re.sub(r"[^a-z0-9]+", "_", name.strip().lower())).strip("_")


def _errors(doc: Document, attr: str) -> dict[str, Any]:
    """Read a `SpellingErrors` / `GrammaticalErrors` collection into a summary dict.

    Returns `{count, errors:[{text, anchor_id, para}]}`. Accessing the collection
    triggers Word's checker, which can raise if proofing is disabled or the
    document is protected — that's reported as `count: None` with no errors
    rather than failing the whole read.
    """
    try:
        with _com.translate_com_errors():
            collection = getattr(doc.com, attr)
            count = int(collection.Count)
            ranges = [collection.Item(i) for i in range(1, min(count, _MAX_ERRORS) + 1)]
    except Exception:
        return {"count": None, "errors": []}
    errors: list[dict[str, Any]] = []
    for rng in ranges:
        with _com.translate_com_errors():
            try:
                start, end = int(rng.Start), int(rng.End)
            except Exception:
                start, end = None, None
            text = str(rng.Text or "").rstrip("\r\n\x07")
        para = doc.paragraphs.at(start) if start is not None else None
        errors.append(
            {
                "text": text,
                "anchor_id": f"range:{start}-{end}" if start is not None else None,
                "para": para.anchor_id if para is not None else None,
            }
        )
    return {"count": count, "errors": errors}


def _readability(doc: Document) -> dict[str, float]:
    """Word's `ReadabilityStatistics` as a `{snake_case_name: value}` dict.

    Accessing this runs a full grammar+readability pass; if proofing is
    unavailable it raises, and we report an empty dict rather than failing.
    """
    out: dict[str, float] = {}
    try:
        with _com.translate_com_errors():
            for stat in doc.com.ReadabilityStatistics:
                name, value = str(stat.Name), float(stat.Value)
                out[_slug(name)] = value
    except Exception:
        return {}
    return out


def read_proofing(doc: Document) -> dict[str, Any]:
    """`{spelling, grammar, readability}` — Word's proofing read for a document.

    `spelling` / `grammar` are `{count, errors:[{text, anchor_id, para}]}` (the
    error list capped, the count exact); `readability` is the snake_cased
    readability statistics. See [`Document.proofing`][wordlive.Document.proofing].
    """
    return {
        "spelling": _errors(doc, "SpellingErrors"),
        "grammar": _errors(doc, "GrammaticalErrors"),
        "readability": _readability(doc),
    }
