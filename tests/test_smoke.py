"""Smoke tests that require a real running Word.

Skipped by default (`pytest -m "not smoke"` runs the unit suite). To exercise
them: `pytest -m smoke` on a Windows box with Word installed and *some*
document open.
"""

from __future__ import annotations

import pytest

pytestmark = pytest.mark.smoke


def test_status_returns_at_least_one_doc(real_word):
    docs = real_word.documents.list()
    assert isinstance(docs, list)
    # If Word is running but nothing is open, this is fine — the smoke is
    # really about "does the COM dance hang together end-to-end".
    for d in docs:
        assert "name" in d and "path" in d and "is_active" in d


def test_outline_runs(real_word):
    try:
        doc = real_word.documents.active
    except Exception:
        pytest.skip("No active document open in Word")
    outline = doc.outline()
    assert isinstance(outline, list)
    for item in outline:
        assert {"level", "text", "anchor_id"} <= set(item)


def test_edit_scope_preserves_selection(real_word):
    """Trivial edit scope round-trip — verifies Selection survives a no-op edit."""
    try:
        doc = real_word.documents.active
    except Exception:
        pytest.skip("No active document open in Word")
    sel = real_word.selection.info()
    with doc.edit("wordlive smoke: no-op"):
        pass
    after = real_word.selection.info()
    assert (sel["start"], sel["end"]) == (after["start"], after["end"])
