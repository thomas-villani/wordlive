"""Fuzzy find/replace: normalization, position mapping, and the public API."""

from __future__ import annotations

import pytest

import wordlive
from wordlive._findreplace import find_matches, normalized_equal
from wordlive.exceptions import (
    AmbiguousMatchError,
    AnchorNotFoundError,
    ReplaceVerificationError,
)

# ---------------------------------------------------------------------------
# Normalization + match math (pure, no fake-COM needed)
# ---------------------------------------------------------------------------


def test_find_matches_plain():
    matches = find_matches("hello world", "hello")
    assert len(matches) == 1
    assert (matches[0].start, matches[0].end, matches[0].text) == (0, 5, "hello")


def test_find_matches_collapses_internal_whitespace():
    """Double-space in haystack should still match single-space needle."""
    matches = find_matches("hello  world", "hello world")
    assert len(matches) == 1
    # Match offsets cover the original (double-spaced) span.
    assert matches[0].text == "hello  world"


def test_find_matches_smart_quotes_fold_to_ascii():
    matches = find_matches("She said “hi”.", 'said "hi"')
    assert len(matches) == 1
    # The matched span in the original includes the smart quotes.
    assert matches[0].text == "said “hi”"


def test_find_matches_em_dash_folds_to_hyphen():
    matches = find_matches("a — b", "a - b")
    assert len(matches) == 1
    assert matches[0].text == "a — b"


def test_find_matches_nbsp_folds_to_space():
    matches = find_matches("foo bar", "foo bar")
    assert len(matches) == 1
    assert matches[0].text == "foo bar"


def test_find_matches_multiple_occurrences():
    matches = find_matches("ab ab ab", "ab")
    assert len(matches) == 3
    assert [m.start for m in matches] == [0, 3, 6]


def test_find_matches_end_anchored_excludes_trailing_paragraph_mark():
    """A whole-paragraph match at a segment boundary must stop at the last real
    character, not swallow the trailing paragraph mark.

    Regression: the normalized sentinel used `len(s)`, so when a folded-away
    trailing `\\r` was stripped, `offsets[end]` pointed past it. Replacing that
    span deleted the paragraph break, fusing the paragraph into whatever followed
    (e.g. the first cell of an adjacent table).
    """
    hay = "Revenue grew.\rCosts held flat.\r"
    (m,) = find_matches(hay, "Costs held flat.")
    assert m.text == "Costs held flat."  # no trailing \r
    assert hay[m.start : m.end] == "Costs held flat."
    assert hay[m.end] == "\r"  # the mark is left intact, just past the match


def test_find_matches_end_anchored_excludes_trailing_cell_mark():
    """The same boundary fix for Word's folded-away cell/end-of-row marker."""
    hay = "Total\x07"  # \x07 folds to "" in normalization
    (m,) = find_matches(hay, "Total")
    assert (m.start, m.end, m.text) == (0, 5, "Total")


def test_find_matches_no_match():
    assert find_matches("hello", "xyz") == []


def test_find_matches_empty_needle_yields_nothing():
    assert find_matches("hello", "") == []


# ---------------------------------------------------------------------------
# Document.find / find_replace
# ---------------------------------------------------------------------------


def test_document_find_locates_text_in_content(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        matches = doc.find("Body text here")
    assert len(matches) == 1
    assert matches[0]["start"] == 13
    assert matches[0]["end"] == 27  # "Body text here" is 14 chars
    assert matches[0]["anchor_id"] == "range:13-27"


def test_document_find_returns_empty_on_no_match(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        assert doc.find("nonexistent phrase") == []


def test_document_find_replace_single_match(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        with doc.edit("fr"):
            applied = doc.find_replace("Body text here", "Replaced")
    assert len(applied) == 1
    assert applied[0]["start"] == 13
    # Verify a Range was created over the match span.
    fake_word.ActiveDocument.Range.assert_any_call(13, 27)


def test_document_find_replace_zero_matches_raises(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        with pytest.raises(AnchorNotFoundError) as exc_info:
            doc.find_replace("not in doc", "x")
    assert exc_info.value.kind == "find"


def test_document_find_replace_ambiguous_raises(fake_word, monkeypatch):
    """Two matches without --all/--occurrence → AmbiguousMatchError."""
    # Patch Content.Text to have two occurrences of the same phrase.
    fake_word.ActiveDocument.Content.Text = "alpha beta alpha beta"
    fake_word.ActiveDocument.Content.End = len("alpha beta alpha beta")

    with wordlive.attach() as word:
        doc = word.documents.active
        with pytest.raises(AmbiguousMatchError) as exc_info:
            doc.find_replace("alpha", "X")
    assert len(exc_info.value.matches) == 2


def test_document_find_replace_all_applies_to_every_match(fake_word):
    fake_word.ActiveDocument.Content.Text = "alpha beta alpha beta"
    fake_word.ActiveDocument.Content.End = len("alpha beta alpha beta")

    with wordlive.attach() as word:
        doc = word.documents.active
        with doc.edit("fr"):
            applied = doc.find_replace("alpha", "X", all=True)
    assert len(applied) == 2


def test_document_find_replace_occurrence_picks_nth(fake_word):
    fake_word.ActiveDocument.Content.Text = "alpha beta alpha beta"
    fake_word.ActiveDocument.Content.End = len("alpha beta alpha beta")

    with wordlive.attach() as word:
        doc = word.documents.active
        with doc.edit("fr"):
            applied = doc.find_replace("alpha", "X", occurrence=2)
    assert len(applied) == 1
    # Second occurrence starts at index 11.
    assert applied[0]["start"] == 11


def test_document_find_replace_occurrence_out_of_range_raises(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        with pytest.raises(AnchorNotFoundError):
            doc.find_replace("Body text here", "x", occurrence=5)


# ---------------------------------------------------------------------------
# Terminal paragraph clamp (Fix 1a)
# ---------------------------------------------------------------------------


def test_find_replace_last_paragraph_targets_before_final_mark(fake_word):
    # A real document always ends with an undeletable paragraph mark; a match in
    # the last paragraph must target up to (not including) that mark or Word
    # raises COM 0x80020009.
    fake_word.ActiveDocument.Content.Text = "alpha beta\r"
    fake_word.ActiveDocument.Content.End = len("alpha beta\r")  # 11; \r sits at 10
    with wordlive.attach() as word:
        doc = word.documents.active
        with doc.edit("fr"):
            applied = doc.find_replace("beta", "X")
    assert len(applied) == 1
    # Target ends at 10 (the \r position), never 11 (== Content.End).
    fake_word.ActiveDocument.Range.assert_any_call(6, 10)


def test_find_replace_clamps_target_away_from_content_end(fake_word):
    # Degenerate: the match offset reaches Content.End (this fake string has no
    # trailing mark). The clamp must pull the write target back to End-1 rather
    # than straddle the final paragraph mark.
    fake_word.ActiveDocument.Content.Text = "alpha beta"
    fake_word.ActiveDocument.Content.End = 10
    with wordlive.attach() as word:
        doc = word.documents.active
        with doc.edit("fr"):
            doc.find_replace("beta", "X")
    calls = [c.args for c in fake_word.ActiveDocument.Range.call_args_list]
    assert (6, 10) not in calls  # never the raw, crashing span
    assert (6, 9) in calls  # clamped to End-1


# ---------------------------------------------------------------------------
# Replace verification safety-net (Fix 2a)
# ---------------------------------------------------------------------------


def test_find_replace_refuses_when_resolved_target_mismatches(fake_word):
    # Simulate an offset map that drifted onto unrelated, non-empty text (the
    # table-cell corruption signature): the write must be refused, not applied.
    fake_word.ActiveDocument.Content.Text = "alpha beta"
    fake_word.ActiveDocument.Content.End = 10
    fake_word.ActiveDocument.Range(0, 5).Text = "WRONG"
    with wordlive.attach() as word:
        doc = word.documents.active
        with pytest.raises(ReplaceVerificationError) as exc_info:
            with doc.edit("fr"):
                doc.find_replace("alpha", "Z")
    assert exc_info.value.expected == "alpha"
    assert exc_info.value.resolved == "WRONG"


def test_find_replace_proceeds_when_resolved_target_empty(fake_word):
    # An empty resolved target (the fake's default, or a genuinely empty range)
    # is treated as unverifiable, not a mismatch — the replace still applies.
    fake_word.ActiveDocument.Content.Text = "alpha beta"
    fake_word.ActiveDocument.Content.End = 10
    with wordlive.attach() as word:
        doc = word.documents.active
        with doc.edit("fr"):
            applied = doc.find_replace("alpha", "Z")
    assert len(applied) == 1
    assert fake_word.ActiveDocument.Range(0, 5).Text == "Z"


def test_normalized_equal_folds_cosmetic_differences():
    assert normalized_equal("say “hi”", 'say "hi"')
    assert normalized_equal("a — b", "a - b")
    assert not normalized_equal("alpha", "beta")


def test_replace_verification_error_classifies_distinctly():
    from wordlive.cli.main import _exit_for
    from wordlive.exceptions import classify

    exc = ReplaceVerificationError("None", "None", "Looks", anchor_id="range:5-9")
    assert classify(exc) == ("replace_verification", False)
    assert _exit_for(exc) == 1  # generic exit code, not anchor/ambiguous
