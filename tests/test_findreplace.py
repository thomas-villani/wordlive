"""Fuzzy find/replace: normalization, position mapping, and the public API."""

from __future__ import annotations

import pytest

import wordlive
from wordlive._findreplace import find_matches
from wordlive.exceptions import AmbiguousMatchError, AnchorNotFoundError

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
