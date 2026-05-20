"""RangeAnchor — arbitrary-offset anchors and `range:START-END` resolution."""

from __future__ import annotations

import pytest

import wordlive
from wordlive.exceptions import AnchorNotFoundError


def test_range_anchor_basic_properties(fake_word):
    with wordlive.attach() as word:
        rng = word.documents.active.range(5, 12)
    assert isinstance(rng, wordlive.RangeAnchor)
    assert rng.kind == "range"
    assert rng.start == 5
    assert rng.end == 12
    assert rng.anchor_id == "range:5-12"
    assert rng.name == "range:5-12"


def test_range_anchor_rejects_inverted_offsets(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        with pytest.raises(ValueError):
            doc.range(10, 4)
        with pytest.raises(ValueError):
            doc.range(-1, 4)


def test_range_anchor_reads_text(fake_word):
    # Seed the cached range so the read has something to return.
    fake_word.ActiveDocument.Range(5, 10).Text = "hello"
    with wordlive.attach() as word:
        assert word.documents.active.range(5, 10).text == "hello"


def test_range_anchor_set_text_round_trip(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        with doc.edit("replace range"):
            doc.range(5, 10).set_text("XY")
    # The write lands on the resolved Range. (The fake keys ranges by exact
    # (start, end) and can't model the content-shrink that a real Word Range
    # reflects, so we read at the write coordinates rather than the synced ones.)
    assert fake_word.ActiveDocument.Range(5, 10).Text == "XY"


def test_range_anchor_set_text_syncs_end(fake_word):
    with wordlive.attach() as word:
        rng = word.documents.active.range(5, 10)
        rng.set_text("ABCD")
    # End follows the new length (UTF-16 code units): 5 + 4 == 9.
    assert rng.end == 9
    assert rng.anchor_id == "range:5-9"


def test_anchor_by_id_resolves_range(fake_word):
    with wordlive.attach() as word:
        anchor = word.documents.active.anchor_by_id("range:5-12")
    assert isinstance(anchor, wordlive.RangeAnchor)
    assert anchor.start == 5 and anchor.end == 12


def test_anchor_by_id_range_is_writable(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        with doc.edit("replace via range id"):
            doc.anchor_by_id("range:3-8").set_text("Z")
        assert doc.range(3, 8).text == "Z"


def test_anchor_by_id_range_missing_separator_raises(fake_word):
    with wordlive.attach() as word:
        with pytest.raises(AnchorNotFoundError) as exc_info:
            word.documents.active.anchor_by_id("range:512")
    assert exc_info.value.kind == "range"


def test_anchor_by_id_range_non_numeric_raises(fake_word):
    with wordlive.attach() as word:
        with pytest.raises(AnchorNotFoundError):
            word.documents.active.anchor_by_id("range:a-b")


def test_anchor_by_id_range_inverted_offsets_raises(fake_word):
    with wordlive.attach() as word:
        with pytest.raises(AnchorNotFoundError) as exc_info:
            word.documents.active.anchor_by_id("range:9-2")
    assert exc_info.value.kind == "range"


def test_find_emits_resolvable_range_ids(fake_word):
    # The default fixture content contains "Body text here." at offset 13.
    with wordlive.attach() as word:
        doc = word.documents.active
        matches = doc.find("Body text here")
        assert matches, "expected at least one match"
        anchor = doc.anchor_by_id(matches[0]["anchor_id"])
    assert isinstance(anchor, wordlive.RangeAnchor)
