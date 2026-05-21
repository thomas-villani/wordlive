"""Paragraph anchors (para:N) + doc.paragraphs collection."""

from __future__ import annotations

import pytest

import wordlive
from wordlive import Paragraph
from wordlive.exceptions import AnchorNotFoundError


def test_paragraphs_len(fake_word):
    with wordlive.attach() as word:
        assert len(word.documents.active.paragraphs) == 3


def test_paragraphs_list_shape(fake_word):
    with wordlive.attach() as word:
        rows = word.documents.active.paragraphs.list()
    assert [r["anchor_id"] for r in rows] == ["para:1", "para:2", "para:3"]
    assert [r["is_heading"] for r in rows] == [True, False, True]
    assert [r["level"] for r in rows] == [1, 10, 2]
    assert [r["text"] for r in rows] == ["Introduction", "Body text here.", "Risks"]
    # Offsets are emitted for range:START-END targeting.
    assert (rows[1]["start"], rows[1]["end"]) == (13, 29)


def test_getitem_returns_paragraph(fake_word):
    with wordlive.attach() as word:
        para = word.documents.active.paragraphs[2]
        assert isinstance(para, Paragraph)
        assert para.anchor_id == "para:2"
        assert para.index == 2
        assert para.is_heading is False
        assert para.level == 10
        assert para.text == "Body text here."


def test_heading_paragraph_is_flagged(fake_word):
    with wordlive.attach() as word:
        para = word.documents.active.paragraphs[1]
        assert para.is_heading is True
        assert para.text == "Introduction"


def test_getitem_out_of_range_raises(fake_word):
    with wordlive.attach() as word:
        with pytest.raises(AnchorNotFoundError) as exc_info:
            _ = word.documents.active.paragraphs[99]
    assert exc_info.value.kind == "paragraph"


def test_getitem_bool_rejected(fake_word):
    with wordlive.attach() as word:
        with pytest.raises(TypeError):
            _ = word.documents.active.paragraphs[True]


def test_iterate_yields_paragraphs(fake_word):
    with wordlive.attach() as word:
        ids = [p.anchor_id for p in word.documents.active.paragraphs]
    assert ids == ["para:1", "para:2", "para:3"]


def test_anchor_by_id_resolves_para(fake_word):
    with wordlive.attach() as word:
        anchor = word.documents.active.anchor_by_id("para:2")
        assert isinstance(anchor, Paragraph)
        assert anchor.text == "Body text here."


def test_anchor_by_id_para_bad_index_raises_on_use(fake_word):
    with wordlive.attach() as word:
        anchor = word.documents.active.anchor_by_id("para:99")
        with pytest.raises(AnchorNotFoundError) as exc_info:
            _ = anchor.text
    assert exc_info.value.kind == "paragraph"


def test_para_and_heading_share_index_space(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        # para:1 and heading:1 are the same heading paragraph.
        assert doc.anchor_by_id("para:1").text == doc.anchor_by_id("heading:1").text


def test_at_maps_offset_to_paragraph(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        assert doc.paragraphs.at(0).anchor_id == "para:1"
        assert doc.paragraphs.at(15).anchor_id == "para:2"
        assert doc.paragraphs.at(9999) is None


def test_set_text_preserves_trailing_mark(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        doc.paragraphs[2].set_text("Rewritten")
        # para:2 spans 13-29; the inner write targets 13-28 (mark preserved).
        assert doc.com.Range(13, 28).Text == "Rewritten"


def test_insert_paragraph_after_on_paragraph(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        doc.paragraphs[1].insert_paragraph_after("New body")
        # para:1 (Introduction) ends at 13; new paragraph lands there.
        assert doc.com.Range(13, 13).Text == "New body\r"


def test_apply_style_inherited_on_paragraph(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        # Inherited from Anchor — validates the style then writes Range.Style.
        doc.paragraphs[2].apply_style("Heading 2")
        assert doc.paragraphs[2]._range().Style == doc.com.Styles("Heading 2")
