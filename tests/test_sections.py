"""Sections, headers, footers — anchors over header/footer ranges + doc.sections."""

from __future__ import annotations

import pytest

import wordlive
from wordlive.exceptions import AnchorNotFoundError


def test_sections_len(fake_word):
    with wordlive.attach() as word:
        assert len(word.documents.active.sections) == 1


def test_header_read_seeded_text(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        assert doc.sections[1].header().text == "Confidential Draft"


def test_footer_read_seeded_text(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        assert doc.sections[1].footer().text == "Page 1"


def test_header_anchor_id_and_kind(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        hf = doc.sections[1].header()
        assert hf.anchor_id == "header:1:primary"
        assert hf.kind == "header"


def test_footer_anchor_id_with_which(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        hf = doc.sections[1].footer("first")
        assert hf.anchor_id == "footer:1:first"
        assert hf.kind == "footer"
        assert hf.which == "first"


def test_header_set_text_round_trip(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        with doc.edit("header"):
            doc.sections[1].header().set_text("ACME Corporation")
        assert doc.sections[1].header().text == "ACME Corporation"


def test_footer_set_text_round_trip(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        with doc.edit("footer"):
            doc.sections[1].footer().set_text("Page 2 of 5")
        assert doc.sections[1].footer().text == "Page 2 of 5"


def test_unseeded_which_is_empty(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        # Only the primary header was seeded; first/even default to empty.
        assert doc.sections[1].header("first").text == ""
        assert doc.sections[1].header("even").text == ""


def test_unknown_which_raises_value_error(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        with pytest.raises(ValueError):
            doc.sections[1].header("sideways")


def test_page_setup_shape(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        ps = doc.sections[1].page_setup()
    assert ps["orientation"] == "portrait"
    assert ps["page_width"] == 612.0
    assert {
        "orientation",
        "top_margin",
        "bottom_margin",
        "left_margin",
        "right_margin",
        "page_width",
        "page_height",
    } == set(ps)


def test_sections_list_shape(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        rows = doc.sections.list()
    assert rows[0]["index"] == 1
    assert "page_setup" in rows[0]


def test_sections_getitem_out_of_range_raises(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        with pytest.raises(AnchorNotFoundError) as exc_info:
            _ = doc.sections[9]
    assert exc_info.value.kind == "section"


def test_sections_getitem_bool_rejected(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        with pytest.raises(TypeError):
            _ = doc.sections[True]


def test_header_exists_and_linked(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        hf = doc.sections[1].header()
        assert hf.exists is True
        assert hf.linked_to_previous is False


def test_anchor_by_id_header_resolves_and_writes(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        hf = doc.anchor_by_id("header:1:primary")
        assert hf.kind == "header"
        with doc.edit("hdr"):
            hf.set_text("Via Id")
        assert doc.sections[1].header().text == "Via Id"


def test_anchor_by_id_footer_first(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        hf = doc.anchor_by_id("footer:1:first")
        assert hf.kind == "footer"
        assert hf.which == "first"


def test_anchor_by_id_bad_which_raises(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        with pytest.raises(AnchorNotFoundError) as exc_info:
            doc.anchor_by_id("header:1:sideways")
    assert exc_info.value.kind == "header"


def test_anchor_by_id_section_out_of_range_raises(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        with pytest.raises(AnchorNotFoundError) as exc_info:
            doc.anchor_by_id("header:9:primary")
    assert exc_info.value.kind == "header"


def test_anchor_by_id_malformed_header_raises(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        with pytest.raises(AnchorNotFoundError):
            doc.anchor_by_id("header:1")  # missing WHICH


def test_header_apply_style_inherited(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        with doc.edit("style header"):
            doc.sections[1].header().apply_style("Normal")
        # apply_style is inherited from Anchor — it writes the style onto the
        # header's range like any other anchor.
        assert doc.com.Sections(1).Headers(1).Range.Style == doc.com.Styles("Normal")
