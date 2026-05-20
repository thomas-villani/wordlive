"""Lists & numbering — apply / read / restart / level on anchors, plus doc.lists."""

from __future__ import annotations

import pytest

import wordlive
from wordlive.exceptions import AnchorNotFoundError


def test_apply_numbered_sets_list_type(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        target = doc.range(0, 12)
        with doc.edit("list"):
            target.apply_list("numbered")
        info = target.list_info()
        assert info["type"] == "numbered"
        assert info["number"] == 1


def test_apply_bulleted(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        target = doc.range(0, 12)
        target.apply_list("bulleted")
        assert target.list_info()["type"] == "bulleted"


def test_apply_outline(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        target = doc.range(0, 12)
        target.apply_list("outline")
        assert target.list_info()["type"] == "outline"


def test_apply_unknown_type_raises_value_error(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        with pytest.raises(ValueError):
            doc.range(0, 12).apply_list("squiggly")


def test_apply_continue_previous_passes_through(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        doc.range(0, 12).apply_list("numbered", continue_previous=True)
        assert doc.com.Range(0, 12).ListFormat._continue is True


def test_apply_default_does_not_continue(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        doc.range(0, 12).apply_list("numbered")
        assert doc.com.Range(0, 12).ListFormat._continue is False


def test_remove_list_clears_formatting(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        target = doc.range(0, 12)
        target.apply_list("numbered")
        with doc.edit("remove"):
            target.remove_list()
        assert target.list_info()["type"] == "none"


def test_list_info_on_plain_range_is_none(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        assert doc.range(0, 12).list_info()["type"] == "none"


def test_list_info_reads_seeded_list(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        info = doc.lists[1].list_info()
        assert info["type"] == "numbered"
        assert info["string"] == "1."


def test_restart_numbering_on_list(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        with doc.edit("restart"):
            doc.lists[1].restart_numbering()
        # Still a numbered list, now restarted at 1.
        info = doc.lists[1].list_info()
        assert info["type"] == "numbered"
        assert info["number"] == 1


def test_restart_numbering_on_non_list_raises(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        with pytest.raises(ValueError):
            doc.range(0, 12).restart_numbering()


def test_indent_and_outdent_change_level(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        target = doc.range(0, 12)
        target.apply_list("numbered")
        assert target.list_info()["level"] == 1
        with doc.edit("indent"):
            target.indent_list()
        assert target.list_info()["level"] == 2
        with doc.edit("outdent"):
            target.outdent_list()
        assert target.list_info()["level"] == 1


def test_doc_lists_len_and_list_shape(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        assert len(doc.lists) == 1
        rows = doc.lists.list()
    assert rows[0]["index"] == 1
    assert rows[0]["type"] == "numbered"
    assert rows[0]["count"] == 2
    assert rows[0]["anchor_id"] == "range:13-29"


def test_doc_lists_getitem_returns_range_anchor(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        anchor = doc.lists[1]
        assert anchor.kind == "range"
        assert anchor.anchor_id == "range:13-29"


def test_doc_lists_getitem_out_of_range_raises(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        with pytest.raises(AnchorNotFoundError) as exc_info:
            _ = doc.lists[5]
    assert exc_info.value.kind == "list"


def test_doc_lists_getitem_bool_rejected(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        with pytest.raises(TypeError):
            _ = doc.lists[True]


def test_doc_lists_iterate(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        ids = [a.anchor_id for a in doc.lists]
    assert ids == ["range:13-29"]
