"""Start-of-document prepend: doc.prepend / doc.prepend_paragraph + `start` anchor.

The mirror of `test_append.py`. The start has no terminal-mark complication, so
a prepended paragraph is written as ``"<text><break>"`` at offset 0.
"""

from __future__ import annotations

import pytest

import wordlive
from wordlive import StartAnchor
from wordlive.exceptions import StyleNotFoundError


def test_prepend_paragraph_writes_text_then_break_at_start(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        with doc.edit("prepend"):
            doc.prepend_paragraph("DRAFT")
    # New first paragraph is "<text><break>" written at offset 0.
    assert fake_word.ActiveDocument.Range(0, 0).Text == "DRAFT\r"


def test_prepend_paragraph_validates_style_before_writing(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        with pytest.raises(StyleNotFoundError):
            doc.prepend_paragraph("x", style="NoSuchStyle")
    assert fake_word.ActiveDocument.Range(0, 0).Text == ""


def test_prepend_paragraph_styles_full_span(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        with doc.edit("prepend-styled"):
            doc.prepend_paragraph("Title", style="Heading 1")
    # Text starts at 0 and spans "Title" (5 UTF-16 units).
    styled = [c.args for c in fake_word.ActiveDocument.Range.call_args_list if c.args == (0, 5)]
    assert styled, (
        f"expected styled Range(0, 5); got {fake_word.ActiveDocument.Range.call_args_list}"
    )


def test_prepend_inline_uses_content_insert_before(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        with doc.edit("prepend-inline"):
            doc.prepend("Note: ")
    fake_word.ActiveDocument.Content.InsertBefore.assert_called_once_with("Note: ")


def test_doc_start_is_a_start_anchor(fake_word):
    with wordlive.attach() as word:
        start = word.documents.active.start
    assert isinstance(start, StartAnchor)
    assert (start.anchor_id, start.kind, start.name) == ("start", "start", "start")


def test_anchor_by_id_start_resolves(fake_word):
    with wordlive.attach() as word:
        anchor = word.documents.active.anchor_by_id("start")
    assert isinstance(anchor, StartAnchor)
    assert anchor.anchor_id == "start"


def test_start_anchor_insert_paragraph_after_prepends(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        with doc.edit("prepend"):
            doc.start.insert_paragraph_after("Head.")
    assert fake_word.ActiveDocument.Range(0, 0).Text == "Head.\r"


def test_start_anchor_insert_after_prepends_inline(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        with doc.edit("prepend"):
            doc.start.insert_after("intro ")
    fake_word.ActiveDocument.Content.InsertBefore.assert_called_once_with("intro ")


def test_start_anchor_set_text_prepends(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        with doc.edit("prepend"):
            doc.start.set_text("prepended")
    fake_word.ActiveDocument.Content.InsertBefore.assert_called_once_with("prepended")


def test_start_anchor_text_is_empty(fake_word):
    with wordlive.attach() as word:
        assert word.documents.active.start.text == ""
