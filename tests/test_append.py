"""End-of-document append: doc.append / doc.append_paragraph + the `end` anchor.

The fixture document's content is ``"Introduction\rBody text here.\rRisks\r"``
(35 UTF-16 units), so ``Content.End`` is 35 and an append writes its leading
break at offset 34 — just before the final paragraph mark.
"""

from __future__ import annotations

import pytest

import wordlive
from wordlive import EndAnchor
from wordlive.exceptions import StyleNotFoundError


def test_append_paragraph_writes_break_then_text_before_final_mark(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        with doc.edit("append"):
            doc.append_paragraph("Closing note.")
    # New paragraph is "<break><text>" written at Content.End - 1 == 34, never
    # at 35 (past the final mark — a "value out of range" COM error).
    assert fake_word.ActiveDocument.Range(34, 34).Text == "\rClosing note."


def test_append_paragraph_validates_style_before_writing(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        with pytest.raises(StyleNotFoundError):
            doc.append_paragraph("x", style="NoSuchStyle")
    # The bad style is rejected before any Range.Text assignment.
    assert fake_word.ActiveDocument.Range(34, 34).Text == ""


def test_append_paragraph_styles_full_span(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        with doc.edit("append-styled"):
            doc.append_paragraph("Body", style="Body Text")
    # Text starts after the leading break at 35 and spans "Body" (4 UTF-16 units).
    styled = [c.args for c in fake_word.ActiveDocument.Range.call_args_list if c.args == (35, 39)]
    assert styled, (
        f"expected styled Range(35, 39); got {fake_word.ActiveDocument.Range.call_args_list}"
    )


def test_append_inline_uses_content_insert_after(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        with doc.edit("append-inline"):
            doc.append(" (verified)")
    fake_word.ActiveDocument.Content.InsertAfter.assert_called_once_with(" (verified)")


def test_doc_end_is_an_end_anchor(fake_word):
    with wordlive.attach() as word:
        end = word.documents.active.end
    assert isinstance(end, EndAnchor)
    assert (end.anchor_id, end.kind, end.name) == ("end", "end", "end")


def test_anchor_by_id_end_resolves(fake_word):
    with wordlive.attach() as word:
        anchor = word.documents.active.anchor_by_id("end")
    assert isinstance(anchor, EndAnchor)
    assert anchor.anchor_id == "end"


def test_end_anchor_insert_paragraph_after_appends(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        with doc.edit("append"):
            doc.end.insert_paragraph_after("Tail.")
    assert fake_word.ActiveDocument.Range(34, 34).Text == "\rTail."


def test_end_anchor_insert_after_appends_inline(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        with doc.edit("append"):
            doc.end.insert_after(" more")
    fake_word.ActiveDocument.Content.InsertAfter.assert_called_once_with(" more")


def test_end_anchor_set_text_appends(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        with doc.edit("append"):
            doc.end.set_text("appended")
    fake_word.ActiveDocument.Content.InsertAfter.assert_called_once_with("appended")


def test_end_anchor_text_is_empty(fake_word):
    with wordlive.attach() as word:
        # The end position holds no content — a zero-width range reads "".
        assert word.documents.active.end.text == ""
