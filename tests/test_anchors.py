"""Anchor lookup + set_text + error translation."""

from __future__ import annotations

import pytest

import wordlive
from wordlive.exceptions import AnchorNotFoundError, ComError


def test_bookmark_missing_raises(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        with pytest.raises(AnchorNotFoundError) as exc_info:
            _ = doc.bookmarks["Nope"]
    assert exc_info.value.kind == "bookmark"
    assert exc_info.value.name == "Nope"


def test_bookmark_list(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        assert doc.bookmarks.list() == ["Address"]


def test_bookmark_set_text_replaces_and_readds(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        with doc.edit("Set address"):
            doc.bookmarks["Address"].set_text("123 Main St")

    # The bookmark was re-added covering the new content. fake_word's registry
    # ends up with one Address bookmark whose range matches the new text length.
    assert fake_word.ActiveDocument.Bookmarks.Exists("Address")


def test_content_control_missing_raises(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        with pytest.raises(AnchorNotFoundError):
            _ = doc.content_controls["NoSuchCC"]


def test_content_control_read(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        assert doc.content_controls["Signatory"].text == "Jane Doe"


def test_heading_missing_raises(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        with pytest.raises(AnchorNotFoundError):
            doc.heading("Nonexistent").text


def test_outline(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        outline = doc.outline()
    assert [item["text"] for item in outline] == ["Introduction", "Risks"]
    assert [item["level"] for item in outline] == [1, 2]
    assert all(item["anchor_id"].startswith("heading:") for item in outline)


def test_from_com_error_classifies_busy():
    from wordlive.exceptions import WordBusyError, from_com_error

    class _FakeComError(Exception):
        def __init__(self, hresult: int) -> None:
            super().__init__("Call was rejected by callee.")
            self.args = (
                hresult,
                "Call was rejected by callee.",
                (0, "Word", "Call was rejected by callee.", None, 0, hresult),
                None,
            )

    exc = _FakeComError(0x80010001)
    classified = from_com_error(exc)
    assert isinstance(classified, WordBusyError)
    assert classified.hresult == 0x80010001


def test_from_com_error_other_is_com_error():
    from wordlive.exceptions import from_com_error

    class _FakeComError(Exception):
        def __init__(self) -> None:
            super().__init__("Some other failure")
            self.args = (
                -2147352567,
                "Some other failure",
                (0, "Word", "Some other failure", None, 0, -2147352567),
                None,
            )

    classified = from_com_error(_FakeComError())
    assert isinstance(classified, ComError)
