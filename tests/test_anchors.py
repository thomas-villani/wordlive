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


# ---------------------------------------------------------------------------
# anchor_by_id (v0.1)
# ---------------------------------------------------------------------------


def test_anchor_by_id_heading_index_resolves(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        anchor = doc.anchor_by_id("heading:1")
        assert anchor.kind == "heading"
        assert anchor.text == "Introduction"


def test_anchor_by_id_heading_skips_body_paragraphs(fake_word):
    """Paragraph 2 in the fixture is body text (OutlineLevel=10), so heading:2 is invalid."""
    with wordlive.attach() as word:
        doc = word.documents.active
        with pytest.raises(AnchorNotFoundError):
            doc.anchor_by_id("heading:2").text


def test_anchor_by_id_bookmark(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        anchor = doc.anchor_by_id("bookmark:Address")
        assert anchor.kind == "bookmark"
        assert anchor.name == "Address"


def test_anchor_by_id_cc(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        anchor = doc.anchor_by_id("cc:Signatory")
        assert anchor.kind == "content control"


def test_anchor_by_id_bad_scheme(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        with pytest.raises(AnchorNotFoundError):
            doc.anchor_by_id("table:1")


def test_anchor_by_id_missing_colon(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        with pytest.raises(AnchorNotFoundError):
            doc.anchor_by_id("no-colon-here")


def test_anchor_by_id_non_integer_heading(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        with pytest.raises(AnchorNotFoundError):
            doc.anchor_by_id("heading:notanint")


def test_anchor_by_id_missing_bookmark(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        with pytest.raises(AnchorNotFoundError):
            doc.anchor_by_id("bookmark:Nope")


# ---------------------------------------------------------------------------
# anchor_id properties
# ---------------------------------------------------------------------------


def test_bookmark_anchor_id(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        assert doc.bookmarks["Address"].anchor_id == "bookmark:Address"


def test_content_control_anchor_id(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        assert doc.content_controls["Signatory"].anchor_id == "cc:Signatory"


def test_heading_anchor_id_resolves_by_text(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        # Introduction is paragraph 1 in the fixture.
        assert doc.heading("Introduction").anchor_id == "heading:1"


def test_indexed_heading_anchor_id(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        anchor = doc.anchor_by_id("heading:1")
        assert anchor.anchor_id == "heading:1"


# ---------------------------------------------------------------------------
# Heading.level + Heading.section_range / section_text
# ---------------------------------------------------------------------------


def test_heading_level(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        assert doc.heading("Introduction").level == 1
        assert doc.heading("Risks").level == 2


def test_heading_section_text_runs_to_next_same_or_higher(fake_word):
    """Introduction (level 1) section runs until the next level-<=1 heading or end of doc.
    Since Risks is level 2 (not <=1), the Introduction section runs to the end.
    """
    with wordlive.attach() as word:
        doc = word.documents.active
        rng = doc.heading("Introduction").section_range()
        # End of Introduction paragraph is 13; end of last paragraph is 35.
        assert int(rng.Start) == 13
        assert int(rng.End) == 35


def test_heading_section_text_stops_at_same_level(fake_word, monkeypatch):
    """When a sibling heading exists, the section stops before it."""
    # Add a sibling level-1 heading after Risks so Introduction's section
    # boundary is exercised.
    from tests.conftest import _make_application, _make_document

    doc_com = _make_document(
        paragraphs=[
            {"level": 1, "text": "Introduction", "start": 0, "end": 13},
            {"level": 10, "text": "Body 1", "start": 13, "end": 20},
            {"level": 1, "text": "Conclusion", "start": 20, "end": 31},
            {"level": 10, "text": "Body 2", "start": 31, "end": 38},
        ],
        content="Introduction\rBody 1\rConclusion\rBody 2\r",
    )
    app = _make_application([doc_com])
    from wordlive import _com as _com_module
    monkeypatch.setattr(_com_module, "get_active_word", lambda: app)

    with wordlive.attach() as word:
        doc = word.documents.active
        rng = doc.heading("Introduction").section_range()
        # Body 1 starts at 13, Conclusion starts at 20 — section is [13, 20).
        assert int(rng.Start) == 13
        assert int(rng.End) == 20
