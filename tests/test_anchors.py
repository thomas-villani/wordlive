"""Anchor lookup + set_text + error translation."""

from __future__ import annotations

from typing import Any

import pytest

import wordlive
from wordlive.exceptions import AnchorNotFoundError, ComError


def test_anchor_abstract_base_cannot_be_instantiated():
    """Regression for D-2: Anchor must refuse direct instantiation (it's ABC).
    Subclasses missing `_range` or `set_text` now fail at construction time
    with a clear TypeError, not silently until the first call.
    """
    from wordlive._anchors import Anchor

    with pytest.raises(TypeError, match=r"abstract"):
        Anchor(None, "x")  # type: ignore[abstract,arg-type]


def test_anchor_subclass_missing_method_fails_at_construction():
    """A would-be Anchor that forgets to override `set_text` must fail when
    instantiated, not at call site."""
    from wordlive._anchors import Anchor

    class IncompleteAnchor(Anchor):
        # _range overridden, but set_text intentionally omitted
        def _range(self) -> Any:
            return None

    with pytest.raises(TypeError, match=r"set_text"):
        IncompleteAnchor(None, "x")  # type: ignore[arg-type]


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


def test_bookmark_list_hides_word_internal_bookmarks(monkeypatch):
    """Regression for B-5: Word auto-creates `_Toc...` and `_Ref...` bookmarks
    for table-of-contents entries and cross-references. They drown the user's
    own bookmarks in noise; default `list()` should filter them out.
    """
    from tests.conftest import _make_application, _make_document

    doc_com = _make_document(
        bookmarks={
            "Address": (10, 20),
            "_Toc1234567": (0, 5),
            "_Ref9876": (30, 35),
        },
    )
    app = _make_application([doc_com])
    from wordlive import _com as _com_module

    monkeypatch.setattr(_com_module, "get_active_word", lambda: app)

    with wordlive.attach() as word:
        doc = word.documents.active
        # Default: hidden bookmarks filtered out.
        assert doc.bookmarks.list() == ["Address"]
        # Opt-in escape hatch surfaces them all.
        assert set(doc.bookmarks.list(include_hidden=True)) == {
            "Address",
            "_Toc1234567",
            "_Ref9876",
        }
        # Direct lookup by name still works (so agents that DO need them aren't blocked).
        assert "_Toc1234567" in doc.bookmarks
        assert doc.bookmarks["_Toc1234567"].name == "_Toc1234567"


def test_bookmark_set_text_replaces_and_readds(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        with doc.edit("Set address"):
            doc.bookmarks["Address"].set_text("123 Main St")

    # The bookmark was re-added covering the new content. fake_word's registry
    # ends up with one Address bookmark whose range matches the new text length.
    assert fake_word.ActiveDocument.Bookmarks.Exists("Address")


def test_bookmark_set_text_with_surrogate_pairs_uses_utf16_length(fake_word):
    """Regression for B-2: Word counts characters in UTF-16 code units, so a
    single emoji is *2* characters from Word's POV. Computing the new bookmark
    end via Python `len()` under-counted and re-added the bookmark covering a
    range shorter than the actual inserted text.
    """
    with wordlive.attach() as word:
        doc = word.documents.active
        with doc.edit("emoji"):
            # 🎉 is one Python char but two UTF-16 code units; " done" is 5.
            doc.bookmarks["Address"].set_text("🎉 done")

    # The Address bookmark starts at offset 13 in the fixture. The new end must
    # be 13 + 7 (UTF-16 units), not 13 + 6 (Python code points).
    registered = fake_word.ActiveDocument.Bookmarks._items["Address"]
    assert registered == (13, 13 + 7), f"got {registered!r}, expected (13, 20)"


def test_insert_paragraph_after_with_emoji_styles_full_span(fake_word):
    """Regression for B-2: style application must cover the whole inserted
    text in UTF-16 units. Python `len("🎉 hi")` is 4, but Word sees 5 chars,
    so the old `end + len(text)` math would leave the trailing character
    un-styled.
    """
    with wordlive.attach() as word:
        doc = word.documents.active
        with doc.edit("emoji-insert"):
            doc.heading("Introduction").insert_paragraph_after("🎉 hi", style="Body Text")

    # End of "Introduction" paragraph is at offset 13. The styled range must
    # span 13 → 13+5 (UTF-16 units), covering all of "🎉 hi".
    styled_ranges = [
        call.args for call in fake_word.ActiveDocument.Range.call_args_list if call.args == (13, 18)
    ]
    assert styled_ranges, (
        f"expected styled Range(13, 18); got {fake_word.ActiveDocument.Range.call_args_list}"
    )


def test_insert_paragraph_after_terminal_paragraph_splits_before_final_mark(fake_word):
    """Regression: inserting after the document's *last* paragraph must not try
    to write at `Range(end, end)` — that position is past the final paragraph
    mark and Word rejects it with a "value out of range" COM error. The new
    paragraph is split in just before the final mark instead (at Content.End-1),
    so appending to a freshly-created document just works.
    """
    # "Risks" is the fixture's terminal paragraph: range [29, 35), and the
    # document content ends at 35, so its end coincides with Content.End.
    with wordlive.attach() as word:
        doc = word.documents.active
        with doc.edit("append"):
            doc.heading("Risks").insert_paragraph_after("New tail paragraph.")

    # Inserted at Content.End - 1 == 34 as "<break><text>", not at 35.
    assert fake_word.ActiveDocument.Range(34, 34).Text == "\rNew tail paragraph."


def test_insert_paragraph_after_terminal_paragraph_styles_inserted_text(fake_word):
    """The style must land on the text *after* the leading break, not on the
    anchor's paragraph. For a terminal insert the text starts at Content.End."""
    with wordlive.attach() as word:
        doc = word.documents.active
        with doc.edit("append-styled"):
            doc.heading("Risks").insert_paragraph_after("Body", style="Body Text")

    # Content.End is 35; text starts at 35 and spans "Body" (4 UTF-16 units).
    styled = [c.args for c in fake_word.ActiveDocument.Range.call_args_list if c.args == (35, 39)]
    assert styled, (
        f"expected styled Range(35, 39); got {fake_word.ActiveDocument.Range.call_args_list}"
    )


def test_content_control_missing_raises(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        with pytest.raises(AnchorNotFoundError):
            _ = doc.content_controls["NoSuchCC"]


def test_content_control_empty_name_does_not_match_untitled_cc(fake_word, monkeypatch):
    """Regression for B-3: a CC with empty Title and Tag must not be returned
    for `doc.content_controls[""]`. Word templates often include untitled
    rich-text or repeating-section CCs and matching them on empty input would
    silently target the wrong control.
    """
    from tests.conftest import _make_application, _make_document

    doc_com = _make_document(
        content_controls=[
            # An untitled, untagged CC — both fields empty.
            {"title": "", "tag": "", "start": 0, "end": 5, "text": "untitled"},
            {"title": "Real", "tag": "real", "start": 10, "end": 14, "text": "real"},
        ],
    )
    app = _make_application([doc_com])
    from wordlive import _com as _com_module

    monkeypatch.setattr(_com_module, "get_active_word", lambda: app)

    with wordlive.attach() as word:
        doc = word.documents.active
        with pytest.raises(AnchorNotFoundError):
            _ = doc.content_controls[""]
        # And the named one still resolves.
        assert doc.content_controls["Real"].text == "real"


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


@pytest.mark.parametrize(
    "hresult",
    [
        0x80010001,  # RPC_E_CALL_REJECTED
        0x8001010A,  # RPC_E_SERVERCALL_RETRYLATER
        0x80010005,  # RPC_E_SERVERCALL_REJECTED
        -2147418111,  # signed RPC_E_CALL_REJECTED
        -2147417846,  # signed RPC_E_SERVERCALL_RETRYLATER
    ],
)
def test_from_com_error_classifies_busy(hresult):
    """T-4: every entry in `_BUSY_HRESULTS` must classify to WordBusyError."""
    from wordlive.exceptions import WordBusyError, from_com_error

    class _FakeComError(Exception):
        def __init__(self, h: int) -> None:
            super().__init__("Call was rejected by callee.")
            self.args = (
                h,
                "Call was rejected by callee.",
                (0, "Word", "Call was rejected by callee.", None, 0, h),
                None,
            )

    classified = from_com_error(_FakeComError(hresult))
    assert isinstance(classified, WordBusyError)
    assert classified.hresult == hresult


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
        assert anchor.kind == "content_control"


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


def test_anchor_by_id_unknown_scheme_hints_valid_types(fake_word):
    # A malformed scheme should say it's an unknown *type*, not look like a
    # valid-scheme-but-missing-target, so the message lists what's accepted.
    with wordlive.attach() as word:
        doc = word.documents.active
        with pytest.raises(AnchorNotFoundError) as exc_info:
            doc.anchor_by_id("banana:7")
    msg = str(exc_info.value)
    assert "unknown anchor type" in msg
    assert "'banana'" in msg
    assert "heading/para/bookmark" in msg


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


# ---------------------------------------------------------------------------
# HeadingCollection (D-1)
# ---------------------------------------------------------------------------


def test_headings_iter_yields_each_heading(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        names = [h.text for h in doc.headings]
    assert names == ["Introduction", "Risks"]


def test_headings_getitem_by_name_and_index(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        by_name = doc.headings["Risks"]
        by_index = doc.headings[3]  # paragraph 3 is the Risks heading in fixture
    assert by_name.text == "Risks"
    assert by_index.text == "Risks"
    assert by_index.anchor_id == "heading:3"


def test_headings_contains_by_name_and_index(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        h = doc.headings
        assert "Introduction" in h
        assert "Nope" not in h
        assert 1 in h  # paragraph 1 is a heading
        assert 2 not in h  # paragraph 2 is body (OutlineLevel=10)
        assert 99 not in h  # past end of doc
        assert object() not in h
        assert True not in h  # bool quirk-guard


def test_headings_getitem_missing_raises(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        with pytest.raises(AnchorNotFoundError):
            _ = doc.headings["Nope"]


def test_headings_list_same_shape_as_outline(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        assert doc.headings.list() == doc.outline()


def test_heading_method_delegates_to_collection(fake_word):
    """`doc.heading(name)` is sugar for `doc.headings[name]` — should resolve identically."""
    with wordlive.attach() as word:
        doc = word.documents.active
        assert doc.heading("Risks").text == doc.headings["Risks"].text


def test_heading_section_range_when_heading_is_last_paragraph(monkeypatch):
    """Regression for T-3: when the target heading is the last paragraph in
    the document, `section_range` should return a zero-width range at the end
    rather than walking past the array bounds.
    """
    from tests.conftest import _make_application, _make_document

    doc_com = _make_document(
        paragraphs=[
            {"level": 10, "text": "Body", "start": 0, "end": 5},
            {"level": 1, "text": "Last", "start": 5, "end": 10},
        ],
        content="Body\rLast\r",
    )
    app = _make_application([doc_com])
    from wordlive import _com as _com_module

    monkeypatch.setattr(_com_module, "get_active_word", lambda: app)

    with wordlive.attach() as word:
        doc = word.documents.active
        rng = doc.heading("Last").section_range()
        # End of Last paragraph is 10; no following paragraphs, so the section
        # runs to the end of that same paragraph — effectively zero-width.
        assert int(rng.Start) == 10
        assert int(rng.End) == 10


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
