"""Style lookup, apply_style, format_paragraph, StyleNotFoundError."""

from __future__ import annotations

import pytest

import wordlive
from wordlive.constants import WdParagraphAlignment
from wordlive.exceptions import AnchorNotFoundError, StyleNotFoundError

# ---------------------------------------------------------------------------
# StyleCollection
# ---------------------------------------------------------------------------


def test_styles_list_returns_known_styles(fake_word):
    with wordlive.attach() as word:
        rows = word.documents.active.styles.list()
    names = [r["name"] for r in rows]
    assert "Normal" in names
    assert "Body Text" in names
    assert "Heading 1" in names
    # type is a string ("paragraph" / "character" / ...)
    assert all(isinstance(r["type"], str) for r in rows)
    assert all(isinstance(r["builtin"], bool) for r in rows)
    assert all(isinstance(r["in_use"], bool) for r in rows)


def test_styles_contains(fake_word):
    with wordlive.attach() as word:
        styles = word.documents.active.styles
        assert "Heading 1" in styles
        assert "NoSuchStyle" not in styles
        assert 123 not in styles  # non-string fail-safe


def test_styles_getitem_missing_raises_style_not_found(fake_word):
    with wordlive.attach() as word:
        with pytest.raises(StyleNotFoundError) as exc_info:
            _ = word.documents.active.styles["NoSuchStyle"]
    assert exc_info.value.kind == "style"
    assert exc_info.value.name == "NoSuchStyle"
    # Subclass of AnchorNotFoundError so existing handlers still catch it.
    assert isinstance(exc_info.value, AnchorNotFoundError)


def test_styles_iter_yields_style_wrappers(fake_word):
    with wordlive.attach() as word:
        styles = list(word.documents.active.styles)
    names = [s.name for s in styles]
    assert "Body Text" in names


def test_style_properties(fake_word):
    with wordlive.attach() as word:
        s = word.documents.active.styles["Body Text"]
    assert s.name == "Body Text"
    assert s.type == "paragraph"
    assert s.builtin is True


# ---------------------------------------------------------------------------
# apply_style
# ---------------------------------------------------------------------------


def test_apply_style_sets_range_style(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        with doc.edit("apply heading 2"):
            doc.bookmarks["Address"].apply_style("Heading 2")
    # Fake range stored .Style; check it was written through to the right style
    bm_range = fake_word.ActiveDocument.Bookmarks("Address").Range
    assert getattr(bm_range, "Style", None) is not None
    assert bm_range.Style.NameLocal == "Heading 2"


def test_apply_style_missing_raises_before_mutation(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        with pytest.raises(StyleNotFoundError):
            doc.bookmarks["Address"].apply_style("NoSuchStyle")


def test_insert_paragraph_with_bad_style_raises_typed(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        with pytest.raises(StyleNotFoundError):
            doc.heading("Introduction").insert_paragraph_after("x", style="NoSuchStyle")


# ---------------------------------------------------------------------------
# format_paragraph
# ---------------------------------------------------------------------------


def test_format_paragraph_alignment_string(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        with doc.edit("center"):
            doc.bookmarks["Address"].format_paragraph(alignment="center")
    pf = fake_word.ActiveDocument.Bookmarks("Address").Range.ParagraphFormat
    assert pf.Alignment == int(WdParagraphAlignment.CENTER)


def test_format_paragraph_alignment_enum(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        doc.bookmarks["Address"].format_paragraph(alignment=WdParagraphAlignment.RIGHT)
    pf = fake_word.ActiveDocument.Bookmarks("Address").Range.ParagraphFormat
    assert pf.Alignment == int(WdParagraphAlignment.RIGHT)


def test_format_paragraph_alignment_invalid_string_raises(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        with pytest.raises(ValueError):
            doc.bookmarks["Address"].format_paragraph(alignment="diagonal")


def test_format_paragraph_indent_and_spacing(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        doc.bookmarks["Address"].format_paragraph(
            left_indent=36.0,
            right_indent=18.0,
            first_line_indent=12.0,
            space_before=6.0,
            space_after=4.5,
        )
    pf = fake_word.ActiveDocument.Bookmarks("Address").Range.ParagraphFormat
    assert pf.LeftIndent == 36.0
    assert pf.RightIndent == 18.0
    assert pf.FirstLineIndent == 12.0
    assert pf.SpaceBefore == 6.0
    assert pf.SpaceAfter == 4.5


def test_format_paragraph_no_kwargs_is_noop(fake_word):
    """Calling format_paragraph with no kwargs must not raise — just do nothing."""
    with wordlive.attach() as word:
        doc = word.documents.active
        doc.bookmarks["Address"].format_paragraph()


# ---------------------------------------------------------------------------
# apply_style / format_paragraph cross-kind coverage (T-1)
#
# These live on the base `Anchor` class, so they should work identically on
# every anchor kind. Bookmark coverage above; mirror it for content controls
# and headings so a regression in `_range()` on either kind gets caught.
# ---------------------------------------------------------------------------


def test_apply_style_on_content_control_writes_through(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        with doc.edit("cc style"):
            doc.content_controls["Signatory"].apply_style("Heading 2")
    # ContentControl.Range was assigned a Style; pull the CC from the fake and
    # verify the Style applied was the Heading 2 mock.
    cc_mock = next(c for c in fake_word.ActiveDocument.ContentControls if c.Title == "Signatory")
    assert cc_mock.Range.Style.NameLocal == "Heading 2"


def test_apply_style_on_heading_writes_through(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        with doc.edit("heading style"):
            doc.heading("Risks").apply_style("Heading 1")
    # Heading anchor exposes the paragraph's Range; check Style landed there.
    para = next(
        p for p in fake_word.ActiveDocument.Paragraphs if p.Range.Text.rstrip("\r\n\x07") == "Risks"
    )
    assert para.Range.Style.NameLocal == "Heading 1"


def test_format_paragraph_on_content_control(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        doc.content_controls["Signatory"].format_paragraph(alignment="right", left_indent=24.0)
    cc_mock = next(c for c in fake_word.ActiveDocument.ContentControls if c.Title == "Signatory")
    pf = cc_mock.Range.ParagraphFormat
    assert pf.Alignment == int(WdParagraphAlignment.RIGHT)
    assert pf.LeftIndent == 24.0


def test_format_paragraph_on_heading(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        doc.heading("Risks").format_paragraph(space_before=12.0, alignment="center")
    para = next(
        p for p in fake_word.ActiveDocument.Paragraphs if p.Range.Text.rstrip("\r\n\x07") == "Risks"
    )
    pf = para.Range.ParagraphFormat
    assert pf.SpaceBefore == 12.0
    assert pf.Alignment == int(WdParagraphAlignment.CENTER)
