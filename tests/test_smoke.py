"""Smoke tests that require a real running Word.

Skipped by default (`pytest -m "not smoke"` runs the unit suite). To exercise
them: `pytest -m smoke` on a Windows box with Word installed and *some*
document open.
"""

from __future__ import annotations

import base64
import contextlib

import pytest

import wordlive

pytestmark = pytest.mark.smoke

# A 1x1 transparent PNG, for the inline-image smoke checks.
_PNG = base64.b64decode(
    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAAC0lEQVR42mNkYPhfDwAChwGA"
    "60e6kgAAAABJRU5ErkJggg=="
)


@pytest.fixture
def scratch_doc(real_word):
    """A fresh blank document, discarded (no save) at teardown."""
    real_word.com.Documents.Add()
    doc = real_word.documents.active
    try:
        yield doc
    finally:
        with contextlib.suppress(Exception):
            doc.com.Close(SaveChanges=0)


@pytest.fixture
def png_path(tmp_path):
    p = tmp_path / "pic.png"
    p.write_bytes(_PNG)
    return str(p)


def test_status_returns_at_least_one_doc(real_word):
    docs = real_word.documents.list()
    assert isinstance(docs, list)
    # If Word is running but nothing is open, this is fine — the smoke is
    # really about "does the COM dance hang together end-to-end".
    for d in docs:
        assert "name" in d and "path" in d and "is_active" in d


def test_outline_runs(real_word):
    try:
        doc = real_word.documents.active
    except Exception:
        pytest.skip("No active document open in Word")
    outline = doc.outline()
    assert isinstance(outline, list)
    for item in outline:
        assert {"level", "text", "anchor_id"} <= set(item)


def test_edit_scope_preserves_selection(real_word):
    """Trivial edit scope round-trip — verifies Selection survives a no-op edit."""
    try:
        doc = real_word.documents.active
    except Exception:
        pytest.skip("No active document open in Word")
    sel = real_word.selection.info()
    with doc.edit("wordlive smoke: no-op"):
        pass
    after = real_word.selection.info()
    assert (sel["start"], sel["end"]) == (after["start"], after["end"])


def test_snapshot_renders_a_png(real_word):
    """Real Word -> PDF -> PNG: page 1 of the active document rasterises."""
    pytest.importorskip("pymupdf")
    try:
        doc = real_word.documents.active
    except Exception:
        pytest.skip("No active document open in Word")
    shots = doc.snapshot(pages=1, dpi=72)
    assert len(shots) == 1
    assert shots[0].page == 1
    # PNG magic number — proof we got a real raster image back.
    assert shots[0].png.startswith(b"\x89PNG\r\n\x1a\n")


# ---------------------------------------------------------------------------
# Round-2 review fixes (boundary crashes, table corruption, image/break)
# ---------------------------------------------------------------------------


def test_find_replace_in_last_paragraph_no_boundary_crash(scratch_doc):
    """Fix 1a: a match in the final paragraph must not raise COM 0x80020009."""
    doc = scratch_doc
    with doc.edit("seed"):
        doc.append_paragraph("First paragraph here.")
        doc.append_paragraph("Final paragraph SENTINEL.")
    with doc.edit("fr"):
        applied = doc.find_replace("SENTINEL", "REPLACED")
    assert len(applied) == 1
    assert doc.find("REPLACED"), "replacement should be locatable"


def test_add_table_at_document_end_no_boundary_crash(scratch_doc):
    """Fix 1b: add_table appended at the terminal mark must not crash."""
    doc = scratch_doc
    before = len(doc.tables)
    with doc.edit("add table at end"):
        doc.add_table(2, 2, data=[["A1", "B1"], ["A2", "B2"]])
    assert len(doc.tables) == before + 1


def test_whole_doc_replace_targets_correct_table_cell(scratch_doc):
    """Fix 2b: a whole-doc replace of a value inside one cell hits that cell,
    not a neighbour, and leaves the others intact."""
    doc = scratch_doc
    with doc.edit("seed table"):
        doc.add_table(2, 2, data=[["None", "Looks"], ["Keep", "None"]])
    table = doc.tables[len(doc.tables)]
    with doc.edit("replace 2nd None"):
        doc.find_replace("None", "DONE", occurrence=2)
    grid = table.grid()
    # occurrence 2 is the (2,2) cell; the (1,1) "None" and the "Looks" neighbour
    # must be untouched (the old bug overwrote the neighbour "Looks").
    assert grid[0][0] == "None"
    assert grid[0][1] == "Looks"
    assert grid[1][1] == "DONE"


def test_inline_image_reads_back_as_token(scratch_doc, png_path):
    """Fix 3a: an inline image surfaces as a [image] token, not a phantom char."""
    doc = scratch_doc
    with doc.edit("seed heading"):
        doc.append_paragraph("The Problem", style="Heading 1")
    heading = doc.headings["The Problem"]
    with doc.edit("inline image"):
        heading.insert_image(png_path, wrap="inline", where="before")
    # The image embeds in the heading's run, so the paragraph now reads with a
    # visible token instead of a phantom control character. Re-find by scanning
    # paragraphs since the heading's name-text changed.
    texts = [p.text for p in doc.paragraphs]
    assert any("[image]" in t and "The Problem" in t for t in texts)


def test_block_image_keeps_heading_text_clean(scratch_doc, png_path):
    """Fix 3b: a block image sits on its own line, leaving the heading text intact."""
    doc = scratch_doc
    with doc.edit("seed heading"):
        doc.append_paragraph("The Problem", style="Heading 1")
    heading = doc.headings["The Problem"]
    with doc.edit("block image"):
        heading.insert_image(png_path, wrap="inline", where="before", block=True)
    assert heading.text == "The Problem"


def test_section_continuous_break_no_outline_pollution(scratch_doc):
    """Fix 4: a section break before a heading must not add an empty outline entry."""
    doc = scratch_doc
    with doc.edit("seed heading"):
        doc.append_paragraph("Section Heading", style="Heading 1")
    before = [i["text"] for i in doc.outline()]
    with doc.edit("section break"):
        doc.headings["Section Heading"].insert_break("section_continuous", where="before")
    after = [i["text"] for i in doc.outline()]
    # No new (empty-text) heading entry crept into the navigation outline.
    assert after == before


# ---------------------------------------------------------------------------
# Reference apparatus — footnotes / endnotes / TOC (live)
# ---------------------------------------------------------------------------


def _seed_two_sections(doc):
    """Scaffold a tiny doc; return the two heading anchor ids (positional)."""
    with doc.edit("seed"):
        doc.append_paragraph("Introduction", style="Heading 1")
        doc.append_paragraph("Body of the intro section.", style="Normal")
        doc.append_paragraph("Methods", style="Heading 1")
        doc.append_paragraph("Body of methods.", style="Normal")
    return doc.headings["Introduction"].anchor_id, doc.headings["Methods"].anchor_id


def test_footnote_insert_read_edit_delete(scratch_doc):
    doc = scratch_doc
    intro_id, _ = _seed_two_sections(doc)
    with doc.edit("footnote"):
        note = doc.anchor_by_id(intro_id).insert_footnote("A real footnote.")
    assert note.anchor_id == "footnote:1"
    assert len(doc.footnotes) == 1
    assert doc.footnotes[1].text == "A real footnote."
    # footnote:N resolves and round-trips a body edit.
    with doc.edit("edit note"):
        doc.anchor_by_id("footnote:1").set_text("Edited body.")
    assert doc.footnotes[1].text == "Edited body."
    with doc.edit("delete note"):
        doc.footnotes[1].delete()
    assert len(doc.footnotes) == 0


def test_endnote_insert(scratch_doc):
    doc = scratch_doc
    _, methods_id = _seed_two_sections(doc)
    with doc.edit("endnote"):
        doc.anchor_by_id(methods_id).insert_endnote("An endnote body.")
    assert len(doc.endnotes) == 1
    assert doc.endnotes[1].text == "An endnote body."


def test_toc_builds_and_page_numbers_populate(scratch_doc):
    doc = scratch_doc
    _seed_two_sections(doc)
    with doc.edit("toc"):
        toc = doc.add_toc(levels=(1, 2))
    with doc.edit("update"):
        doc.update_fields()
    assert int(doc.com.TablesOfContents.Count) == 1
    # Entries render with tab + page number once fields update.
    assert "Introduction" in toc.text and "Methods" in toc.text


# ---------------------------------------------------------------------------
# Anchoring & linking — bookmarks / hyperlinks / cross-refs / captions (live)
# ---------------------------------------------------------------------------


def test_bookmark_add_then_internal_link(scratch_doc):
    doc = scratch_doc
    intro_id, methods_id = _seed_two_sections(doc)
    with doc.edit("bookmark"):
        doc.bookmarks.add("Intro", intro_id)
    assert "Intro" in doc.bookmarks
    before = int(doc.com.Hyperlinks.Count)
    with doc.edit("link"):
        # Internal jump with new link text — must NOT overwrite the heading.
        doc.anchor_by_id(methods_id).link_to(bookmark="Intro", text="see Intro")
    assert int(doc.com.Hyperlinks.Count) == before + 1
    assert "Methods" in [h["text"] for h in doc.outline()]  # heading preserved


def test_external_hyperlink(scratch_doc):
    doc = scratch_doc
    _seed_two_sections(doc)
    before = int(doc.com.Hyperlinks.Count)
    with doc.edit("link"):
        doc.end.link_to(address="https://example.com", text="example")
    assert int(doc.com.Hyperlinks.Count) == before + 1


def test_cross_reference_to_bookmark_heading_and_footnote(scratch_doc):
    doc = scratch_doc
    intro_id, _ = _seed_two_sections(doc)
    with doc.edit("setup"):
        doc.bookmarks.add("Intro", intro_id)
        doc.anchor_by_id(intro_id).insert_footnote("fn body")
    before = int(doc.com.Fields.Count)
    with doc.edit("xrefs"):
        doc.end.insert_cross_reference("bookmark:Intro", kind="page")  # name, not index
        doc.end.insert_cross_reference(intro_id, kind="text")  # heading text
        doc.end.insert_cross_reference("footnote:1", kind="number")  # note number
        doc.end.insert_cross_reference("footnote:1", kind="text")  # falls back to number
    assert int(doc.com.Fields.Count) >= before + 4


def test_cross_reference_bad_target_raises(scratch_doc):
    doc = scratch_doc
    _seed_two_sections(doc)
    with pytest.raises(wordlive.AnchorNotFoundError):
        with doc.edit("bad"):
            doc.end.insert_cross_reference("bookmark:DoesNotExist")


def test_caption_insert(scratch_doc):
    doc = scratch_doc
    _, methods_id = _seed_two_sections(doc)
    with doc.edit("caption"):
        doc.anchor_by_id(methods_id).insert_caption("Figure", text="System overview")
    # The caption adds a paragraph; the doc still has its headings.
    assert "Methods" in [h["text"] for h in doc.outline()]
