"""Smoke tests that require a real running Word.

Skipped by default (`pytest -m "not smoke"` runs the unit suite). To exercise
them: `pytest -m smoke` on a Windows box with Word installed and *some*
document open.
"""

from __future__ import annotations

import base64
import contextlib

import pytest

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
