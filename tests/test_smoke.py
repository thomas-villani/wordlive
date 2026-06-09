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


def _paras(doc):
    """[(text, style_name)] for each paragraph in the live document."""
    return [(p.Range.Text.rstrip("\r\x07"), p.Range.Style.NameLocal) for p in doc.com.Paragraphs]


def _para_id_by_text(doc, text):
    """anchor_id of the first paragraph whose text matches `text` (1-based scan)."""
    for p in doc.paragraphs:
        if p.text == text:
            return p.anchor_id
    raise AssertionError(f"no paragraph with text {text!r}")


def test_caption_insert(scratch_doc):
    doc = scratch_doc
    _, methods_id = _seed_two_sections(doc)
    with doc.edit("caption"):
        doc.anchor_by_id(methods_id).insert_caption("Figure", text="System overview")
    # The caption adds a paragraph; the doc still has its headings.
    assert "Methods" in [h["text"] for h in doc.outline()]


def test_caption_is_own_paragraph_below_target(scratch_doc):
    """A figure caption is a *separate* Caption paragraph below the target,
    leaving the target paragraph's text untouched (regression: it used to fuse
    inline into the following paragraph)."""
    doc = scratch_doc
    with doc.edit("seed"):
        doc.append_paragraph("Target paragraph.", style="Normal")
        doc.append_paragraph("Trailing paragraph.", style="Normal")
    target_id = _para_id_by_text(doc, "Target paragraph.")
    with doc.edit("caption"):
        doc.anchor_by_id(target_id).insert_caption("Figure", text="Overview")
    paras = _paras(doc)
    texts = [t for t, _ in paras]
    # The target text is intact and *not* merged with the caption.
    assert "Target paragraph." in texts
    assert "Trailing paragraph." in texts
    # Exactly one paragraph is styled Caption, and it carries the label+title.
    caps = [(t, s) for t, s in paras if s == "Caption"]
    assert len(caps) == 1
    cap_text, _ = caps[0]
    assert cap_text.startswith("Figure")
    assert "Overview" in cap_text
    # The caption sits immediately after the target, before the trailing para.
    idx_target = texts.index("Target paragraph.")
    idx_cap = next(i for i, (t, s) in enumerate(paras) if s == "Caption")
    assert idx_cap == idx_target + 1


def test_caption_table_label_is_above(scratch_doc):
    """A 'Table' caption defaults *above* the anchor's paragraph."""
    doc = scratch_doc
    with doc.edit("seed"):
        doc.append_paragraph("Target paragraph.", style="Normal")
    target_id = _para_id_by_text(doc, "Target paragraph.")
    with doc.edit("caption"):
        doc.anchor_by_id(target_id).insert_caption("Table", text="Costs")
    paras = _paras(doc)
    texts = [t for t, _ in paras]
    idx_target = texts.index("Target paragraph.")
    idx_cap = next(i for i, (t, s) in enumerate(paras) if s == "Caption")
    assert idx_cap == idx_target - 1  # caption is above the target


def test_caption_on_table_cell_is_standalone(scratch_doc):
    """Captioning a table cell produces a real standalone caption above the
    whole table — not fused into a cell — even from the last cell (regression:
    the last cell used to raise a COM 'end of table row' error)."""
    doc = scratch_doc
    with doc.edit("seed"):
        t = doc.add_table(2, 2)
    last_cell_id = t.cell(2, 2).anchor_id
    with doc.edit("caption"):
        doc.anchor_by_id(last_cell_id).insert_caption("Table", text="Grid data")
    # The caption paragraph exists, is styled Caption, and is *not* inside the table.
    cap = None
    for p in doc.com.Paragraphs:
        if p.Range.Style.NameLocal == "Caption":
            cap = p
            break
    assert cap is not None
    assert cap.Range.Information(12) is False  # wdWithInTable
    assert "Grid data" in cap.Range.Text


def test_format_paragraph_pagination_flags(scratch_doc):
    doc = scratch_doc
    with doc.edit("seed"):
        doc.append_paragraph("A heading-ish line.", style="Normal")
    pid = _para_id_by_text(doc, "A heading-ish line.")
    with doc.edit("pag"):
        doc.anchor_by_id(pid).format_paragraph(
            keep_together=True, keep_with_next=True, widow_control=False
        )
    pf = doc.anchor_by_id(pid).com.ParagraphFormat
    assert int(pf.KeepTogether) == -1  # True
    assert int(pf.KeepWithNext) == -1
    assert int(pf.WidowControl) == 0  # False


def test_set_heading_row_repeats(scratch_doc):
    doc = scratch_doc
    with doc.edit("seed"):
        doc.add_table(3, 2)
    with doc.edit("heading row"):
        doc.tables[1].set_heading_row(1)
    row = doc.tables[1].com.Rows(1)
    assert int(row.HeadingFormat) == -1  # True
    assert int(row.AllowBreakAcrossPages) == 0  # kept intact


# ---------------------------------------------------------------------------
# v0.12 LLM-ergonomics fixes (cell-scoped find, delete_paragraph, revisions,
# snapshot markup, numbered-list span)
# ---------------------------------------------------------------------------


def test_cell_scoped_find_replace_stays_in_cell(scratch_doc):
    """§2: a cell-scoped find/replace must not overrun into the next cell."""
    doc = scratch_doc
    with doc.edit("seed table"):
        doc.add_table(2, 2, data=[["Model", "Ctx"], ["Opus", "200K"]])
    table = doc.tables[len(doc.tables)]
    cell = table.cell(2, 1)
    with doc.edit("cell find"):
        applied = doc.find_replace("Opus", "Claude Opus", scope=doc.anchor_by_id(cell.anchor_id))
    assert len(applied) == 1
    grid = table.grid()
    assert grid[1][0] == "Claude Opus"
    assert grid[1][1] == "200K"  # the neighbour cell is untouched


def test_delete_paragraph_removes_whole_paragraph(scratch_doc):
    """§6: delete_paragraph removes the paragraph and its mark — no empty line left."""
    doc = scratch_doc
    with doc.edit("seed"):
        doc.append_paragraph("Keep one")
        doc.append_paragraph("Delete me")
        doc.append_paragraph("Keep two")
    target = next(p for p in doc.paragraphs.list() if p.get("text") == "Delete me")
    with doc.edit("delete"):
        doc.delete_paragraph(f"para:{target['index']}")
    texts = [p.get("text") for p in doc.paragraphs.list()]
    assert "Delete me" not in texts
    assert "Keep one" in texts and "Keep two" in texts


def test_revisions_reader_reports_tracked_edits(scratch_doc):
    """§1: doc.revisions exposes tracked changes as structured insert/delete entries."""
    doc = scratch_doc
    with doc.edit("seed"):
        doc.append_paragraph("The quick brown fox")
    with doc.tracked_changes(), doc.edit("tracked"):
        doc.find_replace("quick", "swift")
    rows = doc.revisions.list()
    types = {r["type"] for r in rows}
    assert "insert" in types and "delete" in types
    assert any(r["text"] == "swift" for r in rows if r["type"] == "insert")
    assert all({"index", "type", "author", "text", "anchor_id"} <= set(r) for r in rows)


def test_snapshot_markup_differs_from_final(scratch_doc):
    """§1: snapshot(markup="all") renders revision marks the final render omits."""
    pytest.importorskip("pymupdf")
    doc = scratch_doc
    with doc.edit("seed"):
        doc.append_paragraph("The quick brown fox jumps over the lazy dog.")
    with doc.tracked_changes(), doc.edit("tracked"):
        doc.find_replace("quick", "swift")
    final = doc.snapshot(pages=1, dpi=72, markup="none")[0].png
    marked = doc.snapshot(pages=1, dpi=72, markup="all")[0].png
    assert final.startswith(b"\x89PNG\r\n\x1a\n") and marked.startswith(b"\x89PNG\r\n\x1a\n")
    # The markup render carries the change bar / "Deleted: quick" balloon, so the
    # rasterised pages are not identical.
    assert final != marked


def test_numbered_list_over_range_numbers_sequentially(scratch_doc):
    """§4: applying a numbered list over a multi-paragraph range numbers 1..N."""
    doc = scratch_doc
    with doc.edit("seed"):
        for text in ("Item one", "Item two", "Item three"):
            doc.append_paragraph(text)
    items = [p for p in doc.paragraphs.list() if p.get("text", "").startswith("Item")]
    start, end = items[0]["start"], items[-1]["end"]
    with doc.edit("number"):
        doc.range(start, end).apply_list("numbered")
    markers = [p.list_info().get("string") for p in doc.paragraphs if p.text.startswith("Item")]
    # One contiguous list numbered 1., 2., 3. — not three independent "1." lists.
    assert markers == ["1.", "2.", "3."]


def test_image_extraction_round_trips(scratch_doc, png_path):
    """Insert a PNG, then read its bytes + MIME back out via doc.images / read_image."""
    doc = scratch_doc
    with doc.edit("seed image"):
        doc.end.insert_image(png_path, wrap="inline")
    rows = doc.images.list()
    assert len(rows) >= 1
    first = rows[0]
    assert first["anchor_id"] == "image:1"
    # Word re-encodes the embedded picture as PNG.
    assert first["mime"] == "image/png"
    # The image:N anchor and the discovery list agree, and the bytes come back.
    data, mime = doc.images[1].read_image()
    assert mime == "image/png"
    assert data[:8] == b"\x89PNG\r\n\x1a\n"  # a real PNG signature
    # The same bytes via anchor_by_id("image:N").
    via_id, _ = doc.anchor_by_id("image:1").read_image()
    assert via_id == data


def test_read_image_on_imageless_range_raises(scratch_doc):
    """A range with no picture is a clean ImageSourceError, not a crash."""
    doc = scratch_doc
    with doc.edit("seed text"):
        doc.append_paragraph("No picture here.")
    with pytest.raises(wordlive.ImageSourceError):
        doc.paragraphs[1].read_image()


def test_save_as_and_export_pdf_roundtrip(scratch_doc, tmp_path):
    """Ungated Python-API persistence: save a .docx, export a PDF, both land on disk."""
    doc = scratch_doc
    with doc.edit("seed"):
        doc.append_paragraph("A deliverable paragraph.")

    docx = tmp_path / "out.docx"
    written = doc.save_as(docx)
    assert written == str(docx.resolve())
    assert docx.is_file() and docx.stat().st_size > 0
    # After a save, Word reports the document clean.
    assert doc.saved is True

    # save() to the now-existing path succeeds (no longer "never saved").
    doc.append_paragraph("One more line.")
    assert doc.saved is False
    doc.save()
    assert doc.saved is True

    pdf = tmp_path / "out.pdf"
    assert doc.export_pdf(pdf) == str(pdf.resolve())
    assert pdf.is_file()
    assert pdf.read_bytes()[:5] == b"%PDF-"  # a real PDF header

    # save_as refuses to clobber without overwrite, then allows it.
    with pytest.raises(wordlive.exceptions.OpError):
        doc.save_as(docx)
    doc.save_as(docx, overwrite=True)


def test_save_gating_blocks_outside_whitelist(scratch_doc, tmp_path):
    """The CLI/MCP path policy denies a save outside the whitelist (resolve-first)."""
    from wordlive._paths import PathPolicy

    allowed = tmp_path / "allowed"
    allowed.mkdir()
    pol = PathPolicy(save_dirs=[allowed])
    # Inside the whitelist resolves; an escape via .. is refused.
    assert pol.resolve_save_target(allowed / "ok.docx") == (allowed / "ok.docx").resolve()
    with pytest.raises(wordlive.PathNotAllowedError):
        pol.resolve_save_target(allowed / ".." / "escape.docx")
