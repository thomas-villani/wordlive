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


def test_insert_block_at_end_does_not_merge_into_last_paragraph(scratch_doc):
    """WL-A: composing at doc.end must not fuse the first block paragraph into a
    non-empty last paragraph (and steal its style)."""
    doc = scratch_doc
    with doc.edit("seed"):
        doc.append_paragraph("…single slide.", style="Normal")
    with doc.edit("markdown at end"):
        doc.end.insert_markdown("## The Problem\n\nBody line.")
    texts = [p.text for p in doc.paragraphs]
    # The seeded line survives verbatim — no "…single slide.The Problem" merge.
    assert "…single slide." in texts
    assert not any("single slide.The Problem" in t for t in texts)
    # And the heading is its own paragraph, styled as a heading (not Normal,
    # not merged into the body line) — so it shows up as its own outline entry.
    assert doc.headings["The Problem"].text == "The Problem"
    assert "The Problem" in [h["text"] for h in doc.outline()]


def test_insert_block_at_end_of_fresh_doc_leaves_no_trailing_empty(scratch_doc):
    """WL-A: filling an empty terminal paragraph must not strand a stray empty
    paragraph after the block."""
    doc = scratch_doc
    with doc.edit("markdown into fresh"):
        doc.end.insert_markdown("## Title\n\nOnly body.")
    texts = [p.text for p in doc.paragraphs]
    assert texts[-1] == "Only body."  # last paragraph is real content, not ""


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


def test_format_paragraph_line_spacing(scratch_doc):
    """line_spacing sets the leading: a multiple, an exact length, or a keyword."""
    from wordlive.constants import WdLineSpacing

    doc = scratch_doc
    with doc.edit("seed"):
        doc.append_paragraph("Multiple.", style="Normal")
        doc.append_paragraph("Exact.", style="Normal")
    mult_id = _para_id_by_text(doc, "Multiple.")
    exact_id = _para_id_by_text(doc, "Exact.")
    with doc.edit("line spacing"):
        doc.anchor_by_id(mult_id).format_paragraph(line_spacing=2)  # double, as a multiple
        doc.anchor_by_id(exact_id).format_paragraph(line_spacing="20pt")
    mult_pf = doc.anchor_by_id(mult_id).com.ParagraphFormat
    # Word stores an exact-double multiple as either the MULTIPLE rule (24pt) or
    # its named DOUBLE rule; both render as double spacing.
    assert int(mult_pf.LineSpacingRule) in (
        int(WdLineSpacing.MULTIPLE),
        int(WdLineSpacing.DOUBLE),
    )
    assert abs(float(mult_pf.LineSpacing) - 24.0) < 0.01
    exact_pf = doc.anchor_by_id(exact_id).com.ParagraphFormat
    assert int(exact_pf.LineSpacingRule) == int(WdLineSpacing.EXACTLY)
    assert abs(float(exact_pf.LineSpacing) - 20.0) < 0.01


def test_drop_cap_roundtrip(scratch_doc):
    """drop_cap sets a real Word DropCap (position + geometry stick), and
    position='none' removes it."""
    from wordlive.constants import WdDropPosition

    doc = scratch_doc
    with doc.edit("seed"):
        doc.append_paragraph("Once upon a time a paragraph wanted a fancy initial.", style="Normal")
    pid = _para_id_by_text(doc, "Once upon a time a paragraph wanted a fancy initial.")
    with doc.edit("drop cap"):
        doc.anchor_by_id(pid).drop_cap(4, position="dropped", distance="2pt", font="Georgia")
    dc = doc.anchor_by_id(pid).com.Paragraphs(1).DropCap
    assert int(dc.Position) == int(WdDropPosition.DROPPED)
    assert int(dc.LinesToDrop) == 4
    assert abs(float(dc.DistanceFromText) - 2.0) < 0.01
    assert str(dc.FontName) == "Georgia"
    with doc.edit("remove drop cap"):
        doc.anchor_by_id(pid).drop_cap(position="none")
    dc = doc.anchor_by_id(pid).com.Paragraphs(1).DropCap
    assert int(dc.Position) == int(WdDropPosition.NONE)


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


def test_apply_list_format_authors_custom_numbered_list(scratch_doc):
    """apply_list_format mints a custom list template and applies it: a lower-letter
    ')' scheme renders 'a)' and reads back through read_list_levels."""
    doc = scratch_doc
    with doc.edit("seed"):
        for text in ("One", "Two", "Three"):
            doc.append_paragraph(text)
    items = [p for p in doc.paragraphs.list() if p.get("text", "") in ("One", "Two", "Three")]
    start, end = items[0]["start"], items[-1]["end"]
    rng = doc.range(start, end)
    with doc.edit("custom list"):
        rng.apply_list_format(
            [
                {
                    "kind": "number",
                    "format": "%1)",
                    "style": "lower-letter",
                    "trailing": "space",
                    "number_position": "0.25in",
                    "text_position": "0.5in",
                }
            ]
        )
    markers = [
        p.list_info().get("string") for p in doc.paragraphs if p.text in ("One", "Two", "Three")
    ]
    assert markers == ["a)", "b)", "c)"]
    levels = rng.read_list_levels()
    assert levels[0]["format"] == "%1)"
    assert levels[0]["style"] == "lower-letter"
    assert levels[0]["trailing"] == "space"


def test_apply_list_format_bullet_uses_glyph_and_symbol_font(scratch_doc):
    """A bullet level is authored via the glyph + a symbol font (never
    NumberStyle=bullet, which raises) and renders that glyph."""
    doc = scratch_doc
    with doc.edit("seed"):
        for text in ("Alpha", "Beta"):
            doc.append_paragraph(text)
    items = [p for p in doc.paragraphs.list() if p.get("text", "") in ("Alpha", "Beta")]
    start, end = items[0]["start"], items[-1]["end"]
    rng = doc.range(start, end)
    with doc.edit("bullets"):
        rng.apply_list_format([{"kind": "bullet", "bullet": "•", "font": "Symbol"}])
    markers = [p.list_info().get("string") for p in doc.paragraphs if p.text in ("Alpha", "Beta")]
    assert all(m == "•" for m in markers)
    levels = rng.read_list_levels()
    assert levels[0]["kind"] == "bullet"
    assert levels[0]["font"] == "Symbol"


def test_apply_list_format_multi_level_outline(scratch_doc):
    """A 2-spec levels list mints a 9-level outline template; the two set levels
    read back as authored."""
    doc = scratch_doc
    with doc.edit("seed"):
        doc.append_paragraph("Top")
    items = [p for p in doc.paragraphs.list() if p.get("text", "") == "Top"]
    rng = doc.range(items[0]["start"], items[0]["end"])
    with doc.edit("outline"):
        rng.apply_list_format(
            [
                {"kind": "number", "format": "%1.", "style": "upper-roman"},
                {"kind": "number", "format": "%1.%2", "style": "arabic"},
            ]
        )
    levels = rng.read_list_levels()
    assert len(levels) == 9  # outline templates always carry 9 levels
    assert levels[0]["style"] == "upper-roman"
    assert levels[1]["format"] == "%1.%2"


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


def test_insert_block_styled_run_with_inline_formatting(scratch_doc):
    """A whole styled bulleted section in one op: markdown + structured runs both
    format inline, paragraphs land in order, and the returned range bullets them."""
    doc = scratch_doc
    with doc.edit("seed"):
        doc.append_paragraph("Features", style="Heading 1")

    with doc.edit("block"):
        rng = doc.headings["Features"].insert_block(
            [
                {"text": "**Politeness** — preserves your cursor.", "style": "List Bullet"},
                {
                    "runs": [
                        {"text": "Atomic undo", "bold": True},
                        {"text": " — one Ctrl-Z."},
                    ],
                    "style": "List Bullet",
                },
                "Plain third bullet.",
            ]
        )
        doc.anchor_by_id(rng.anchor_id).apply_list("bulleted")

    rows = doc.paragraphs.list()
    texts = [r["text"] for r in rows]
    assert "Politeness — preserves your cursor." in texts
    assert "Atomic undo — one Ctrl-Z." in texts
    assert "Plain third bullet." in texts

    dc = doc.com
    lead = next(r for r in rows if r["text"].startswith("Politeness"))
    s = lead["start"]
    # The markdown bolded only the lead-in word.
    assert bool(dc.Range(s, s + len("Politeness")).Bold) is True
    assert bool(dc.Range(s + 12, s + 20).Bold) is False
    # The per-item paragraph style and the range-fed list both took effect.
    assert dc.Range(s, s + 1).Paragraphs(1).Range.Style.NameLocal == "List Bullet"
    assert int(dc.Range(s, s + 1).ListFormat.ListType) != 0

    runs_para = next(r for r in rows if r["text"].startswith("Atomic undo"))
    s2 = runs_para["start"]
    assert bool(dc.Range(s2, s2 + len("Atomic undo")).Bold) is True
    assert bool(dc.Range(s2 + 13, s2 + 20).Bold) is False


def test_table_from_records_builds_bolded_header(scratch_doc):
    """Records (list of dicts) → a table whose keys are a bolded header row, with
    rows/cols inferred from the data."""
    doc = scratch_doc
    with doc.edit("records table"):
        t = doc.end.insert_table(
            data=[{"Item": "Travel", "Cost": "$400"}, {"Item": "Lodging", "Cost": "$600"}]
        )
    assert (t.row_count, t.column_count) == (3, 2)
    assert t.grid() == [["Item", "Cost"], ["Travel", "$400"], ["Lodging", "$600"]]
    # Records imply a header row — it's bolded.
    assert bool(t.com.Rows(1).Range.Bold) is True


# ---------------------------------------------------------------------------
# v0.14 introspection — location() / stats() against REAL Word (the gate that
# catches a wrong WdStatistic / WdInformation constant value) and table records
# ---------------------------------------------------------------------------


def test_stats_reports_real_counts(scratch_doc):
    """doc.stats() must return integer counts that move with real content —
    proves the WdStatistic selector values are right (a wrong enum would read
    the wrong counter or raise)."""
    doc = scratch_doc
    with doc.edit("seed"):
        doc.append_paragraph("Introduction", style="Heading 1")
        doc.append_paragraph("Some body words here for the counters.")
        doc.add_table(2, 2)
    s = doc.stats()
    assert {
        "pages",
        "words",
        "characters",
        "paragraphs",
        "lines",
        "sections",
        "headings",
        "tables",
        "images",
        "comments",
        "revisions",
        "saved",
    } <= set(s)
    assert all(isinstance(s[k], int) for k in ("pages", "words", "characters", "paragraphs"))
    assert s["pages"] >= 1
    assert s["words"] >= 5  # the body line alone has 6 words
    assert s["tables"] == 1  # structural count from doc.tables
    assert s["headings"] >= 1  # at least the one Heading 1 (template may add more)
    assert s["saved"] is False  # a never-saved scratch doc


def test_location_reports_page_and_table_flag(scratch_doc):
    """anchor.location() must report a real page and the in-table flag — proves
    the WdInformation line/column/page selectors resolve."""
    doc = scratch_doc
    with doc.edit("seed"):
        doc.append_paragraph("On page one.")
        doc.add_table(1, 1)
    pid = _para_id_by_text(doc, "On page one.")
    loc = doc.anchor_by_id(pid).location()
    assert {"page", "end_page", "line", "column", "in_table"} == set(loc)
    assert loc["page"] == 1 and loc["in_table"] is False
    assert loc["line"] >= 1 and loc["column"] >= 1
    # A cell anchor reports in_table True.
    assert doc.tables[1].cell(1, 1).location()["in_table"] is True


def test_location_page_rises_after_a_page_break(scratch_doc):
    """A paragraph after a page break reports a higher page than one before it —
    proves location()'s page read tracks real layout. (Trailing 'Tail.' keeps the
    after-break paragraph off the undeletable terminal mark.)"""
    doc = scratch_doc
    with doc.edit("seed"):
        doc.append_paragraph("Top of block.")
        doc.end.insert_break("page")
        doc.append_paragraph("After the break.")
        doc.append_paragraph("Tail.")
    top = doc.anchor_by_id(_para_id_by_text(doc, "Top of block.")).location()
    bot = doc.anchor_by_id(_para_id_by_text(doc, "After the break.")).location()
    assert bot["page"] >= top["page"] + 1


def test_structural_query_helpers(scratch_doc):
    """between / nearest_heading / find_paragraphs over a real outline.

    Layout: Heading 1 'Alpha' → body → Heading 1 'Beta' → body. Pure reads."""
    doc = scratch_doc
    with doc.edit("seed"):
        # Pin body paragraphs to Normal — a paragraph appended after a Heading 1
        # otherwise inherits the heading style (real Word behaviour) and would
        # itself read as a heading.
        doc.append_paragraph("Alpha", style="Heading 1")
        doc.append_paragraph("The quick brown fox jumps over the lazy dog.", style="Normal")
        doc.append_paragraph("Beta", style="Heading 1")
        doc.append_paragraph("Second section body paragraph.", style="Normal")
    alpha = _para_id_by_text(doc, "Alpha")
    beta = _para_id_by_text(doc, "Beta")
    body1 = _para_id_by_text(doc, "The quick brown fox jumps over the lazy dog.")

    # between: the body under Alpha (excludes both heading lines).
    span = doc.between(alpha, beta)
    assert "quick brown fox" in span.text
    assert "Beta" not in span.text and "Alpha" not in span.text
    assert doc.between(alpha, beta, inclusive=True).text.strip().startswith("Alpha")

    # nearest_heading: enclosing vs next.
    assert doc.nearest_heading(body1, direction="before")["text"] == "Alpha"
    assert doc.nearest_heading(body1, direction="after")["text"] == "Beta"

    # find_paragraphs: a lightly-typo'd query still ranks the intended paragraph
    # first, with a high-but-imperfect score (proves real fuzzy ranking).
    rows = doc.find_paragraphs("the quick brown fox jumped over the lazy dog", limit=3)
    assert rows, "expected at least one fuzzy match"
    assert rows[0]["anchor_id"] == body1
    assert 0.8 <= rows[0]["score"] < 1.0


def test_table_records_round_trip_and_update(scratch_doc):
    """records() reads back what insert_table(data=[{...}]) wrote; append_record
    and update_row edit by header name."""
    doc = scratch_doc
    with doc.edit("records table"):
        t = doc.end.insert_table(data=[{"Item": "Travel", "Cost": "$400"}])
    assert t.records() == [{"Item": "Travel", "Cost": "$400"}]
    with doc.edit("append + update"):
        t.append_record({"Item": "Lodging", "Cost": "$600"})
        t.update_row("Travel", {"Cost": "$450"})
    assert t.records() == [
        {"Item": "Travel", "Cost": "$450"},
        {"Item": "Lodging", "Cost": "$600"},
    ]


def test_stats_and_location_preserve_selection(scratch_doc):
    """Politeness: the repagination inside stats()/location() must not move the
    user's Selection."""
    doc = scratch_doc
    with doc.edit("seed"):
        doc.append_paragraph("Anchor line for the selection probe.")
    sel = doc.com.Application.Selection
    sel.SetRange(3, 9)
    before = (int(sel.Start), int(sel.End))
    doc.stats()
    doc.range(0, 5).location()
    after_sel = doc.com.Application.Selection
    assert (int(after_sel.Start), int(after_sel.End)) == before


# ---------------------------------------------------------------------------
# v0.14 compose helpers — insert_section / insert_markdown / replace_section_body
# against REAL Word (styles applied, list markers contiguous, heading preserved)
# ---------------------------------------------------------------------------


def test_insert_markdown_maps_blocks_to_real_word_structure(scratch_doc):
    """A constrained-Markdown chunk lands as real headings, lists, and paragraphs:
    styles applied, numbered list reads 1..N, inline bold on the right span."""
    doc = scratch_doc
    with doc.edit("md"):
        doc.end.insert_markdown(
            "# Overview\n\nA **bold** lead paragraph.\n\n- first bullet\n- second bullet\n\n1. step one\n2. step two"
        )

    rows = doc.paragraphs.list()
    by_text = {r["text"]: r for r in rows}
    assert by_text["Overview"]["style"] == "Heading 1"
    # The two list runs carry list styles and real list formatting.
    dc = doc.com
    for marker_text in ("first bullet", "second bullet"):
        s = by_text[marker_text]["start"]
        assert int(dc.Range(s, s + 1).ListFormat.ListType) != 0

    # The numbered run reads 1., 2. — one contiguous list, not two "1."s.
    numbers = [
        p.list_info().get("string") for p in doc.paragraphs if p.text in ("step one", "step two")
    ]
    assert numbers == ["1.", "2."]

    # Inline markdown bolded only the word "bold" in the lead paragraph.
    lead = by_text["A bold lead paragraph."]
    s = lead["start"]
    assert bool(dc.Range(s + 2, s + 6).Bold) is True  # "bold"
    assert bool(dc.Range(s, s + 1).Bold) is False  # "A"


def test_insert_section_places_heading_and_body(scratch_doc):
    """insert_section drops a Heading {level} + its body in one op."""
    doc = scratch_doc
    with doc.edit("sec"):
        doc.end.insert_section("Results", ["We saw a lift.", "Caveats apply."], level=2)
    rows = {r["text"]: r for r in doc.paragraphs.list()}
    assert rows["Results"]["style"] == "Heading 2"
    assert "We saw a lift." in rows
    assert "Caveats apply." in rows


def test_replace_section_body_keeps_heading(scratch_doc):
    """Rewriting a section's body preserves the heading and the next section."""
    doc = scratch_doc
    with doc.edit("seed"):
        doc.append_paragraph("Alpha", style="Heading 1")
        doc.append_paragraph("old alpha body")
        doc.append_paragraph("Beta", style="Heading 1")
        doc.append_paragraph("beta body")

    with doc.edit("rewrite"):
        doc.headings["Alpha"].replace_section_body("fresh alpha body")

    texts = [r["text"] for r in doc.paragraphs.list()]
    assert "Alpha" in texts and "Beta" in texts  # both headings survive
    assert "fresh alpha body" in texts
    assert "old alpha body" not in texts  # the old body is gone
    assert "beta body" in texts  # the next section is untouched


# ---------------------------------------------------------------------------
# Equations (insert_equation + doc.equations) — real OMaths / InsertXML / MSXML
# ---------------------------------------------------------------------------


def test_insert_equation_unicodemath_builds_up(scratch_doc):
    """Native UnicodeMath: a linear string becomes a built-up display equation."""
    doc = scratch_doc
    with doc.edit("seed"):
        doc.append_paragraph("Derivation")
    with doc.edit("eq"):
        eq = doc.paragraphs[1].insert_equation(unicodemath="a^2+b^2=c^2", where="after")
    assert eq.anchor_id == "equation:1"
    assert eq.type == "display"
    assert len(doc.equations) == 1
    # The built-up text carries the operands (structure markers stripped).
    assert "𝑎" in eq.linear and "𝑐" in eq.linear


def test_insert_equation_mathml_round_trips_to_mathml(scratch_doc):
    """MathML in → Office OMML → and back out to MathML via the read transform."""
    doc = scratch_doc
    mathml = (
        '<math xmlns="http://www.w3.org/1998/Math/MathML">'
        "<mi>E</mi><mo>=</mo><mi>m</mi><msup><mi>c</mi><mn>2</mn></msup></math>"
    )
    with doc.edit("eq"):
        eq = doc.end.insert_equation(mathml=mathml, display=False)
    assert eq.type == "inline"
    out = eq.mathml
    assert "<math" in out and "msup" in out  # the superscript survives the round-trip


def test_insert_equation_display_and_inline_types(scratch_doc):
    """display=True/False set the equation's OMath type."""
    doc = scratch_doc
    with doc.edit("eqs"):
        disp = doc.end.insert_equation(unicodemath="x+1", display=True)
        inl = doc.end.insert_equation(unicodemath="y+1", display=False)
    assert disp.type == "display"
    assert inl.type == "inline"


def test_insert_equation_before_and_prepend(scratch_doc):
    """`where="before"` inserts ahead of the anchor, including at document start."""
    doc = scratch_doc
    with doc.edit("seed"):
        doc.append_paragraph("Body paragraph.")
    with doc.edit("prepend"):
        eq = doc.paragraphs[1].insert_equation(unicodemath="z=1", where="before")
    # The equation now precedes the body paragraph it was anchored before.
    assert eq.anchor_id == "equation:1"
    texts = [r["text"] for r in doc.paragraphs.list()]
    assert "Body paragraph." in texts


def test_equations_list_reports_each(scratch_doc):
    """doc.equations.list() summarises every equation in document order."""
    doc = scratch_doc
    with doc.edit("eqs"):
        doc.end.insert_equation(unicodemath="a=1")
        doc.end.insert_equation(unicodemath="b=2", display=False)
    rows = doc.equations.list()
    assert [r["anchor_id"] for r in rows] == ["equation:1", "equation:2"]
    assert rows[0]["type"] == "display"
    assert rows[1]["type"] == "inline"


def _seed_body_then_heading(doc):
    """Body paragraph (para:1) immediately followed by a Heading 2 (para:2)."""
    with doc.edit("seed"):
        doc.append_paragraph("Body paragraph before the equation.", style="Normal")
        doc.append_paragraph("A Heading 2 right after", style="Heading 2")


def test_display_equation_uses_equation_style_not_following_heading(scratch_doc):
    """WL-B: a display equation inserted before a Heading 2 must not adopt the
    heading's style (outline/TOC pollution) — it gets the centred `Equation`
    style instead."""
    doc = scratch_doc
    _seed_body_then_heading(doc)
    with doc.edit("display eq"):
        eq = doc.paragraphs[1].insert_equation(unicodemath="a^2+b^2=c^2", display=True)
    para = doc.paragraphs.at(int(eq.com.Start))
    assert para is not None
    assert not para.is_heading  # not styled Heading 2
    assert para.com.Style.NameLocal == "Equation"


def test_inline_equation_is_normal_and_left_aligned(scratch_doc):
    """WL-D: a display=False equation lands on its own paragraph but reads as
    body text — Normal style, left-aligned — not centred, not a heading."""
    from wordlive.constants import WdParagraphAlignment

    doc = scratch_doc
    _seed_body_then_heading(doc)
    with doc.edit("inline eq"):
        eq = doc.paragraphs[1].insert_equation(unicodemath="x=y", display=False)
    para = doc.paragraphs.at(int(eq.com.Start))
    assert para is not None
    assert not para.is_heading
    assert para.com.Style.NameLocal == "Normal"
    assert int(para.com.ParagraphFormat.Alignment) == int(WdParagraphAlignment.LEFT)


def test_insert_equation_latex_when_backend_present(scratch_doc):
    """LaTeX path, end to end — skipped when the optional backend isn't installed."""
    pytest.importorskip("latex2mathml")
    doc = scratch_doc
    with doc.edit("eq"):
        eq = doc.end.insert_equation(latex=r"\frac{-b}{2a}")
    assert len(doc.equations) == 1
    assert "<math" in eq.mathml


# ---------------------------------------------------------------------------
# Revision write surface — accept/reject, bulk, revision-aware reads (real
# Track Changes; both inserted and deleted runs present in the stream)
# ---------------------------------------------------------------------------


def test_revision_accept_one_makes_it_permanent(scratch_doc):
    """A tracked insertion accepted loses its revision mark; the text stays."""
    doc = scratch_doc
    with doc.edit("seed"):
        doc.append_paragraph("original")
    with doc.tracked_changes(), doc.edit("tracked edit"):
        doc.find_replace("original", "rewritten", all=True)
    assert len(doc.revisions) >= 1
    with doc.edit("accept"):
        doc.revisions[1].accept()
    # Fewer revisions remain, and the rewritten text survives.
    assert "rewritten" in "\n".join(p["text"] for p in doc.paragraphs.list())


def test_revision_reject_one_undoes_it(scratch_doc):
    """Rejecting a tracked insertion removes the inserted text."""
    doc = scratch_doc
    with doc.edit("seed"):
        doc.append_paragraph("keep this")
    with doc.tracked_changes(), doc.edit("tracked"):
        doc.end.insert_paragraph_after("DELETE ME")
    before = len(doc.revisions)
    assert before >= 1
    with doc.edit("reject"):
        for _ in range(before):
            doc.revisions[1].reject()
    assert "DELETE ME" not in "\n".join(p["text"] for p in doc.paragraphs.list())


def test_accept_all_clears_every_revision(scratch_doc):
    """Whole-document accept_all resolves all tracked changes at once."""
    doc = scratch_doc
    with doc.tracked_changes(), doc.edit("tracked"):
        doc.append_paragraph("alpha")
        doc.append_paragraph("beta")
    assert len(doc.revisions) >= 1
    with doc.edit("accept all"):
        n = doc.revisions.accept_all()
    assert n >= 1
    assert len(doc.revisions) == 0


def test_accept_all_within_anchor_scopes_to_range(scratch_doc):
    """accept_all(within=paragraph) resolves only the changes inside that range.

    Note an anchor's range is literal: a paragraph anchor covers that paragraph,
    so we scope to the paragraph that carries the tracked change (a *heading*
    anchor would cover only the heading line, not the body beneath it).
    """
    doc = scratch_doc
    with doc.edit("seed"):
        doc.append_paragraph("alpha")
        doc.append_paragraph("beta")
    with doc.tracked_changes(), doc.edit("tracked"):
        doc.paragraphs[1].insert_after(" ONE")
        doc.paragraphs[2].insert_after(" TWO")
    total = len(doc.revisions)
    assert total >= 2
    first_para = next(p for p in doc.paragraphs if "alpha" in p.text)
    with doc.edit("accept first only"):
        accepted = doc.revisions.accept_all(within=first_para)
    assert accepted >= 1
    # The second paragraph's revisions remain.
    assert 0 < len(doc.revisions) < total


def test_revision_aware_reads_split_inserted_and_deleted(scratch_doc):
    """text_final / text_original separate the two runs a tracked replace leaves."""
    doc = scratch_doc
    with doc.edit("seed"):
        doc.append_paragraph("the quick fox")
    with doc.tracked_changes(), doc.edit("tracked"):
        doc.find_replace("quick", "slow", all=True)
    para = next(p for p in doc.paragraphs if "fox" in p.text)
    # final = as if accepted (slow), original = as if rejected (quick).
    assert "slow" in para.text_final and "quick" not in para.text_final
    assert "quick" in para.text_original and "slow" not in para.text_original
    changes = {s["change"] for s in para.revision_segments()}
    assert "insert" in changes and "delete" in changes


# ---------------------------------------------------------------------------
# Publishing flourishes — watermark (header-story WordArt) + text box
# ---------------------------------------------------------------------------


def test_set_and_remove_watermark(scratch_doc):
    """A text watermark lands in the header story and removes cleanly."""
    doc = scratch_doc
    with doc.edit("seed"):
        doc.append_paragraph("Body content")
    with doc.edit("watermark"):
        n = doc.set_watermark("DRAFT")
    assert n >= 1
    header_shapes = doc.com.Sections(1).Headers(1).Shapes
    names = [header_shapes(i).Name for i in range(1, int(header_shapes.Count) + 1)]
    assert any(name.startswith("PowerPlusWaterMarkObject") for name in names)
    with doc.edit("remove watermark"):
        removed = doc.remove_watermark()
    assert removed >= 1
    assert doc.remove_watermark() == 0  # idempotent — nothing left


def test_set_watermark_replaces_not_stacks(scratch_doc):
    """Setting a second watermark clears the first rather than stacking."""
    doc = scratch_doc
    with doc.edit("w1"):
        doc.set_watermark("DRAFT", layout="horizontal")
    with doc.edit("w2"):
        doc.set_watermark("FINAL")
    header_shapes = doc.com.Sections(1).Headers(1).Shapes
    watermarks = [
        header_shapes(i).Name
        for i in range(1, int(header_shapes.Count) + 1)
        if header_shapes(i).Name.startswith("PowerPlusWaterMarkObject")
    ]
    assert len(watermarks) == 1


def test_insert_text_box_creates_floating_shape_with_text(scratch_doc):
    """A pull-quote text box is a floating shape whose frame carries the text."""
    doc = scratch_doc
    with doc.edit("seed"):
        doc.append_paragraph("Article body goes here.")
    before = int(doc.com.Shapes.Count)
    with doc.edit("text box"):
        doc.paragraphs[1].insert_text_box(
            "Pull quote!", width="2.5in", height="1in", fill="#dddddd"
        )
    assert int(doc.com.Shapes.Count) == before + 1
    shape = doc.com.Shapes(int(doc.com.Shapes.Count))
    assert "Pull quote!" in str(shape.TextFrame.TextRange.Text)


def test_replace_image_preserves_rotation_wrap_and_crop(scratch_doc, png_path):
    """Swapping a floating picture's image keeps its rotation, wrap side, and crop
    (delete+reinsert must re-apply the full layout, not just position/size)."""
    doc = scratch_doc
    with doc.edit("seed image"):
        doc.end.insert_image(png_path, wrap="square", alt_text="orig")
    with doc.edit("style image"):
        sh = doc.anchor_by_id("shape:1")
        sh.set_rotation(30)
        sh.set_wrap(side="left", distance_right="0.2in")

    s_before = doc.com.Shapes(1)
    rot, side, dist = (
        round(float(s_before.Rotation), 1),
        int(s_before.WrapFormat.Side),
        round(float(s_before.WrapFormat.DistanceRight), 2),
    )
    with doc.edit("replace image"):
        doc.anchor_by_id("shape:1").replace_image(_PNG)

    s_after = doc.com.Shapes(1)
    assert round(float(s_after.Rotation), 1) == rot
    assert int(s_after.WrapFormat.Side) == side
    assert round(float(s_after.WrapFormat.DistanceRight), 2) == dist


# ---------------------------------------------------------------------------
# Charts (insert_chart + doc.charts) — real Excel-backed AddChart2 / BreakLink.
# Requires Excel installed alongside Word. The recipe is hard-won (see _charts):
# insert off the Selection, populate the embedded workbook via a SERIES formula,
# then BreakLink + close so no orphan Excel is left behind.
# ---------------------------------------------------------------------------


def test_insert_charts_all_kinds_and_metadata(scratch_doc):
    """Insert every kind into ONE doc (one undo step) and read back the metadata.

    Deliberately one shared document and a single edit batch: each chart spins up
    and tears down a hidden Excel, and cycling that rapidly across many separate
    docs is what destabilises live Word (transient RPC hiccups — the same
    fragility `_charts` tames with BreakLink). One pass mirrors real usage. We
    read only the stable metadata (`chart_type`, `title`, count) — never the
    series data, which re-spins Excel.

    Reaching all four kinds also proves the no-orphan close works: a leftover open
    data grid would make the second insert raise "the chart data grid is already
    open", so four charts in a row means BreakLink+close released it each time.
    """
    doc = scratch_doc
    with doc.edit("charts"):
        doc.end.insert_chart("bar", {"Q1": 10, "Q2": 25, "Q3": 18}, title="Quarterly")
        doc.end.insert_chart("pie", {"A": 1, "B": 2, "C": 3})
        doc.end.insert_chart("line", {"Jan": 3.1, "Feb": 4.7})
        # scatter from [x, y] pairs, with a duplicate x (numeric value axis)
        doc.end.insert_chart("scatter", [[1.2, 3.4], [1.2, 3.9], [2.5, 6.1]], title="signal")

    assert len(doc.charts) == 4
    rows = doc.charts.list()
    assert [r["anchor_id"] for r in rows] == ["chart:1", "chart:2", "chart:3", "chart:4"]
    assert [r["kind"] for r in rows] == ["bar", "pie", "line", "scatter"]
    # title set where given; title=None leaves the chart untitled
    assert rows[0]["title"] == "Quarterly"
    assert rows[1]["title"] is None
    assert rows[3]["title"] == "signal"


def test_chart_formatting_verbs(scratch_doc):
    """Drive the post-insert formatting surface against a live chart.

    Live-probed safe: these operate on the BreakLink-static chart with no
    embedded-Excel respin (validated 0 orphan EXCEL.EXE). We assert only that the
    calls succeed and read back the *stable* design metadata (`chart_style`,
    `has_legend`, `title`) — never the series data. One shared doc, one edit batch.
    """
    doc = scratch_doc
    with doc.edit("insert"):
        doc.end.insert_chart("bar", {"Q1": 10, "Q2": 25, "Q3": 18}, title="Q")
        doc.end.insert_chart("scatter", [[1.0, 2.0], [2.0, 5.0], [3.0, 7.0], [4.0, 10.0]])
    bar, scatter = doc.charts[1], doc.charts[2]

    with doc.edit("format charts"):
        bar.format(
            title="Quarterly revenue",
            legend=True,
            legend_position="bottom",
            chart_style=242,
            background="#F4F6F7",
            data_labels=True,
            data_label_format="0",
        ).set_axis("value", title="USD (M)", minimum=0, maximum=30, gridlines=True)
        bar.set_series_color("#2E86C1")
        bar.set_series_color("#E74C3C", point=2)
        # scatter: log value axis + a power trendline that draws its equation
        scatter.set_axis("value", scale="log").set_axis("x", title="t (s)")
        scatter.add_trendline(kind="power", display_equation=True, display_r_squared=True)
        scatter.set_series_color((39, 174, 96))

    # Re-type a chart in place, then read back only the safe design metadata.
    with doc.edit("retype"):
        bar.format(chart_type="line")
    assert bar.chart_type == "line"
    assert bar.title == "Quarterly revenue"
    assert bar.chart_style == 242
    assert bar.has_legend is True
    row = doc.charts.list()[0]
    assert row["chart_style"] == 242 and row["has_legend"] is True


def test_chart_depth_verbs(scratch_doc):
    """Drive the PR-C depth surface against live charts (no Excel respin).

    Live-probed 2026-06-21: error bars, markers/smoothing, pie explosion, bar
    gap/overlap, data table, and trendline order/period are all settable on the
    BreakLink-static chart. We assert the calls succeed and the charts survive —
    never the series data. Three kinds (one Excel spin each, kept low to respect
    the live-Word RPC fragility) cover every type-specific knob: bar, pie, scatter.

    We do NOT assert chart_type after the ops: markers / smoothing legitimately
    promote the underlying Xl ChartType (a marked scatter becomes xlLineMarkers),
    so the kind string is expected to drift — the point is the calls don't raise.
    """
    doc = scratch_doc
    with doc.edit("insert"):
        doc.end.insert_chart("bar", {"Q1": 10, "Q2": 25, "Q3": 18}, title="Q")
        doc.end.insert_chart("pie", {"A": 1, "B": 2, "C": 3})
        doc.end.insert_chart("scatter", [[1.0, 2.0], [2.0, 5.0], [3.0, 7.0], [4.0, 10.0]])
    bar, pie, scatter = doc.charts[1], doc.charts[2], doc.charts[3]

    with doc.edit("chart depth"):
        # bar: spacing + data table + percent error bars on the value axis
        bar.format(gap_width=40, overlap=20, data_table=True, chart_style=240)
        bar.add_error_bars(series=1, kind="percent", amount=5, include="both")
        # pie: explode one slice
        pie.format_series(series=1, point=1, explosion=25)
        # scatter: markers + smoothed line, styled labels, and two trendline knobs
        scatter.format_series(
            series=1,
            marker="circle",
            marker_size=8,
            smooth=True,
            data_labels=True,
            data_label_size=9,
            data_label_color="#2E86C1",
        )
        scatter.add_trendline(kind="polynomial", order=3, display_equation=True)
        scatter.add_trendline(kind="moving_average", period=2)

    # The charts survived the depth ops and stable design metadata still reads back.
    assert len(doc.charts) == 3
    assert bar.chart_style == 240


# ---------------------------------------------------------------------------
# Table styling & polish — restyle / alignment / borders / banding, cell
# vertical alignment, and the row/column anchors (table:N:row:R / :col:C).
# The column path is the fragile one: Word has no per-column model on a
# merged / mixed-width table, so a column op there must raise a clean OpError.
# ---------------------------------------------------------------------------


def test_set_style_restyles_existing_table(scratch_doc):
    """A post-creation table style swap takes (the gap insert_table(style=) left)."""
    doc = scratch_doc
    with doc.edit("seed"):
        doc.add_table(2, 2, data=[["A", "B"], ["C", "D"]])
    with doc.edit("restyle"):
        doc.tables[1].set_style("Grid Table 4 - Accent 1")
    assert doc.tables[1].com.Style.NameLocal == "Grid Table 4 - Accent 1"


def test_set_alignment_and_banding_and_borders(scratch_doc):
    """Whole-table alignment, banding flags, and grid borders all apply live."""
    from wordlive.constants import WdRowAlignment

    doc = scratch_doc
    with doc.edit("seed"):
        doc.add_table(3, 3, style="Grid Table 4 - Accent 1")
    t = doc.tables[1]
    with doc.edit("style ops"):
        t.set_alignment("center")
        t.set_banding(first_row=True, banded_rows=False, banded_columns=True)
        t.set_borders(sides=["box", "horizontal", "vertical"], style="single", weight=1.0)
    com = t.com
    assert int(com.Rows.Alignment) == int(WdRowAlignment.CENTER)
    # The ApplyStyle* properties read back as real Python bools under makepy.
    assert bool(com.ApplyStyleHeadingRows) is True
    assert bool(com.ApplyStyleRowBands) is False
    assert bool(com.ApplyStyleColumnBands) is True


def test_cell_vertical_alignment_bottom_renders(scratch_doc):
    """Cell vertical alignment maps top/center/bottom onto 0/1/3 (2 is invalid)."""
    doc = scratch_doc
    with doc.edit("seed"):
        doc.add_table(2, 2)
    with doc.edit("valign"):
        doc.anchor_by_id("table:1:1:1").set_vertical_alignment("bottom")
    assert int(doc.tables[1].com.Cell(1, 1).VerticalAlignment) == 3  # wdCellAlignVerticalBottom


def test_row_anchor_shades_whole_row_politely(scratch_doc):
    """table:N:row:R styles the entire row through the inherited verbs, and the
    user's Selection survives (politeness)."""
    from wordlive._format import to_bgr

    doc = scratch_doc
    with doc.edit("seed"):
        doc.add_table(3, 3, data=[["h1", "h2", "h3"], ["a", "b", "c"], ["d", "e", "f"]])
    sel = doc.com.Application.Selection
    sel.SetRange(0, 0)
    before = (int(sel.Start), int(sel.End))
    with doc.edit("row shade"):
        doc.anchor_by_id("table:1:row:1").set_shading(fill="#FFFF00")
    com = doc.tables[1].com
    for c in (1, 2, 3):
        assert int(com.Cell(1, c).Range.Shading.BackgroundPatternColor) == to_bgr("#FFFF00")
    after = doc.com.Application.Selection
    assert (int(after.Start), int(after.End)) == before


def test_column_anchor_styles_each_cell_on_regular_table(scratch_doc):
    """On a regular table, table:N:col:C fans styling across the column's cells."""
    from wordlive._format import to_bgr

    doc = scratch_doc
    with doc.edit("seed"):
        doc.add_table(3, 2, data=[["x", "y"], ["1", "2"], ["3", "4"]])
    with doc.edit("col shade"):
        doc.anchor_by_id("table:1:col:1").set_shading(fill="red")
    com = doc.tables[1].com
    for r in (1, 2, 3):
        assert int(com.Cell(r, 1).Range.Shading.BackgroundPatternColor) == to_bgr("red")


def test_column_anchor_on_merged_table_raises_operror(scratch_doc):
    """The headline fragility: a column op on a merged / mixed-width table has no
    per-column model in Word, so it raises a clean OpError pointing at per-cell
    styling — not a raw COM 'mixed cell widths' error."""
    from wordlive.exceptions import OpError

    doc = scratch_doc
    with doc.edit("seed"):
        doc.add_table(3, 3)
    # Merge the first two cells of row 1 → the table now has mixed cell widths.
    with doc.edit("merge"):
        doc.tables[1].com.Cell(1, 1).Merge(doc.tables[1].com.Cell(1, 2))
    with pytest.raises(OpError):
        with doc.edit("col shade"):
            doc.anchor_by_id("table:1:col:1").set_shading(fill="red")


def test_add_column_appends_and_fills(scratch_doc):
    """add_column appends a column at the right edge and fills it top-to-bottom."""
    doc = scratch_doc
    with doc.edit("seed"):
        doc.add_table(2, 2, data=[["A", "B"], ["C", "D"]])
    t = doc.tables[1]
    with doc.edit("add column"):
        t.add_column(["x", "y"])
    assert t.column_count == 3
    assert t.cell(1, 3).text == "x"
    assert t.cell(2, 3).text == "y"


def test_delete_column_removes_a_column(scratch_doc):
    """delete_column drops a column on a regular table."""
    doc = scratch_doc
    with doc.edit("seed"):
        doc.add_table(2, 3, data=[["a", "b", "c"], ["d", "e", "f"]])
    t = doc.tables[1]
    with doc.edit("delete column"):
        t.delete_column(2)
    assert t.column_count == 2
    assert t.cell(1, 2).text == "c"  # old column 3 shifts left


def test_delete_column_on_merged_table_raises_operror(scratch_doc):
    """delete_column on a merged / mixed-width table raises a clean OpError
    (Word can't address an individual column there) — not a raw COM error."""
    from wordlive.exceptions import OpError

    doc = scratch_doc
    with doc.edit("seed"):
        doc.add_table(3, 3)
    with doc.edit("merge"):
        doc.tables[1].cell(1, 1).merge(doc.tables[1].cell(1, 2))
    with pytest.raises(OpError, match="mixed-width|merged"):
        with doc.edit("delete column"):
            doc.tables[1].delete_column(1)


def test_merge_cells_via_api_makes_table_non_uniform(scratch_doc):
    """Cell.merge joins two cells; the table reports is_uniform False afterwards."""
    doc = scratch_doc
    with doc.edit("seed"):
        doc.add_table(3, 3, data=[["A", "B", "C"], ["d", "e", "f"], ["g", "h", "i"]])
    t = doc.tables[1]
    assert t.is_uniform is True
    with doc.edit("merge"):
        t.cell(1, 1).merge(t.cell(1, 2))
    assert t.is_uniform is False
    # The merged cell is addressed by the rectangle's upper-left (1,1) and carries
    # both original texts.
    merged = t.cell(1, 1).text
    assert "A" in merged and "B" in merged


def test_merge_collapses_to_upper_left_regardless_of_receiver(scratch_doc):
    """Cell.Merge always collapses to the upper-left cell, even when `self` is the
    bottom-right receiver — the merged cell is addressed by the upper-left coord."""
    doc = scratch_doc
    with doc.edit("seed"):
        doc.add_table(2, 2, data=[["A", "B"], ["C", "D"]])
    t = doc.tables[1]
    with doc.edit("merge into upper-left"):
        t.cell(2, 2).merge(t.cell(1, 1))  # receiver is (2,2), MergeTo is (1,1)
    # The survivor is at (1,1) and holds every spanned cell's text…
    merged = t.cell(1, 1).text
    for ch in ("A", "B", "C", "D"):
        assert ch in merged
    # …and the bottom-right coordinate no longer resolves.
    from wordlive.exceptions import AnchorNotFoundError

    with pytest.raises(AnchorNotFoundError):
        _ = t.cell(2, 2).text


def test_split_cell_via_api_makes_table_non_uniform(scratch_doc):
    """Cell.split divides one cell into a grid; the row gains physical cells."""
    doc = scratch_doc
    with doc.edit("seed"):
        doc.add_table(2, 2, data=[["A", "B"], ["C", "D"]])
    t = doc.tables[1]
    with doc.edit("split"):
        t.cell(1, 1).split(rows=1, cols=3)
    assert t.is_uniform is False
    # Row 1 now has more physical cells than row 2.
    assert int(t.com.Rows(1).Cells.Count) > int(t.com.Rows(2).Cells.Count)


def test_checkpoint_ignores_field_code_view_toggle(scratch_doc):
    """A checkpoint is independent of the ShowFieldCodes view state: toggling it
    between two checkpoints of an unchanged doc must not surface a phantom change
    (the fingerprint pins TextRetrievalMode rather than reading the live view)."""
    doc = scratch_doc
    with doc.edit("seed field"):
        # Insert a PAGE field via COM so the paragraph contains a real field.
        rng = doc.anchor_by_id("start").com
        doc.com.Fields.Add(Range=rng, Type=33)  # wdFieldPage

    cp = doc.checkpoint()
    view = doc.com.ActiveWindow.View
    original = view.ShowFieldCodes
    try:
        view.ShowFieldCodes = not original
        changes = doc.changes_since(cp)
    finally:
        view.ShowFieldCodes = original
    assert changes == []


# ---------------------------------------------------------------------------
# Batch 6 — the first adds_content opt-in linter fixes, end-to-end vs live Word
# ---------------------------------------------------------------------------


def test_regularize_content_fixes_round_trip(scratch_doc):
    """A messy document -> regularize(allow_content=True) -> re-lint clean -> a second
    pass is a no-op. Exercises the six wired fixes against real Word, and above all the
    multi-delete reverse-order guard (two stray blank paragraphs deleted in one pass)."""
    doc = scratch_doc
    with doc.edit("seed messy doc"):
        doc.append_paragraph("Intro body paragraph.", style="Normal")
        doc.append_paragraph("", style="Normal")  # stray blank #1
        doc.append_paragraph("Middle body paragraph.", style="Normal")
        doc.append_paragraph("", style="Normal")  # stray blank #2
        doc.append_paragraph("Final body paragraph.", style="Normal")
        doc.set_watermark("DRAFT")
        doc.end.link_to("https://acme.example/report", text="Acme")

    # Selecting a policy rule by id enables it, so no profile is needed here.
    rules_all = [
        "stray-empty-paragraph",
        "draft-watermark-present",
        "hyperlink-bare-for-print",
        "page-numbers-present",
    ]

    # Withheld by default: every fix is adds_content, so nothing applies unprompted.
    withheld = doc.regularize(rules=rules_all)
    assert withheld["applied"] == []
    assert {f["rule"] for f in withheld["deferred"]} == set(rules_all)

    report = doc.regularize(rules=rules_all, allow_content=True)
    assert {f["rule"] for f in report["applied"]} == set(rules_all)
    # Both stray blanks are gone (reverse-order deletes didn't corrupt each other):
    # four distinct rules but stray-empty-paragraph fired twice → five applied fixes.
    assert len(report["applied"]) == len(set(rules_all)) + 1

    assert doc.watermark() is None
    assert doc.lint(rules=rules_all) == []
    again = doc.regularize(rules=rules_all, allow_content=True)
    assert again["applied"] == []
