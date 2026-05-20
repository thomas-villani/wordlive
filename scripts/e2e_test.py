#!/usr/bin/env python
"""End-to-end test for wordlive against a real Microsoft Word installation.

Creates a fresh document via COM, exercises the public Python API + the CLI,
and prints PASS/FAIL per scenario. The temp document is closed without saving
on exit (override with --keep). The user's other open documents are not
touched.

Usage:

    python scripts/e2e_test.py                # run, close test doc, summarise
    python scripts/e2e_test.py --keep         # leave the test doc open
    python scripts/e2e_test.py --verbose      # print extra context per test
    python scripts/e2e_test.py --no-cli       # skip the subprocess CLI smoke

Requires: Windows + Microsoft Word installed. Will launch Word if it isn't
already running.
"""

from __future__ import annotations

import argparse
import json
import subprocess
import sys
import traceback
from contextlib import contextmanager
from typing import Any, Callable, Iterator

import wordlive as wl
from wordlive import AmbiguousMatchError, AnchorNotFoundError, StyleNotFoundError
from wordlive._document import Document


# ---------------------------------------------------------------------------
# Test document layout
# ---------------------------------------------------------------------------

# Each tuple is (paragraph text, style or None for body text).
# Headings give us outline coverage; tokens give us deterministic find targets:
#   - SINGLE_TOKEN   appears exactly once  → unambiguous find_replace
#   - MULTI_TOKEN    appears three times   → ambiguous + --occurrence + --all
#   - the curly-quote phrase tests Unicode-fuzzy matching
#   - [PLACEHOLDER] becomes a content control
#   - "wordlive E2E" becomes a bookmark
DOC_PARAGRAPHS: list[tuple[str, str | None]] = [
    ("Introduction", "Heading 1"),
    ("Welcome to the wordlive E2E test document.", None),
    ("Risks", "Heading 2"),
    ("Risk one is documented. Risk two is documented. The token SINGLE_TOKEN appears only here.", None),
    ("Action items", "Heading 2"),
    ("MULTI_TOKEN appears here. Then MULTI_TOKEN again. And MULTI_TOKEN a third time.", None),
    ("Smart quotes test", "Heading 2"),
    ("The phrase “hello world” uses curly quotes.", None),
    ("Conclusion", "Heading 1"),
    ("Signed by [PLACEHOLDER] on this date.", None),
]

BOOKMARK_NAME = "Project"
BOOKMARK_NEEDLE = "wordlive E2E"
CC_TITLE = "Signatory"
CC_PLACEHOLDER = "[PLACEHOLDER]"
CC_DEFAULT_TEXT = "Jane Doe"

TABLE_TITLE = "E2E Grid"
TABLE_CELLS = [["R1C1", "R1C2"], ["R2C1", "R2C2"]]


# ---------------------------------------------------------------------------
# Tiny test harness
# ---------------------------------------------------------------------------


class Harness:
    def __init__(self, *, verbose: bool) -> None:
        self.verbose = verbose
        self.results: list[tuple[str, bool, str]] = []

    def run(self, name: str, fn: Callable[[], None]) -> None:
        try:
            fn()
        except Exception as e:  # noqa: BLE001
            tb = traceback.format_exc() if self.verbose else f"{type(e).__name__}: {e}"
            print(f"[FAIL] {name}: {tb.strip().splitlines()[-1]}")
            if self.verbose:
                print(tb)
            self.results.append((name, False, str(e)))
            return
        print(f"[PASS] {name}")
        self.results.append((name, True, ""))

    @property
    def passed(self) -> int:
        return sum(1 for _, ok, _ in self.results if ok)

    @property
    def failed(self) -> int:
        return sum(1 for _, ok, _ in self.results if not ok)

    def summary(self) -> str:
        total = len(self.results)
        return f"{self.passed}/{total} passed, {self.failed} failed"


def expect(cond: bool, msg: str) -> None:
    if not cond:
        raise AssertionError(msg)


# ---------------------------------------------------------------------------
# Document setup / teardown
# ---------------------------------------------------------------------------


def build_doc(word: wl.Word) -> Document:
    """Create a fresh blank document and populate it via COM."""
    app = word.com
    new = app.Documents.Add()

    # Wipe the default empty paragraph so paragraph indices line up with our list.
    new.Content.Text = ""

    body = "".join(text + "\r" for text, _ in DOC_PARAGRAPHS)
    new.Content.Text = body

    # Apply heading styles by paragraph index (1-based).
    for idx, (_, style) in enumerate(DOC_PARAGRAPHS, start=1):
        if style is None:
            continue
        try:
            new.Paragraphs(idx).Range.Style = new.Styles(style)
        except Exception as e:  # noqa: BLE001
            print(f"  warn: could not apply style {style!r} to paragraph {idx}: {e}")

    # Add a bookmark around a known phrase in the body.
    full_text = str(new.Content.Text)
    bm_pos = full_text.find(BOOKMARK_NEEDLE)
    if bm_pos >= 0:
        bm_range = new.Range(bm_pos, bm_pos + len(BOOKMARK_NEEDLE))
        new.Bookmarks.Add(Name=BOOKMARK_NAME, Range=bm_range)

    # Convert [PLACEHOLDER] into a Text content control with a default value.
    cc_pos = str(new.Content.Text).find(CC_PLACEHOLDER)
    if cc_pos >= 0:
        target = new.Range(cc_pos, cc_pos + len(CC_PLACEHOLDER))
        wd_content_control_text = 0  # WdContentControlType.wdContentControlText
        cc = new.ContentControls.Add(wd_content_control_text, target)
        cc.Title = CC_TITLE
        cc.Tag = CC_TITLE.lower()
        cc.Range.Text = CC_DEFAULT_TEXT

    # Append a 2x2 table at the end of the document for the table tests.
    tail = new.Content
    tail.Collapse(0)  # wdCollapseEnd
    tbl = new.Tables.Add(tail, len(TABLE_CELLS), len(TABLE_CELLS[0]))
    try:
        tbl.Title = TABLE_TITLE
    except Exception as e:  # noqa: BLE001 — Title needs a recent Word; non-fatal
        print(f"  warn: could not set table title: {e}")
    for r, row in enumerate(TABLE_CELLS, start=1):
        for c, val in enumerate(row, start=1):
            tbl.Cell(r, c).Range.Text = val

    return Document(word, new)


@contextmanager
def temp_doc(word: wl.Word, *, keep: bool) -> Iterator[Document]:
    doc = build_doc(word)
    try:
        yield doc
    finally:
        if keep:
            print(f"(leaving test document open: {doc.name})")
            return
        try:
            doc.com.Close(SaveChanges=0)  # wdDoNotSaveChanges
        except Exception as e:  # noqa: BLE001
            print(f"  warn: could not close test document: {e}")


# ---------------------------------------------------------------------------
# Individual tests
# ---------------------------------------------------------------------------


def t_status(word: wl.Word, doc: Document) -> None:
    rows = word.documents.list()
    expect(isinstance(rows, list), "status should return a list")
    names = [r["name"] for r in rows]
    expect(doc.name in names, f"test doc {doc.name!r} missing from status: {names}")


def t_outline(_word: wl.Word, doc: Document) -> None:
    items = doc.outline()
    texts = [it["text"] for it in items]
    for expected in ("Introduction", "Risks", "Action items", "Smart quotes test", "Conclusion"):
        expect(expected in texts, f"heading {expected!r} missing from outline: {texts}")
    expect(all(it["anchor_id"].startswith("heading:") for it in items), "outline anchor_ids must start with 'heading:'")


def t_anchor_by_id_heading(_word: wl.Word, doc: Document) -> None:
    items = doc.outline()
    # Pick "Risks" by anchor_id.
    risks = next(it for it in items if it["text"] == "Risks")
    h = doc.anchor_by_id(risks["anchor_id"])
    expect(h.text == "Risks", f"anchor_by_id heading text mismatch: {h.text!r}")
    expect(h.kind == "heading", f"expected heading kind, got {h.kind!r}")


def t_section_text(_word: wl.Word, doc: Document) -> None:
    body = doc.heading("Risks").section_text()
    expect("Risk one is documented" in body, f"section body missing expected line: {body!r}")
    # Section should stop before the next same-level heading.
    expect("Action items" not in body, f"section body bled past next heading: {body!r}")


def t_bookmark_read_and_write(_word: wl.Word, doc: Document) -> None:
    bm = doc.bookmarks[BOOKMARK_NAME]
    expect(bm.text == BOOKMARK_NEEDLE, f"bookmark text mismatch: {bm.text!r}")
    new_text = "wordlive smoke run"
    with doc.edit("E2E: write bookmark"):
        bm.set_text(new_text)
    expect(doc.bookmarks[BOOKMARK_NAME].text == new_text, "bookmark set_text round-trip failed")


def t_cc_read_and_write(_word: wl.Word, doc: Document) -> None:
    cc = doc.content_controls[CC_TITLE]
    expect(cc.text == CC_DEFAULT_TEXT, f"cc default text mismatch: {cc.text!r}")
    new_text = "Tomás Villani"
    with doc.edit("E2E: write cc"):
        cc.set_text(new_text)
    expect(doc.content_controls[CC_TITLE].text == new_text, "cc set_text round-trip failed")


def t_find_smart_quotes(_word: wl.Word, doc: Document) -> None:
    # Search with straight quotes; the doc has curly quotes.
    matches = doc.find('"hello world"')
    expect(len(matches) == 1, f"expected 1 fuzzy match for curly-quote phrase, got {len(matches)}: {matches}")
    expect("hello world" in matches[0]["text"], f"matched text missing phrase: {matches[0]['text']!r}")


def t_find_replace_single(_word: wl.Word, doc: Document) -> None:
    with doc.edit("E2E: find/replace SINGLE_TOKEN"):
        applied = doc.find_replace("SINGLE_TOKEN", "ONCE_REPLACED")
    expect(len(applied) == 1, f"expected exactly 1 replacement, got {len(applied)}")
    expect("ONCE_REPLACED" in str(doc.com.Content.Text), "replacement string missing from doc content")
    expect("SINGLE_TOKEN" not in str(doc.com.Content.Text), "original token still present after single replace")


def t_find_replace_ambiguous(_word: wl.Word, doc: Document) -> None:
    try:
        with doc.edit("E2E: ambiguous"):
            doc.find_replace("MULTI_TOKEN", "WRONG")
    except AmbiguousMatchError as e:
        expect(len(e.matches) >= 2, f"AmbiguousMatchError must carry >=2 matches, got {len(e.matches)}")
        expect("MULTI_TOKEN" in str(doc.com.Content.Text), "ambiguous call must not mutate the document")
        return
    raise AssertionError("expected AmbiguousMatchError for 3-match MULTI_TOKEN replace")


def t_find_replace_occurrence(_word: wl.Word, doc: Document) -> None:
    with doc.edit("E2E: occurrence=1"):
        applied = doc.find_replace("MULTI_TOKEN", "OCC1", occurrence=1)
    expect(len(applied) == 1, f"expected 1 replacement at occurrence=1, got {len(applied)}")
    text = str(doc.com.Content.Text)
    expect("OCC1" in text, "OCC1 missing from doc after occurrence=1 replace")
    expect(text.count("MULTI_TOKEN") == 2, f"expected 2 MULTI_TOKEN left, got {text.count('MULTI_TOKEN')}")


def t_find_replace_all(_word: wl.Word, doc: Document) -> None:
    with doc.edit("E2E: all"):
        applied = doc.find_replace("MULTI_TOKEN", "ALL", all=True)
    expect(len(applied) == 2, f"expected 2 replacements with --all, got {len(applied)}")
    expect("MULTI_TOKEN" not in str(doc.com.Content.Text), "MULTI_TOKEN remained after --all")
    expect(str(doc.com.Content.Text).count("ALL") >= 2, "fewer than 2 'ALL' tokens after --all replace")


def t_find_replace_zero(_word: wl.Word, doc: Document) -> None:
    try:
        with doc.edit("E2E: zero matches"):
            doc.find_replace("ZZZ_DOES_NOT_EXIST_ZZZ", "x")
    except AnchorNotFoundError as e:
        expect(e.kind == "find", f"AnchorNotFoundError for find should set kind='find', got {e.kind!r}")
        return
    raise AssertionError("expected AnchorNotFoundError for zero-match find_replace")


def t_section_scoped_replace(_word: wl.Word, doc: Document) -> None:
    """Replace inside the Risks section only; verify other sections untouched."""
    before_conclusion = doc.heading("Conclusion").section_text()
    with doc.edit("E2E: scoped replace"):
        applied = doc.find_replace(
            "documented",
            "noted",
            scope=doc.heading("Risks"),
            all=True,
        )
    expect(len(applied) >= 1, "expected at least one replacement inside Risks section")
    risks_after = doc.heading("Risks").section_text()
    expect("documented" not in risks_after, "'documented' should be gone from Risks section")
    expect(doc.heading("Conclusion").section_text() == before_conclusion, "Conclusion section must not be mutated by Risks-scoped replace")


def t_selection_preserved(word: wl.Word, doc: Document) -> None:
    # Park the cursor at a known offset, then make an edit elsewhere.
    word.com.Selection.SetRange(0, 0)
    before = word.selection.info()
    with doc.edit("E2E: no-op edit"):
        # touch the bookmark again so the edit is real
        doc.bookmarks[BOOKMARK_NAME].set_text(doc.bookmarks[BOOKMARK_NAME].text)
    after = word.selection.info()
    expect(
        (before["start"], before["end"]) == (after["start"], after["end"]),
        f"Selection drifted: before={before} after={after}",
    )


def t_insert_paragraph_after(_word: wl.Word, doc: Document) -> None:
    # Count headings before / after to confirm the insert landed.
    before = len(doc.outline())
    marker = "Inserted by E2E."
    with doc.edit("E2E: insert paragraph"):
        doc.heading("Conclusion").insert_paragraph_after(marker)
    expect(marker in str(doc.com.Content.Text), "inserted paragraph not visible in doc content")
    # New paragraph is body, not a heading; outline count unchanged.
    after = len(doc.outline())
    expect(before == after, f"outline count changed unexpectedly: {before} -> {after}")


def t_styles_list(_word: wl.Word, doc: Document) -> None:
    rows = doc.styles.list()
    names = [r["name"] for r in rows]
    expect("Heading 1" in names, f"'Heading 1' missing from style list (got {len(names)} styles)")
    expect("Normal" in names, "'Normal' missing from style list")
    paragraph_count = sum(1 for r in rows if r["type"] == "paragraph")
    expect(paragraph_count > 0, "no paragraph-typed styles found")


def t_apply_style_to_bookmark(_word: wl.Word, doc: Document) -> None:
    with doc.edit("E2E: apply style"):
        doc.bookmarks[BOOKMARK_NAME].apply_style("Heading 3")
    # Read the style back off the bookmark's range.
    rng = doc.com.Bookmarks(BOOKMARK_NAME).Range
    applied = str(rng.ParagraphFormat.Style.NameLocal)
    expect(applied == "Heading 3", f"expected style 'Heading 3', got {applied!r}")
    # Restore so later read-tests aren't surprised.
    with doc.edit("E2E: revert style"):
        doc.bookmarks[BOOKMARK_NAME].apply_style("Normal")


def t_format_paragraph(_word: wl.Word, doc: Document) -> None:
    target = doc.heading("Action items")
    with doc.edit("E2E: format paragraph"):
        target.format_paragraph(alignment="center", space_before=6.0, left_indent=18.0)
    # Read back through the paragraph range.
    pf = target.com.ParagraphFormat
    expect(int(pf.Alignment) == 1, f"alignment not centered: got {int(pf.Alignment)}")
    expect(abs(float(pf.SpaceBefore) - 6.0) < 0.01, f"SpaceBefore mismatch: {pf.SpaceBefore}")
    expect(abs(float(pf.LeftIndent) - 18.0) < 0.01, f"LeftIndent mismatch: {pf.LeftIndent}")


def t_apply_style_missing_raises(_word: wl.Word, doc: Document) -> None:
    try:
        doc.bookmarks[BOOKMARK_NAME].apply_style("NoSuchStyleXyz")
    except StyleNotFoundError as exc:
        expect(exc.kind == "style", f"expected kind='style', got {exc.kind!r}")
        expect(exc.name == "NoSuchStyleXyz", f"expected name='NoSuchStyleXyz', got {exc.name!r}")
        # StyleNotFoundError must be catchable as AnchorNotFoundError for backwards compat.
        expect(isinstance(exc, AnchorNotFoundError), "StyleNotFoundError should subclass AnchorNotFoundError")
        return
    raise AssertionError("expected StyleNotFoundError, but no exception was raised")


def t_tables_list(_word: wl.Word, doc: Document) -> None:
    rows = doc.tables.list()
    expect(len(rows) >= 1, f"expected at least one table, got {len(rows)}")
    t = rows[0]
    expect(t["rows"] == 2 and t["columns"] == 2, f"unexpected table size: {t}")


def t_table_read(_word: wl.Word, doc: Document) -> None:
    grid = doc.tables[1].read()
    expect(grid["cells"][0][0]["text"] == "R1C1", f"cell (1,1) mismatch: {grid['cells'][0][0]}")
    expect(grid["cells"][1][1]["anchor_id"] == "table:1:2:2", "cell anchor_id mismatch")
    # Cell text must be stripped of Word's trailing cell markers.
    expect("\x07" not in grid["cells"][0][0]["text"], "cell text leaked the \\x07 cell mark")


def t_cell_set_text(_word: wl.Word, doc: Document) -> None:
    with doc.edit("E2E: set cell"):
        doc.tables[1].cell(1, 2).set_text("UPDATED")
    expect(doc.tables[1].cell(1, 2).text == "UPDATED", "cell set_text round-trip failed")


def t_cell_via_anchor_id(_word: wl.Word, doc: Document) -> None:
    anchor = doc.anchor_by_id("table:1:2:1")
    expect(anchor.kind == "cell", f"expected cell kind, got {anchor.kind!r}")
    with doc.edit("E2E: replace cell via id"):
        anchor.set_text("VIA_ID")
    expect(doc.tables[1].cell(2, 1).text == "VIA_ID", "cell write via anchor_by_id failed")


def t_apply_style_to_cell(_word: wl.Word, doc: Document) -> None:
    with doc.edit("E2E: style cell"):
        doc.tables[1].cell(1, 1).apply_style("Heading 3")
    rng = doc.com.Tables(1).Cell(1, 1).Range
    applied = str(rng.ParagraphFormat.Style.NameLocal)
    expect(applied == "Heading 3", f"expected cell style 'Heading 3', got {applied!r}")


def t_add_row(_word: wl.Word, doc: Document) -> None:
    before = doc.tables[1].row_count
    with doc.edit("E2E: add row"):
        doc.tables[1].add_row(["NR1", "NR2"])
    t = doc.tables[1]
    expect(t.row_count == before + 1, f"row count did not grow: {before} -> {t.row_count}")
    expect(t.cell(t.row_count, 1).text == "NR1", "new row first cell missing value")


def t_delete_row(_word: wl.Word, doc: Document) -> None:
    before = doc.tables[1].row_count
    with doc.edit("E2E: delete row"):
        doc.tables[1].delete_row(before)
    expect(doc.tables[1].row_count == before - 1, "row count did not shrink after delete")


def t_bookmark_in_cell_roundtrip(_word: wl.Word, doc: Document) -> None:
    """The roadmap's open question: do bookmarks inside cells round-trip via set_text?"""
    bm_name = "CellBookmark"
    cell_range = doc.com.Tables(1).Cell(2, 2).Range
    start, end = int(cell_range.Start), int(cell_range.End)
    # Exclude the trailing cell mark from the bookmark range.
    bm_range = doc.com.Range(start, max(start, end - 1))
    doc.com.Bookmarks.Add(Name=bm_name, Range=bm_range)
    with doc.edit("E2E: write bookmark in cell"):
        doc.bookmarks[bm_name].set_text("InCell")
    expect(doc.bookmarks[bm_name].text == "InCell", "bookmark-in-cell set_text round-trip failed")
    expect("InCell" in doc.tables[1].cell(2, 2).text, "cell did not reflect bookmark write")


def t_cli_status(_word: wl.Word, _doc: Document) -> None:
    """Smoke the CLI via subprocess — proves the install entry point works."""
    result = subprocess.run(
        [sys.executable, "-m", "wordlive.cli", "status"],
        capture_output=True,
        text=True,
        timeout=30,
    )
    expect(result.returncode == 0, f"wordlive status exited {result.returncode}: {result.stderr.strip()}")
    payload = json.loads(result.stdout.strip().splitlines()[-1])
    expect(isinstance(payload, list), f"wordlive status JSON should be a list, got {type(payload).__name__}")


# ---------------------------------------------------------------------------
# Orchestration
# ---------------------------------------------------------------------------


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--keep", action="store_true", help="Leave the test document open after the run.")
    parser.add_argument("--verbose", action="store_true", help="Print full tracebacks on failures.")
    parser.add_argument("--no-cli", action="store_true", help="Skip the subprocess CLI smoke test.")
    args = parser.parse_args()

    h = Harness(verbose=args.verbose)

    print("Attaching to Word (launching if necessary)…")
    with wl.connect(launch_if_missing=True, visible=True) as word:
        with temp_doc(word, keep=args.keep) as doc:
            # Make the test doc the active one (Documents.Add already does this,
            # but we re-assert in case the user clicks away mid-run).
            try:
                doc.com.Activate()
            except Exception:
                pass

            # Read-only checks first.
            h.run("status lists test document", lambda: t_status(word, doc))
            h.run("outline returns expected headings", lambda: t_outline(word, doc))
            h.run("anchor_by_id resolves heading", lambda: t_anchor_by_id_heading(word, doc))
            h.run("section_text reads body under heading", lambda: t_section_text(word, doc))
            h.run("find tolerates smart quotes", lambda: t_find_smart_quotes(word, doc))

            # Polite writes.
            h.run("bookmark read + set_text round-trip", lambda: t_bookmark_read_and_write(word, doc))
            h.run("content control read + set_text round-trip", lambda: t_cc_read_and_write(word, doc))
            h.run("selection preserved across edit", lambda: t_selection_preserved(word, doc))
            h.run("insert_paragraph_after lands in body", lambda: t_insert_paragraph_after(word, doc))

            # Fuzzy find/replace matrix. Order matters: ambiguous before occurrence, occurrence before all.
            h.run("find_replace single match", lambda: t_find_replace_single(word, doc))
            h.run("find_replace ambiguous raises AmbiguousMatchError", lambda: t_find_replace_ambiguous(word, doc))
            h.run("find_replace occurrence=1 hits exactly one", lambda: t_find_replace_occurrence(word, doc))
            h.run("find_replace --all hits the rest", lambda: t_find_replace_all(word, doc))
            h.run("find_replace zero matches raises AnchorNotFoundError(find)", lambda: t_find_replace_zero(word, doc))
            h.run("section-scoped replace only touches that section", lambda: t_section_scoped_replace(word, doc))

            # Styles + paragraph formatting.
            h.run("doc.styles.list includes built-ins", lambda: t_styles_list(word, doc))
            h.run("apply_style writes through to the range", lambda: t_apply_style_to_bookmark(word, doc))
            h.run("format_paragraph sets alignment/indent/spacing", lambda: t_format_paragraph(word, doc))
            h.run("apply_style with missing name raises StyleNotFoundError", lambda: t_apply_style_missing_raises(word, doc))

            # Tables.
            h.run("tables.list reports the test table", lambda: t_tables_list(word, doc))
            h.run("table.read returns cells with anchor ids", lambda: t_table_read(word, doc))
            h.run("cell set_text round-trips", lambda: t_cell_set_text(word, doc))
            h.run("cell resolves + writes via anchor_by_id", lambda: t_cell_via_anchor_id(word, doc))
            h.run("bookmark inside a cell round-trips via set_text", lambda: t_bookmark_in_cell_roundtrip(word, doc))
            h.run("apply_style writes through to a cell", lambda: t_apply_style_to_cell(word, doc))
            h.run("table.add_row appends and fills cells", lambda: t_add_row(word, doc))
            h.run("table.delete_row removes a row", lambda: t_delete_row(word, doc))

            if not args.no_cli:
                h.run("CLI: wordlive status (subprocess)", lambda: t_cli_status(word, doc))

    print()
    print(h.summary())
    return 0 if h.failed == 0 else 1


if __name__ == "__main__":
    sys.exit(main())
