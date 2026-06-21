"""Smoke tests for `doc.to_markdown` — require a real running Word.

Skipped by default (`pytest -m "not smoke"`). Run with `pytest -m smoke` on a
Windows box with Word installed. These validate the COM document-walk that the
fake fixture can't model (per-word emphasis, ListFormat, table interleave) and
the headline contract: a constrained-subset `insert_markdown` round-trips back
through `to_markdown`.
"""

from __future__ import annotations

import base64
import contextlib

import pytest

from wordlive._markdown import parse_markdown

pytestmark = pytest.mark.smoke

# A 1x1 transparent PNG for the inline-image check.
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


def _kinds(md: str) -> list[tuple[str, int | None]]:
    """The block kind/level sequence of a Markdown string (drops empty blocks)."""
    return [(b.kind, b.level) for b in parse_markdown(md)]


def test_to_markdown_round_trips_constrained_subset(scratch_doc):
    """insert_markdown(X) then to_markdown() recovers X's block/run structure."""
    doc = scratch_doc
    md_in = (
        "# Title\n\n"
        "Body with **bold** and *italic* and ***both*** words.\n\n"
        "- one\n"
        "- two\n\n"
        "1. first\n"
        "2. second\n\n"
        "## Subhead\n\n"
        "Closing paragraph."
    )
    with doc.edit("smoke: insert markdown"):
        doc.anchor_by_id("start").insert_markdown(md_in)

    out = doc.to_markdown()
    # The numbered list renders as "1." for every item (GFM renumbers), so compare
    # the parsed kind/level sequence, not bytes.
    assert _kinds(out) == _kinds(md_in)
    # Emphasis survives the round-trip through Range.Words coalescing.
    assert "**bold**" in out
    assert "*italic*" in out
    assert "***both***" in out


def test_to_markdown_emits_gfm_table(scratch_doc):
    doc = scratch_doc
    with doc.edit("smoke: table"):
        doc.end.insert_table(data=[{"Name": "Widget", "Qty": "3"}], header=True)

    out = doc.to_markdown()
    assert "| Name | Qty |" in out
    assert "| Widget | 3 |" in out
    # The alignment separator row is present (undirected or directed).
    assert "\n| " in out and "--" in out


def test_to_markdown_inline_image(scratch_doc, tmp_path):
    doc = scratch_doc
    png = tmp_path / "pic.png"
    png.write_bytes(_PNG)
    with doc.edit("smoke: image"):
        doc.end.insert_image(str(png), wrap="inline", alt_text="logo")

    out = doc.to_markdown()
    assert "![logo](image:1)" in out


def test_to_markdown_hyperlink(scratch_doc):
    doc = scratch_doc
    with doc.edit("smoke: link"):
        doc.anchor_by_id("start").link_to("https://example.test", text="Example")

    out = doc.to_markdown()
    assert "[Example](https://example.test" in out


def test_to_html_renders_headings_lists_emphasis(scratch_doc):
    doc = scratch_doc
    with doc.edit("smoke: html"):
        doc.anchor_by_id("start").insert_markdown("# Title\n\nBody **bold**.\n\n- one\n- two")

    html = doc.to_html()
    assert "<h1>Title</h1>" in html
    assert "<strong>bold</strong>" in html
    assert "<ul>" in html and "<li>one</li>" in html


def test_to_markdown_within_scopes_to_anchor(scratch_doc):
    """within=heading:N emits only that heading line (literal-range semantics)."""
    doc = scratch_doc
    with doc.edit("smoke: scope"):
        doc.anchor_by_id("start").insert_markdown("# Alpha\n\nbody a\n\n# Beta\n\nbody b")

    # The whole doc has both headings; a single heading's range is just its line.
    whole = doc.to_markdown()
    assert "# Alpha" in whole and "# Beta" in whole
    alpha = doc.to_markdown(within="heading:1")
    assert "Alpha" in alpha
    assert "Beta" not in alpha


def _multi_section_markdown(sections: int = 6, para_words: int = 60) -> str:
    body = " ".join(f"word{i}" for i in range(para_words))
    parts = []
    for s in range(1, sections + 1):
        parts.append(f"# Section {s}")
        parts.append(body)
        parts.append(f"## Subsection {s}.1")
        parts.append(body)
    return "\n\n".join(parts)


def test_read_digest_keeps_every_heading_under_budget(scratch_doc, capsys):
    """doc.read(budget) on a multi-section doc: every heading verbatim, under budget."""
    doc = scratch_doc
    with doc.edit("smoke: digest"):
        doc.anchor_by_id("start").insert_markdown(_multi_section_markdown())

    digest = doc.read(budget=400)
    # Every heading is verbatim (the navigation spine).
    for s in range(1, 7):
        assert f"# Section {s}" in digest
        assert f"## Subsection {s}.1" in digest
    # Body was sampled, not dumped whole — elision/truncation markers appear.
    assert "…(" in digest and ("words elided" in digest or "more words" in digest)
    # The budget bounds the output (chars/4 token estimate, generous slack for the
    # heading spine + per-section lead snippets the budget can't compress away).
    assert len(digest) // 4 <= 400 * 2

    with capsys.disabled():
        print(
            "\n----- doc.read(budget=400) -----\n" + digest + "\n--------------------------------"
        )
