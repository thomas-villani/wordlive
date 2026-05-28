"""snapshot — page resolution, PDF-export args, file naming, and errors.

The PDF rasteriser (PyMuPDF) is the one piece that needs a heavy optional
dependency and a real PDF on disk, so it's isolated behind
`_snapshot._rasterize_pdf` and faked here. That lets these tests cover all of
wordlive's own logic — which pages get exported, how they're numbered, where
files land — against the mock Word, with no PyMuPDF install. One test at the
bottom exercises the real rasteriser when PyMuPDF *is* available.
"""

from __future__ import annotations

import base64
import json

import pytest
from click.testing import CliRunner

import wordlive
from wordlive import _snapshot
from wordlive.cli.main import EXIT_ANCHOR_NOT_FOUND, EXIT_OK, EXIT_OTHER, main
from wordlive.constants import WdExportFormat, WdExportRange
from wordlive.exceptions import SnapshotError


@pytest.fixture
def fake_pages(monkeypatch):
    """Make `_rasterize_pdf` return a controllable list of fake PNGs.

    Returns a setter: call `fake_pages(n)` to have the next render yield `n`
    pages (`b"PNG0"`, `b"PNG1"`, …) and to capture the args it was called with.
    """
    state: dict = {"pngs": [b"PNG0"], "calls": []}

    def fake_rasterize(pdf_path, dpi):
        state["calls"].append({"pdf_path": pdf_path, "dpi": dpi})
        return list(state["pngs"])

    monkeypatch.setattr(_snapshot, "_rasterize_pdf", fake_rasterize)

    def configure(n):
        state["pngs"] = [f"PNG{i}".encode() for i in range(n)]
        return state

    configure.state = state  # type: ignore[attr-defined]
    return configure


def _export_kwargs(fake_word):
    fake_word.ActiveDocument.ExportAsFixedFormat.assert_called_once()
    return fake_word.ActiveDocument.ExportAsFixedFormat.call_args.kwargs


# ---------------------------------------------------------------------------
# Library: page selection -> export args + page numbering
# ---------------------------------------------------------------------------


def test_snapshot_all_pages_exports_whole_document(fake_word, fake_pages):
    state = fake_pages(2)
    with wordlive.attach() as word:
        shots = word.documents.active.snapshot()
    kwargs = _export_kwargs(fake_word)
    assert kwargs["ExportFormat"] == int(WdExportFormat.PDF)
    assert kwargs["Range"] == int(WdExportRange.ALL_DOCUMENT)
    assert "From" not in kwargs and "To" not in kwargs
    # Whole-document render numbers pages from 1.
    assert [s.page for s in shots] == [1, 2]
    assert [s.png for s in shots] == state["pngs"]
    assert all(s.path is None for s in shots)


def test_snapshot_single_page(fake_word, fake_pages):
    fake_pages(1)
    with wordlive.attach() as word:
        shots = word.documents.active.snapshot(pages=3)
    kwargs = _export_kwargs(fake_word)
    assert kwargs["Range"] == int(WdExportRange.FROM_TO)
    assert (kwargs["From"], kwargs["To"]) == (3, 3)
    assert [s.page for s in shots] == [3]


def test_snapshot_page_span_numbers_from_start(fake_word, fake_pages):
    fake_pages(3)
    with wordlive.attach() as word:
        shots = word.documents.active.snapshot(pages=(2, 4))
    kwargs = _export_kwargs(fake_word)
    assert (kwargs["From"], kwargs["To"]) == (2, 4)
    # The exported PDF holds only pages 2-4, so they map back to 2, 3, 4.
    assert [s.page for s in shots] == [2, 3, 4]


def test_snapshot_forwards_dpi(fake_word, fake_pages):
    state = fake_pages(1)
    with wordlive.attach() as word:
        word.documents.active.snapshot(pages=1, dpi=222)
    assert state["calls"][0]["dpi"] == 222


def test_snapshot_default_dpi_is_150(fake_word, fake_pages):
    state = fake_pages(1)
    with wordlive.attach() as word:
        word.documents.active.snapshot(pages=1)
    assert state["calls"][0]["dpi"] == 150


@pytest.mark.parametrize("bad", [0, -1, (0, 2), (3, 2), (1, 2, 3)])
def test_snapshot_rejects_bad_pages(fake_word, fake_pages, bad):
    fake_pages(1)
    with wordlive.attach() as word:
        with pytest.raises(ValueError):
            word.documents.active.snapshot(pages=bad)


# ---------------------------------------------------------------------------
# Library: writing files via out=
# ---------------------------------------------------------------------------


def test_snapshot_out_single_page_writes_exact_path(fake_word, fake_pages, tmp_path):
    state = fake_pages(1)
    out = tmp_path / "section.png"
    with wordlive.attach() as word:
        shots = word.documents.active.snapshot(out, pages=3)
    assert out.read_bytes() == state["pngs"][0]
    assert shots[0].path == out


def test_snapshot_out_multi_page_suffixes_page_number(fake_word, fake_pages, tmp_path):
    state = fake_pages(2)
    out = tmp_path / "doc.png"
    with wordlive.attach() as word:
        shots = word.documents.active.snapshot(out, pages=(4, 5))
    assert (tmp_path / "doc-p4.png").read_bytes() == state["pngs"][0]
    assert (tmp_path / "doc-p5.png").read_bytes() == state["pngs"][1]
    assert [s.path.name for s in shots] == ["doc-p4.png", "doc-p5.png"]
    # The original bare path is not written when there's more than one page.
    assert not out.exists()


# ---------------------------------------------------------------------------
# Library: anchor snapshots
# ---------------------------------------------------------------------------


def test_anchor_snapshot_uses_anchor_page_span(fake_word, fake_pages):
    fake_pages(1)
    with wordlive.attach() as word:
        doc = word.documents.active
        shots = doc.bookmarks["Address"].snapshot()
    kwargs = _export_kwargs(fake_word)
    # The fake document is a single page, so the anchor resolves to page 1.
    assert kwargs["Range"] == int(WdExportRange.FROM_TO)
    assert (kwargs["From"], kwargs["To"]) == (1, 1)
    assert [s.page for s in shots] == [1]


def test_heading_snapshot_renders_its_section(fake_word, fake_pages):
    """A heading anchor expands to its section span (start of heading -> body end)."""
    fake_pages(1)
    with wordlive.attach() as word:
        doc = word.documents.active
        shots = doc.heading("Introduction").snapshot()
    kwargs = _export_kwargs(fake_word)
    assert (kwargs["From"], kwargs["To"]) == (1, 1)
    assert len(shots) == 1


def test_anchor_snapshot_writes_file(fake_word, fake_pages, tmp_path):
    state = fake_pages(1)
    out = tmp_path / "anchor.png"
    with wordlive.attach() as word:
        doc = word.documents.active
        shots = doc.bookmarks["Address"].snapshot(out)
    assert out.read_bytes() == state["pngs"][0]
    assert shots[0].path == out


# ---------------------------------------------------------------------------
# Library: missing PyMuPDF -> SnapshotError
# ---------------------------------------------------------------------------


def test_snapshot_without_pymupdf_raises(fake_word, monkeypatch):
    def boom():
        raise SnapshotError("PyMuPDF missing")

    # Don't fake _rasterize_pdf here — let the real one run and hit the import.
    monkeypatch.setattr(_snapshot, "_import_fitz", boom)
    with wordlive.attach() as word:
        with pytest.raises(SnapshotError):
            word.documents.active.snapshot(pages=1)


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------


def _invoke(args):
    result = CliRunner().invoke(main, args, catch_exceptions=False)
    return result.exit_code, result.stdout, result.stderr


def test_cli_snapshot_base64_inline(fake_word, fake_pages):
    state = fake_pages(1)
    code, out, _ = _invoke(["snapshot", "--page", "1"])
    assert code == EXIT_OK
    data = json.loads(out)
    assert data["ok"] is True and data["count"] == 1
    img = data["images"][0]
    assert img["page"] == 1
    assert base64.b64decode(img["base64"]) == state["pngs"][0]
    assert "path" not in img


def test_cli_snapshot_out_writes_file(fake_word, fake_pages, tmp_path):
    state = fake_pages(1)
    out = tmp_path / "s.png"
    code, stdout, _ = _invoke(["snapshot", "--anchor-id", "heading:1", "--out", str(out)])
    assert code == EXIT_OK
    assert out.read_bytes() == state["pngs"][0]
    img = json.loads(stdout)["images"][0]
    assert img["path"] == str(out)
    assert "base64" not in img


def test_cli_snapshot_pages_range(fake_word, fake_pages):
    fake_pages(3)
    code, out, _ = _invoke(["snapshot", "--pages", "2-4"])
    assert code == EXIT_OK
    data = json.loads(out)
    assert data["count"] == 3
    assert [im["page"] for im in data["images"]] == [2, 3, 4]


def test_cli_snapshot_bad_pages_is_usage_error(fake_word, fake_pages):
    fake_pages(1)
    code, _, err = _invoke(["snapshot", "--pages", "notarange"])
    assert code != EXIT_OK
    assert "--pages" in err


def test_cli_snapshot_two_targets_is_usage_error(fake_word, fake_pages):
    fake_pages(1)
    code, _, err = _invoke(["snapshot", "--page", "1", "--anchor-id", "heading:1"])
    assert code != EXIT_OK
    assert "at most one" in err.lower()


def test_cli_snapshot_missing_anchor(fake_word, fake_pages):
    fake_pages(1)
    code, _, err = _invoke(["snapshot", "--anchor-id", "heading:99"])
    assert code == EXIT_ANCHOR_NOT_FOUND
    assert "heading" in err.lower()


def test_cli_snapshot_without_pymupdf_exit_other(fake_word, monkeypatch):
    monkeypatch.setattr(
        _snapshot, "_import_fitz", lambda: (_ for _ in ()).throw(SnapshotError("missing"))
    )
    code, _, err = _invoke(["snapshot", "--page", "1"])
    assert code == EXIT_OTHER
    assert "error" in err.lower()


# ---------------------------------------------------------------------------
# Real rasteriser — only when PyMuPDF is actually installed
# ---------------------------------------------------------------------------


def test_rasterize_pdf_real(tmp_path):
    fitz = pytest.importorskip("pymupdf")
    pdf_path = tmp_path / "two.pdf"
    doc = fitz.open()
    doc.new_page()
    doc.new_page()
    doc.save(str(pdf_path))
    doc.close()

    pngs = _snapshot._rasterize_pdf(str(pdf_path), dpi=72)
    assert len(pngs) == 2
    assert all(p.startswith(b"\x89PNG\r\n\x1a\n") for p in pngs)
