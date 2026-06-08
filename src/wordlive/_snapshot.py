"""Render document pages to PNG — let a vision model *see* the live document.

The pipeline is two steps. Word's COM already renders itself perfectly, so we
ask it to export a pixel-faithful PDF (`Document.ExportAsFixedFormat`), then
rasterise the wanted page(s) to PNG with PyMuPDF. That gives a true WYSIWYG
image of the page — real fonts, layout, and formatting — which is exactly what
a model needs to judge style choices.

PyMuPDF is an optional dependency (the `snapshot` extra): the import is lazy and
deferred to render time, so the rest of wordlive works without it and a missing
backend surfaces as a clean `SnapshotError`.
"""

from __future__ import annotations

import os
import tempfile
from dataclasses import dataclass
from pathlib import Path
from typing import TYPE_CHECKING, Any

from . import _com
from .constants import WdExportFormat, WdExportItem, WdExportRange
from .exceptions import SnapshotError

if TYPE_CHECKING:
    from collections.abc import Sequence


@dataclass(frozen=True)
class Snapshot:
    """One rendered page of a document.

    `page` is the 1-based document page number; `png` is the PNG-encoded image
    bytes — feed it straight to a vision model, or write it yourself. `path` is
    where the image was written when a `snapshot(out=...)` call saved it to disk,
    otherwise `None`.
    """

    page: int
    png: bytes
    path: Path | None = None


def _import_fitz() -> Any:
    """Return the PyMuPDF module, or raise `SnapshotError` if it isn't installed.

    PyMuPDF imports as `pymupdf` (modern) or `fitz` (legacy); accept either.
    """
    try:
        import pymupdf  # type: ignore[import-not-found]

        return pymupdf
    except ImportError:
        pass
    try:
        import fitz  # type: ignore[import-not-found]

        return fitz
    except ImportError as e:
        raise SnapshotError(
            "rendering snapshots requires PyMuPDF; install the extra with "
            '`pip install "wordlive[snapshot]"` (or `uv add "wordlive[snapshot]"`)'
        ) from e


def _rasterize_pdf(pdf_path: str, dpi: int) -> list[bytes]:
    """Rasterise every page of `pdf_path` to PNG bytes at `dpi`.

    Isolated as its own function so tests can substitute a fake renderer without
    a real PyMuPDF install or a real PDF on disk.
    """
    fitz = _import_fitz()
    pages: list[bytes] = []
    try:
        with fitz.open(pdf_path) as pdf:
            for page in pdf:
                pages.append(page.get_pixmap(dpi=dpi).tobytes("png"))
    except SnapshotError:
        raise
    except Exception as e:  # noqa: BLE001 — any PyMuPDF failure is a render failure
        raise SnapshotError(f"failed to rasterise the exported PDF: {e}") from e
    return pages


def _export_pdf(
    doc_com: Any,
    out_path: str,
    *,
    from_page: int | None,
    to_page: int | None,
    markup: bool = False,
) -> None:
    """Export the document (or a page span) to a PDF at `out_path` via COM.

    `from_page`/`to_page` are 1-based, inclusive; pass both as `None` to export
    the whole document. `markup=True` exports with tracked-change marks and
    comment balloons visible (`Item=wdExportDocumentWithMarkup`) — the export
    parameter, not a view change, so the user's on-screen markup mode is left
    untouched.
    """
    kwargs: dict[str, Any] = {
        "OutputFileName": out_path,
        "ExportFormat": int(WdExportFormat.PDF),
        "OpenAfterExport": False,
        "Item": int(WdExportItem.WITH_MARKUP if markup else WdExportItem.DOCUMENT_CONTENT),
    }
    if from_page is None:
        kwargs["Range"] = int(WdExportRange.ALL_DOCUMENT)
    else:
        kwargs["Range"] = int(WdExportRange.FROM_TO)
        kwargs["From"] = int(from_page)
        kwargs["To"] = int(to_page if to_page is not None else from_page)
    with _com.translate_com_errors():
        doc_com.ExportAsFixedFormat(**kwargs)


def render(
    doc_com: Any,
    *,
    from_page: int | None = None,
    to_page: int | None = None,
    dpi: int = 150,
    markup: bool = False,
) -> list[tuple[int, bytes]]:
    """Render `[from_page, to_page]` (or the whole doc) to `(page_number, png)` pairs.

    Exports to a temporary PDF, rasterises it, and removes the temp file. Page
    numbers are the document's own 1-based numbers: when a span is exported the
    PDF holds only those pages, so the first rasterised page is `from_page`.
    `markup=True` renders tracked changes and comments as visible marks.
    """
    # mkstemp creates the file and returns an open fd; close it immediately so
    # Word can write the PDF to the path without a sharing conflict on Windows.
    fd, pdf_path = tempfile.mkstemp(suffix=".pdf", prefix="wordlive_snap_")
    os.close(fd)
    try:
        _export_pdf(doc_com, pdf_path, from_page=from_page, to_page=to_page, markup=markup)
        pngs = _rasterize_pdf(pdf_path, dpi)
    finally:
        try:
            os.unlink(pdf_path)
        except OSError:
            pass
    base = from_page if from_page is not None else 1
    return [(base + i, png) for i, png in enumerate(pngs)]


def build_snapshots(
    rendered: Sequence[tuple[int, bytes]], out: str | Path | None
) -> list[Snapshot]:
    """Wrap `(page, png)` pairs as `Snapshot`s, writing files when `out` is given.

    A single page writes straight to `out`. Multiple pages can't share one path,
    so each is written next to `out` as `<stem>-p<N><suffix>` (N = page number).
    """
    if out is None:
        return [Snapshot(page=page, png=png, path=None) for page, png in rendered]

    out_path = Path(out)
    single = len(rendered) == 1
    snaps: list[Snapshot] = []
    for page, png in rendered:
        dest = (
            out_path if single else out_path.with_name(f"{out_path.stem}-p{page}{out_path.suffix}")
        )
        dest.write_bytes(png)
        snaps.append(Snapshot(page=page, png=png, path=dest))
    return snaps
