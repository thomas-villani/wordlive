#!/usr/bin/env python
"""Render a page of the open document to a PNG so a vision model can *see* it.

Word exports a pixel-faithful PDF of the live document and wordlive rasterises
the page you ask for — real fonts, spacing, and page geometry, not just text.
Read-only: the document and your cursor are untouched.

Requires the optional ``snapshot`` extra:
    pip install "wordlive[snapshot]"     (or: uv add "wordlive[snapshot]")

Run:
    python snapshot_page.py                 # page 1 -> page.png
    python snapshot_page.py 2 out.png       # page 2 -> out.png
"""

from __future__ import annotations

import sys

import wordlive as wl


def main(argv: list[str]) -> int:
    page = int(argv[1]) if len(argv) > 1 else 1
    out = argv[2] if len(argv) > 2 else "page.png"

    try:
        with wl.attach() as word:
            doc = word.documents.active
            shots = doc.snapshot(out, pages=page)
            for shot in shots:
                print(f"Rendered page {shot.page} -> {shot.path} ({len(shot.png):,} bytes)")
    except wl.SnapshotError as exc:
        print(
            f'Snapshot failed: {exc}\nInstall the extra: pip install "wordlive[snapshot]"',
            file=sys.stderr,
        )
        return 1
    except wl.WordNotRunningError:
        print("Word isn't running. Open Word and a document, then retry.", file=sys.stderr)
        return 4
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv))
