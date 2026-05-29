#!/usr/bin/env python
"""Print the outline of the Word document you have open right now.

Read-only — it never changes the document or moves your cursor. The best first
script to confirm wordlive can see your document.

Prerequisites:
    * Windows, with Microsoft Word running and a document open.
    * pip install wordlive   (or: uv add wordlive)

Run:
    python read_outline.py
    uv run python read_outline.py
"""

from __future__ import annotations

import sys

import wordlive as wl


def main() -> int:
    try:
        with wl.attach() as word:
            doc = word.documents.active
            print(f"Active document: {doc.name}\n")

            outline = doc.outline()
            if not outline:
                print("(no headings — every paragraph is still a para:N anchor)")
            for entry in outline:
                indent = "  " * (entry["level"] - 1)
                print(f"{indent}{entry['text']}  [{entry['anchor_id']}]")

            paras = doc.paragraphs.list()
            print(f"\n{len(paras)} paragraph(s), {len(outline)} heading(s).")
    except wl.WordNotRunningError:
        print("Word isn't running. Open Word and a document, then retry.", file=sys.stderr)
        return 4
    except wl.DocumentNotFoundError:
        print("No active document. Open a .docx in Word, then retry.", file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
