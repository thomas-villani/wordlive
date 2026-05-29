#!/usr/bin/env python
"""Append a timestamped note to the end of the open document — politely.

Demonstrates the two core wordlive guarantees:
    * atomic undo  — the whole edit collapses into a single Ctrl-Z.
    * politeness   — your cursor and scroll position are restored on exit.

Safe to run on any document: it only *appends* one paragraph, which you can
remove with a single Ctrl-Z.

Run:
    python append_note.py
    python append_note.py "Custom note text"
"""

from __future__ import annotations

import datetime as _dt
import sys

import wordlive as wl


def main(argv: list[str]) -> int:
    if len(argv) > 1:
        note = argv[1]
    else:
        stamp = _dt.datetime.now().strftime("%Y-%m-%d %H:%M")
        note = f"Note added {stamp} by append_note.py"

    try:
        with wl.attach() as word:
            doc = word.documents.active

            before = doc.selection.info()  # where the user's cursor is
            with doc.edit("Append note (example)"):
                doc.append_paragraph(note)
            after = doc.selection.info()

            print(f"Appended to {doc.name!r}: {note!r}")
            if before["start"] == after["start"]:
                print("Cursor preserved — you weren't moved. One Ctrl-Z undoes this.")
    except wl.WordNotRunningError:
        print("Word isn't running. Open Word and a document, then retry.", file=sys.stderr)
        return 4
    except wl.WordBusyError:
        print("Word is busy (a dialog may be open). Try again.", file=sys.stderr)
        return 3
    except wl.DocumentNotFoundError:
        print("No active document. Open a .docx in Word, then retry.", file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv))
