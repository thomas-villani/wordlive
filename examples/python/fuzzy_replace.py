#!/usr/bin/env python
"""Fuzzy find-and-replace, recorded as tracked changes.

The classic LLM editing flow: swap one phrase for another, tolerant of the
cosmetic drift (smart quotes, NBSPs, em-dashes) that creeps in when text is
round-tripped through a model. Wrapped in ``tracked_changes()`` so every edit
lands as an accept/reject-able revision — nothing is silently overwritten, and
the user stays in control.

Run:
    python fuzzy_replace.py "old phrase" "new phrase"
    python fuzzy_replace.py "utilise" "use" --all
"""

from __future__ import annotations

import sys

import wordlive as wl


def main(argv: list[str]) -> int:
    replace_all = "--all" in argv
    positional = [a for a in argv[1:] if a != "--all"]
    if len(positional) != 2:
        print('usage: python fuzzy_replace.py "old phrase" "new phrase" [--all]', file=sys.stderr)
        return 1
    find, replace = positional

    try:
        with wl.attach() as word:
            doc = word.documents.active
            with doc.tracked_changes(), doc.edit("Fuzzy replace (example)"):
                applied = doc.find_replace(find, replace, all=replace_all)
            print(f"Replaced {len(applied)} occurrence(s) of {find!r} as tracked changes.")
    except wl.AmbiguousMatchError as exc:
        print(
            f"{find!r} matched {len(exc.matches)} places — re-run with --all, or narrow it.",
            file=sys.stderr,
        )
        return 5
    except wl.AnchorNotFoundError:
        print(f"{find!r} wasn't found in the document.", file=sys.stderr)
        return 2
    except wl.WordNotRunningError:
        print("Word isn't running. Open Word and a document, then retry.", file=sys.stderr)
        return 4
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv))
