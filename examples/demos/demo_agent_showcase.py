"""Choreography for the *agent drives Word* demo GIF.

Plays the part of an LLM agent building a styled document from nothing in a
throwaway window: headings and body land section by section, a data table
appears under a heading, and a review comment is left for the human — all
politely (the cursor never moves) and each step one Ctrl-Z.

Run it while screen-recording the Word window (see README.md in this folder):

    python examples/demos/demo_agent_showcase.py

Word must already be running. The throwaway document is left open and unsaved;
close it without saving when you're done.
"""

import time

import wordlive as wl

PAUSE = 2.5  # seconds between beats, so the recording has time to breathe


def beat(msg: str) -> None:
    print(msg, flush=True)
    time.sleep(PAUSE)


def main() -> None:
    with wl.attach() as word:
        word.com.Documents.Add()  # throwaway — close without saving afterwards
        doc = word.documents.active

        beat("Agent: drafting the title and framing…")
        with doc.edit("Draft title and framing"):
            doc.end.insert_markdown(
                "# Project Aurora — Kickoff Brief\n\n"
                "This brief outlines the scope, timeline, and budget for "
                "Project Aurora, prepared for the steering committee.\n"
            )

        beat("Agent: adding the scope section…")
        with doc.edit("Add scope section"):
            doc.end.insert_markdown(
                "## Scope\n\n"
                "- Migrate the reporting pipeline to the new platform\n"
                "- Retire the legacy exports by end of quarter\n"
                "- **Out of scope:** the archival warehouse\n"
            )

        beat("Agent: adding the budget section with a table…")
        with doc.edit("Add budget table"):
            doc.end.insert_markdown("## Budget\n")
            doc.heading("Budget").insert_table(
                data=[
                    ["Item", "Cost"],
                    ["Platform licences", "$40,000"],
                    ["Contract engineering", "$85,000"],
                    ["Contingency", "$12,500"],
                ],
                header=True,
            )

        beat("Agent: leaving a comment for the human reviewer…")
        with doc.edit("Comment for reviewer"):
            doc.comments.add(
                doc.heading("Budget"),
                "Contingency is set at 10% — flag if the committee wants 15%.",
                author="wordlive agent",
            )

        beat("\nDone — a styled brief, built politely:")
        print("  · your cursor never moved")
        print("  · each step above is exactly one Ctrl-Z in Word")
        print("  · the comment sits in the margin, ready for a human reply")


if __name__ == "__main__":
    main()
