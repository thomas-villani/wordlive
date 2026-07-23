"""Choreography for the *regularize* demo GIF.

Builds a deliberately messy document in a throwaway window (never your own
document), shows `lint` naming everything that's off, then `regularize` fixing
the safe findings in one atomic-undo pass — so a single Ctrl-Z in Word visibly
reverts the whole cleanup on camera.

Run it while screen-recording the Word window (see README.md in this folder):

    python examples/demos/demo_regularize.py

Word must already be running. The throwaway document is left open and unsaved;
close it without saving when you're done.
"""

import time

import wordlive as wl

PAUSE = 3.0  # seconds between beats, so the recording has time to breathe


def beat(msg: str) -> None:
    print(msg, flush=True)
    time.sleep(PAUSE)


def main() -> None:
    with wl.attach() as word:
        word.com.Documents.Add()  # throwaway — close without saving afterwards
        doc = word.documents.active

        beat("Building a deliberately messy draft…")
        with doc.edit("Demo: messy draft"):
            doc.end.insert_markdown(
                "# Quarterly Report\n\n"
                "Revenue grew 12% over the prior quarter , driven by services .\n\n"
                "Costs held flat ; headcount was unchanged .\n\n"
                "Outlook remains positive for Q4 .\n"
            )
        # Hand-mangle formatting the way a rushed human would — direct font
        # overrides fighting the applied style (`.com` is the anchor's raw
        # COM Range, the escape hatch):
        with doc.edit("Demo: hand-mangled formatting"):
            for para in doc.paragraphs:
                text = para.com.Text
                if text.startswith("Costs held flat"):
                    para.com.Font.Name = "Arial"
                    para.com.Font.Size = 14
                elif text.startswith("Outlook remains"):
                    para.com.Font.Bold = True

        beat("\nlint — what's off about this document?")
        for f in doc.lint():
            print(f"  {f['severity']:>7}  {f['rule']:<32} {f['anchor_id']}")

        beat("\nregularize — apply the safe fixes, one atomic undo…")
        report = doc.regularize()
        print(f"\n{report['ops_run']} fixes applied.")
        print("Now press Ctrl-Z in Word: the entire cleanup reverts in one step.")


if __name__ == "__main__":
    main()
