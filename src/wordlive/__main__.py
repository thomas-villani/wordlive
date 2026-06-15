"""Run the CLI as `python -m wordlive` — an alias for the `wordlive` console script.

Lets tooling (and the e2e test suite) drive the CLI through the *current*
interpreter without depending on the console-script being on PATH.
"""

from __future__ import annotations

from .cli.main import main

if __name__ == "__main__":
    main()
