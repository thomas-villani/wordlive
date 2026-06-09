# Contributing to wordlive

Thanks for considering a contribution. wordlive drives a *running* Microsoft
Word instance over COM (`pywin32`) — "xlwings, but for Word" — and is built to be
first-class for LLM tool use. This guide covers the dev setup, the conventions
that keep the codebase coherent, and how to get a change merged.

## Ground rules before you start

- **Windows + Word for the full story.** The library only runs on Windows
  against a real Word install. Most of the test suite uses a fake-COM fixture and
  runs anywhere, but anything that actually touches Word is marked `smoke` and
  needs Word running (see [Testing](#testing)).
- **Open an issue first for anything non-trivial.** A quick "here's what I want
  to do" avoids building something that cuts against the design (below). Bug
  fixes and docs typos can go straight to a PR.
- **No direct commits to `main`.** Branch, then open a PR via `gh` (the project
  has been PR-only since the PyPI release).

## Development setup

wordlive is [`uv`](https://docs.astral.sh/uv/)-managed. The project `.venv` has
no `pip` — use `uv`, never `pip` / `python -m build`.

```bash
git clone https://github.com/thomas-villani/wordlive
cd wordlive
uv sync --extra dev          # dev tools (pytest, ruff, mypy)
# add --extra docs / --extra mcp / --extra snapshot if you're touching those
```

Run anything through uv: `uv run pytest`, `uv run ruff check .`, etc.

### Pre-commit hook (recommended)

A pre-commit hook keeps `uv.lock` in sync with `pyproject.toml`. It does **not**
install itself — once per clone:

```bash
uvx pre-commit install
```

Thereafter every `git commit` re-locks if needed; `uvx pre-commit run -a` runs it
over the whole tree on demand.

## The design — don't break these

These four invariants drive every API decision. A change that violates one is
almost certainly wrong; if you think yours is the exception, say so in the PR.

1. **Politeness first.** Operations preserve the user's `Selection`, view, and
   scroll. `doc.edit()` snapshots and restores selection + scroll on exit. Only
   `go_to()` and `scope.allow_cursor_move()` may move the user — and they say so.
2. **Semantic anchors over `Selection`.** Operations target named handles
   (bookmarks, content controls, headings, paragraphs, cells, ranges,
   headers/footers, `start`/`end`) — never the live cursor unless explicitly
   asked. Anchor ids (`heading:N`, `para:N`, `bookmark:NAME`, `table:N:R:C`, …)
   are the stable, LLM-visible addressing scheme.
3. **Atomic undo.** Every `doc.edit("label")` opens a Word `UndoRecord`, so one
   Ctrl-Z reverts the whole intent. A multi-op `exec` batch is one undo step.
4. **Structured I/O.** Reads return dataclasses/dicts; the CLI emits exactly one
   JSON object per invocation; exit codes are deterministic. No string scraping.

Underlying all four is the **`.com` escape hatch**: every wrapper exposes the raw
pywin32 object via `.com`. Prefer extending the high-level API, but never block a
user from dropping to COM.

### The four surfaces must agree

The same capability is exposed four ways, and they have to stay parallel. Adding
or changing a feature usually means touching **all** of:

- the **Python API** (the wrapper class / method),
- the **CLI** verb (`src/wordlive/cli/commands.py`),
- the **`exec` op** in the batch vocabulary (`src/wordlive/_ops.py`),
- the **MCP** dispatch command (`src/wordlive/mcp/server.py` — one of the four
  `word_read` / `word_write` / `word_exec` / `word_snapshot` tools; prefer a new
  `command` value over a new tool).

And the docs that track code, in the same PR:
`docs/cli.md`, `docs/mcp.md`, `docs/python-api.md`, `README.md`,
`src/wordlive/_skill/*/SKILL.md`, and `CHANGELOG.md`.

See [`CLAUDE.md`](CLAUDE.md) for the repo layout and the longer-form version of
all this.

## Testing

```bash
uv run pytest                # unit tests — fake-COM fixture, run anywhere
uv run pytest -m smoke       # integration — needs a real Word install on Windows
```

- Unit tests use the `fake_word` COM fixture (`tests/conftest.py`), so they run
  off-Windows and in CI on Linux. Add unit coverage for any logic that can be
  exercised without Word (parsing, validation, op dispatch, exit-code mapping).
- `smoke` tests attach to a live Word and assert real behaviour. If your change
  touches COM interaction, add a smoke test and run it on Windows before opening
  the PR.

## Lint, format, types

CI runs all three; run them locally first:

```bash
uv run ruff format .         # formatter (owns line wrapping / E501)
uv run ruff check .          # lint: E/F/W/I/UP/B
uv run mypy                  # type-check src/wordlive (py310 target)
```

If you touch the docs, also build them strictly — the docs CI fails on any
warning:

```bash
uv run --no-sync mkdocs build --strict   # needs --extra docs
```

## Commits & pull requests

- **Conventional-commit style** subject lines: `feat:`, `fix:`, `docs:`,
  `chore:`, `refactor:`, `test:`. Keep the subject imperative and scoped.
- **Update `CHANGELOG.md`** under `## [Unreleased]`, in the right
  [Keep a Changelog](https://keepachangelog.com/) section (Added / Changed /
  Fixed / Security / …). The project follows [SemVer](https://semver.org/).
- **One logical change per PR.** Include the *why*, not just the *what*, and note
  how you verified it (which checks you ran, smoke-tested on Windows or not).
- Make sure `ruff format --check`, `ruff check`, `mypy`, and `pytest` all pass.

Releases are cut by a maintainer with `bump-my-version` (which bumps
`pyproject.toml`, commits `chore: release vX.Y.Z`, and tags it; the tag push
triggers PyPI trusted-publishing). You don't need to bump versions in a PR.

## Reporting security issues

Please **don't** open a public issue for a vulnerability — see
[`SECURITY.md`](SECURITY.md) for private reporting.

## License

By contributing, you agree that your contributions are licensed under the
project's [MIT License](LICENSE).
