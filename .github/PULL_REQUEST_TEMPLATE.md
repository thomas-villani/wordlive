<!-- See CONTRIBUTING.md for the full guide. Keep one logical change per PR. -->

## What & why

<!-- What does this change, and why? Link any issue (e.g. Fixes #123). -->

## The four surfaces

wordlive exposes most capabilities four ways that must stay in sync. Tick what
this PR touches — and confirm the rest genuinely don't need it:

- [ ] Python API
- [ ] CLI verb (`cli/commands.py`)
- [ ] `exec` op (`_ops.py`)
- [ ] MCP command (`mcp/server.py`)
- [ ] N/A — not a user-facing capability change

## Checklist

- [ ] `uv run ruff format .` and `uv run ruff check .` pass
- [ ] `uv run mypy` passes
- [ ] `uv run pytest` passes; added/updated tests for the change
- [ ] Smoke-tested against real Word on Windows (or N/A — no COM behaviour changed)
- [ ] Updated the docs that track this change (`cli.md` / `mcp.md` / `python-api.md` / `README.md` / `SKILL.md`)
- [ ] Added a `CHANGELOG.md` entry under `[Unreleased]`
