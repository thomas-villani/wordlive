# CLAUDE.md

Guidance for working in the **wordlive** repo. wordlive drives a *running*
Microsoft Word instance over COM (`pywin32`) — "xlwings, but for Word" — and is
built to be first-class for LLM tool use, not retrofitted. Windows-only.

## Core invariants — do not break these

These four principles drive every API decision. A change that violates one is
almost certainly wrong:

1. **Politeness first.** Operations preserve the user's `Selection`, view, and
   scroll. `doc.edit()` snapshots and restores selection + scroll on exit. Only
   `go_to()` and `scope.allow_cursor_move()` are allowed to move the user, and
   they say so explicitly.
2. **Semantic anchors over `Selection`.** Operations target named handles
   (bookmarks, content controls, headings, paragraphs, cells, ranges,
   headers/footers, `start`/`end`) — never the live cursor unless explicitly
   asked. Anchor ids are the stable, LLM-visible addressing scheme:
   `heading:N`, `para:N`, `bookmark:NAME`, `cc:NAME`, `table:N:R:C`,
   `range:START-END`, `header:S:WHICH` / `footer:S:WHICH`, `start`, `end`.
3. **Atomic undo.** Every `doc.edit("label")` opens a Word `UndoRecord`, so one
   Ctrl-Z reverts the whole intent. A multi-op `exec` batch is one undo step.
   (Pre-2010 Word silently falls back to N undo entries.)
4. **Structured I/O.** Reads return dataclasses/dicts; the CLI emits exactly one
   JSON object per invocation; exit codes are deterministic. No string scraping.

Underlying all four: the **`.com` escape hatch** — every wrapper exposes the raw
pywin32 object via `.com`. Prefer extending the high-level API, but never block
the user from dropping to COM.

## Layout

Flat by design — Word → Document → Anchor, ~no hierarchy beyond that. The two
central types are big enough to live in packages; everything else is one module.

- `src/wordlive/__init__.py` — the public surface (`__all__`). Keep it and
  `docs/python-api.md` in sync.
- `_app.py` `attach`/`connect`/`Word`; `_edit.py` `EditScope`; `_selection.py`
  snapshot/restore.
- `_document/` — `Document`, assembled from `DocumentCore` (`_core.py`: the COM
  handle, `edit`, `anchor_by_id` (anchor-id resolution lives here), the
  collection accessors) plus the `_editing` / `_reading` / `_structure` /
  `_persistence` mixins. `__init__.py` composes them and keeps
  `DocumentCollection`.
- `_anchors/` — `_base.py` assembles the `Anchor` ABC from `AnchorCore`
  (`_anchor_core.py`) + the `_anchor_{insert,media,references,format,lists,read}`
  mixins. Concrete anchors sit in siblings (`_range`, `_headings`, `_bookmarks`,
  `_content_controls`, `_paragraphs`, and `_{image,equation,chart,shape}_anchors`
  — suffixed to stay distinct from the same-named feature modules); free helpers
  in `_helpers.py` / `_refs.py`.
- Feature modules: `_tables.py`, `_lists.py`, `_sections.py`, `_comments.py`,
  `_styles.py`, `_images.py`, `_snapshot.py`, `_findreplace.py`.
- `_ops.py` — the `exec` batch op vocabulary (shared by CLI and MCP).
- `_com.py` COM plumbing; `exceptions.py` typed errors + HRESULT mapping;
  `constants.py` typed `IntEnum` mirrors of Word `Wd*` magic numbers.
- `cli/` — Click CLI (`commands/` = one module per verb cluster, aggregated by
  `register()`; `main.py` = exit-code boundary). `mcp/` — the `wordlive-mcp`
  dispatch-tool server (`server.py` builds it; `_read`/`_write`/`_exec`/
  `_snapshot` hold the impls).
- `_skill/{wordlive-cli,wordlive-python,wordlive-mcp}/SKILL.md` — the three
  bundled agent guides. The cli/python ones are surfaced by `llm-help` /
  `install-skill`; the **mcp** one is what `word_read(command="guide")` and the
  `wordlive://guide` resource serve (so an MCP agent is taught the `word_*` tools,
  not CLI verbs). `_guide.py` reads them (`skill_body(kind)`).

**Mixins and typing.** `DocumentCore`/`AnchorCore` expose a TYPE_CHECKING-only
`_as_document` / `_as_anchor` property (`cast("Document", self)`) — use it when a
mixin hands `self` to a collaborator annotated on the concrete class, or calls a
sibling mixin's method. Do *not* try `def m(self: Document)`: mypy requires a
`self` annotation to be a **supertype** of its class, and `Document` is a subtype
of its own mixins.

## Design surfaces stay parallel

The same capability is exposed three ways and they must agree:

- **CLI** verb (`cli/commands.py`), **`exec` op** (`_ops.py`), and **MCP**
  dispatch command (`mcp/server.py`, `word_read`/`word_write`/`word_exec`/
  `word_snapshot`). Adding a feature usually means touching all three plus the
  Python API.
- Prefer **few dispatch tools with a `command` arg** over many fine-grained
  tools (this is why MCP has just four `word_*` tools).
- Docs that must track code: `docs/cli.md`, `docs/mcp.md`, `docs/python-api.md`,
  `README.md`, `src/wordlive/_skill/SKILL.md`, and `CHANGELOG.md`. When you add
  or change a command/op/tool, update the relevant ones in the same change.

## Exit codes (CLI) — defined in `cli/main.py`

`0` ok · `1` other/bad-input (incl. `ImageSourceError`, `SnapshotError`,
`DocumentNotFoundError`) · `2` anchor/style not found or zero `find` matches ·
`3` Word busy (retryable) · `4` Word not running · `5` ambiguous `find` match.

## Build / test / lint

`uv`-managed; the `.venv` has no `pip`. Use `uv`, not `pip`/`python -m build`.

```bash
uv run pytest            # unit tests (smoke/e2e excluded by default)
uv run pytest -m smoke   # requires a real Word install on Windows
uv run pytest -m e2e     # CLI-subprocess lifecycle vs live Word (tests/test_e2e_cli.py)
uv run ruff format .     # formatter
uv run ruff check .      # lint (E/F/W/I/UP/B; E501 left to the formatter)
uv run mypy              # type-check src/wordlive (py310 target)
uv build                 # build sdist/wheel
```

**Once per clone**, install the git hooks — they don't self-install:

```bash
uvx pre-commit install   # then every commit re-locks uv.lock if it drifted
```

The `uv-lock` hook keeps `uv.lock` in sync with `pyproject.toml`; skipping it is
how a release once shipped a stale lockfile that failed `uv lock --locked` in CI.

Tests use a `fake_word` COM fixture (`tests/conftest.py`), so most run
off-Windows; anything marked `smoke` needs Word (`e2e` implies `smoke` — it
shells out to `python -m wordlive` against a live instance). CI
(`.github/workflows/ci.yml`) runs lint + tests across Python 3.10–3.15. Optional extras: `snapshot` (PyMuPDF,
for rendering), `mcp` (the server), `docs` (mkdocs).

## Release / workflow

- Branch + PR via `gh` — no longer direct-to-main since the PyPI release.
- `bump-my-version` bumps `pyproject.toml`, commits `chore: release vX.Y.Z`, and
  tags `vX.Y.Z`; the tag push triggers PyPI trusted-publishing
  (`.github/workflows/release.yml`). Update `CHANGELOG.md` (Keep a Changelog /
  SemVer) as part of the release.
