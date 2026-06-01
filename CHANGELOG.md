# Changelog

All notable changes to **wordlive** are documented here.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

### Changed
- CI: the release workflow's `actions/setup-node` is bumped `v4` → `v5` (off the
  deprecated Node 20 action runtime; GitHub forces Node 24 after 2026-06-16), and
  the bundle build now uses Node 22 LTS instead of Node 20.

## [0.10.2] — 2026-05-31

### Fixed
- `insert_image` now resolves a relative path to an absolute one before handing
  it to `InlineShapes.AddPicture`. Word resolves a relative filename against
  *its own* working directory, not the caller's, so a relative `--path` (or
  `image=` argument) previously failed with COM `0x80020009` ("not a valid file
  name"). Relative paths from the CLI's working directory now embed correctly.

## [0.10.1] — 2026-05-29

### Fixed
- `word_snapshot` no longer double-encodes its rendered pages. The tool returns
  each page as an MCP image content block, but its `-> list[Any]` return made
  FastMCP infer a structured-output schema and re-serialise the base64 PNG bytes
  into `structuredContent` as well — sending every page twice (a large, silent
  token cost on hosts that forward `structuredContent`). Marked the tool
  `structured_output=False` so the image is sent exactly once.

### Changed
- CI: the release workflow now packs `mcpb/` into `wordlive.mcpb` and attaches it
  to the GitHub Release (built outside the PyPI upload, so it never reaches PyPI).

## [0.10.0] — 2026-05-29

### Added
- Runnable example scripts under `examples/` (Python + PowerShell) and an
  **Examples** docs page, linked from the README and getting-started.
- **Python-API agent skill** (`wordlive-python`) alongside the existing CLI
  skill (now `wordlive-cli`). `install-skill` installs the CLI skill by default;
  `--python` installs just the Python one, `--both` installs both. `llm-help
  --python` prints the Python guide.
- **MCP bundle** (`mcpb/`) — a one-click `.mcpb` for Claude Desktop, kept in
  version lock-step with the package via `bump-my-version`.
- **`wordlive install-mcp`** — register the MCP server in Claude Desktop or
  Claude Code (`--client`, `--directory`, `--config`, `--print`, `--force`).
- `wordlive-mcp` console script (`[project.scripts]`), which the MCP docs and
  bundle already reference.
- MIT `LICENSE`, with `license` / `license-files` declared in `pyproject.toml`
  and the bundle manifest.

### Changed
- CI: bumped GitHub Actions off the deprecated Node 20 runtime to current
  Node 24 majors (`checkout` v6, `setup-python` v6, `setup-uv` v8,
  `upload-artifact` v7, `download-artifact` v8, `upload-pages-artifact` v5,
  `deploy-pages` v5).
- Added this changelog.
- Docs audit for v0.9.0: corrected the documented Python floor (3.10+), the
  README exit-code list (added `5`), the MCP `word_write` command list (added
  `insert_break`), the `design.md` roadmap (snapshot/MCP/tables/breaks shipped),
  and populated the previously-empty `CLAUDE.md`.

## [0.9.0] — 2026-05-29

First release since 0.8.3. Bundles four features that were developed earlier but
had not yet been published.

### Added
- **Snapshots** — `Document.snapshot(...)` / `Anchor.snapshot(...)` and the
  `wordlive snapshot` command render page(s) or a section to PNG (Word exports a
  pixel-faithful PDF, PyMuPDF rasterises it) so a vision model can *see* the
  layout. Requires the optional `snapshot` extra (PyMuPDF).
- **MCP server** (`wordlive-mcp`) — four dispatch tools (`word_read`,
  `word_write`, `word_exec`, `word_snapshot`) plus a `wordlive://guide` resource,
  for Claude Desktop and other agents. Requires the optional `mcp` extra.
- **Table creation / deletion** — `Document.add_table(...)`,
  `Anchor.insert_table(...)`, and `Table.delete()`; the `wordlive table create`
  / `table delete` commands; and the `create_table` / `delete_table` exec ops.
  Populates cells from a row-major `data` grid, defaults to the `Table Grid`
  style, and separates appended tables so Word doesn't silently merge adjacent
  ones.
- **Page / column / section breaks** — `Anchor.insert_break(kind=...)` and
  `format_paragraph(page_break_before=...)`; the `wordlive insert-break` command
  and a `--page-break-before` flag on `format-paragraph`; the `insert_break` exec
  op and a `page_break_before` field on `format_paragraph`.

## [0.8.3] — 2026-05-26

### Added
- `llm-help` command that dumps the full agent guide to stdout.

## [0.8.2] — 2026-05-21

### Added
- `append` / `prepend` helpers and `start` / `end` anchors for the document
  edges, so a document can be built top-down from a blank page.

### Changed
- CI: lint + test workflow across Python 3.10–3.15.

## [0.8.1] — 2026-05-21

### Added
- Inline `exec` JSON via `--ops` (and `--ops -` for stdin), terminal-paragraph
  append, and `before`/`after` placement on exec insert ops.

### Changed
- Tooling: `bump-my-version` configuration.

## [0.8.0] — 2026-05-21

Initial PyPI release. Drives a running Microsoft Word instance over COM
(Windows), with a JSON-in / JSON-out CLI built for LLM agents. Highlights of the
v0–v0.8 development line bundled here:

### Added
- Live automation core: `attach()` / `connect()`, anchors (bookmark, content
  control, heading, paragraph, range, cell, header/footer, start/end),
  `Document.edit()` atomic-undo with cursor/selection preservation, typed
  exceptions, and deterministic CLI exit codes.
- Reading: `status`, `outline`, `paragraphs`, `find` (fuzzy), `read section`.
- Writing: `replace`, `insert`, fuzzy find/replace, styles + `format-paragraph`,
  tables (read/edit), comments, track changes, lists & numbering, sections /
  headers / footers, and image insertion (path / bytes / base64).
- Batch edits via `exec` (single atomic undo), the bundled agent skill +
  `wordlive install-skill`, an mkdocs Material docs site, and PyPI
  trusted-publishing on tag push.

[Unreleased]: https://github.com/thomas-villani/wordlive/compare/v0.10.0...HEAD
[0.10.0]: https://github.com/thomas-villani/wordlive/compare/v0.9.0...v0.10.0
[0.9.0]: https://github.com/thomas-villani/wordlive/compare/v0.8.3...v0.9.0
[0.8.3]: https://github.com/thomas-villani/wordlive/compare/v0.8.2...v0.8.3
[0.8.2]: https://github.com/thomas-villani/wordlive/compare/v0.8.1...v0.8.2
[0.8.1]: https://github.com/thomas-villani/wordlive/compare/v0.8.0...v0.8.1
[0.8.0]: https://github.com/thomas-villani/wordlive/releases/tag/v0.8.0
