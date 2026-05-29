# wordlive MCP bundle (`.mcpb`)

This directory packages wordlive's MCP server as an **MCP bundle** — a single
`.mcpb` file you drag onto Claude Desktop's *Extensions* pane to install the
server in one click, no JSON editing.

The bundle carries no wordlive code: `manifest.json` + `pyproject.toml` declare a
dependency on the published `wordlive[mcp,snapshot]` package, and `src/server.py`
is a thin launcher. At load time Claude Desktop's `uv` runtime resolves and runs
`wordlive-mcp` from PyPI.

## Requirements

- Windows with Microsoft Word installed (wordlive drives a live Word over COM).
- A recent Claude Desktop (the one with the *Extensions* / MCP-bundle installer).
- [`uv`](https://docs.astral.sh/uv/) on PATH — the bundle's declared runtime.

## Install (end users)

1. Download `wordlive.mcpb` (a release asset, or build it below).
2. Open Claude Desktop → **Settings → Extensions** and drop the file in (or
   double-click it).
3. Open a `.docx` in Word, and the `word_read` / `word_write` / `word_exec` /
   `word_snapshot` tools appear.

Prefer to wire it up by hand? `wordlive install-mcp` writes the same
`mcpServers` entry into your client config — see the
[MCP docs](https://thomas-villani.github.io/wordlive/mcp/).

## Build / repack

```bash
npm install -g @anthropic-ai/mcpb   # one-time
mcpb validate manifest.json
mcpb pack . wordlive.mcpb            # from this dir (bare `mcpb pack` names it mcpb.mcpb)
```

The bundle version is kept in lock-step with the package by `bump-my-version`
(see `[tool.bumpversion]` in the root `pyproject.toml`): a release bumps
`manifest.json`, this `pyproject.toml`, and its `wordlive[mcp,snapshot]>=` pin
together.
