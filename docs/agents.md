# Agents & LLM tools

wordlive is built to be driven by an LLM, not just a human. There are two ways
to wire it into a coding agent or assistant, and they compose:

- **Skills** — drop a `SKILL.md` file where a coding agent (Claude Code, Cursor,
  …) discovers it, and the agent learns the whole CLI / Python surface. The
  agent then shells out to `wordlive …` itself.
- **MCP** — run wordlive's MCP server so an MCP client (Claude Desktop, Claude
  Code, …) calls four `word_*` tools directly, no shell.

Both talk to the **already-open** Word document on the same Windows machine.
Pick whichever your tool supports best (Claude Code supports both; Claude
Desktop is MCP-only; a plain coding agent in your editor wants the skill).

## The zero-install path: `llm-help`

Before configuring anything, any agent that can run a shell command can read the
entire guide in one shot — no install, no Word needed:

```bash
wordlive llm-help            # the CLI workflow guide
wordlive llm-help --python   # the `import wordlive as wl` guide
```

Point your agent at `wordlive --help` and it's told to run exactly this. This is
the lowest-friction way to give a one-off agent everything it needs.

## Install the skill

A skill teaches a coding agent the whole wordlive surface so it shells out to
`wordlive …` (or `import wordlive`) itself — no MCP server needed. It lands as a
`SKILL.md` that any skill-aware agent (Claude Code, Cursor, …) auto-discovers:

```bash
wordlive install-skill            # CLI skill → ./.agents/skills/wordlive-cli/SKILL.md
wordlive install-skill --python   # the Python-API skill instead
wordlive install-skill --both     # both
```

Project-local by default (commit it to share); `--system` installs to
`~/.agents/skills/` for every project.

## Connect the MCP server

The MCP server is one stdio process exposing the four `word_*` tools. Every
client launches it the same way — this is the entry to register:

```json
{
  "mcpServers": {
    "wordlive": {
      "command": "uvx",
      "args": ["--from", "wordlive[mcp,snapshot]", "wordlive-mcp"]
    }
  }
}
```

`uvx` runs the published package straight from PyPI — no separate install — and
the `snapshot` extra enables the `word_snapshot` vision tool (renders a page to
an image the model can *see*). Two variants on the launch command:

- **Pinned install:** `pip install "wordlive[mcp,snapshot]"`, then use
  `"command": "wordlive-mcp"` with no `args`.
- **Local checkout (dev):** `"command": "uv"`, `"args": ["run", "--directory",
  "C:\\path\\to\\wordlive", "wordlive-mcp"]`.

`wordlive install-mcp` writes this entry for you (Claude Desktop / Claude Code);
`wordlive install-mcp --print` emits the snippet to paste anywhere else. **Restart
the client** after editing its config. Because wordlive drives Word over COM, the
client must run on the **same Windows machine** as Word — the paths below are the
Windows locations.

### Claude Desktop (MCP-only)

Easiest is the one-click bundle: download `wordlive.mcpb` (built from
[`mcpb/`](https://github.com/thomas-villani/wordlive/tree/main/mcpb)) and drop it
onto **Settings → Extensions**. Otherwise register the server entry:

```bash
wordlive install-mcp        # writes %APPDATA%\Claude\claude_desktop_config.json
```

or edit that file by hand (**Settings → Developer → Edit Config** opens it) with
the `mcpServers` entry above.

### Claude Code (skill **and** MCP)

```bash
# Built-in CLI — user scope, so it's available in every project:
claude mcp add --scope user --transport stdio wordlive \
    -- uvx --from "wordlive[mcp,snapshot]" wordlive-mcp

# …or a committable project-local .mcp.json:
wordlive install-mcp --client claude-code
```

`--scope user` stores it in `~/.claude.json` (this machine only); `--scope
project` (and `install-mcp --client claude-code`) write `.mcp.json` at the repo
root, which you commit to share. Run `/mcp` inside Claude Code to confirm it
connected.

### Cursor

Paste the `mcpServers` entry above into `.cursor/mcp.json` (this project) or
`%USERPROFILE%\.cursor\mcp.json` (every project) — same shape, so
`wordlive install-mcp --print` gives you exactly what to drop in. Then enable it
under **Settings → Tools & MCP**.

### VS Code (Copilot agent mode)

VS Code uses a **different shape**: a top-level `servers` key plus an explicit
`"type": "stdio"`. Put this in `.vscode/mcp.json` (workspace) or your user config
(command palette → **MCP: Open User Configuration**):

```json
{
  "servers": {
    "wordlive": {
      "type": "stdio",
      "command": "uvx",
      "args": ["--from", "wordlive[mcp,snapshot]", "wordlive-mcp"]
    }
  }
}
```

Start it from the gutter **Start** action (or **MCP: List Servers**), then select
it in the Copilot Chat **Agent** tool picker.

### Windsurf

Add the `mcpServers` entry above to
`%USERPROFILE%\.codeium\windsurf\mcp_config.json` (**Cascade → MCP servers →
Manage → View raw config**), then hit **Refresh**.

### Other clients (Cline, …)

Anything that speaks MCP takes the same launch command, and most reuse the
`mcpServers` shape — `wordlive install-mcp --print` (or `--config PATH` to write a
specific file) gives you the snippet. Drop it wherever that client keeps its
server list (e.g. Cline: **MCP Servers → Configure → Installed**).

## What the agent works with

Whichever path you choose, the agent drives the same model:

- **Anchor IDs** — every range is addressed by a stable id (`heading:3`,
  `para:12`, `bookmark:Address`, `table:1:2:2`, …), so the agent targets named
  handles, never the live cursor. See [Concepts](concepts.md#anchor-ids).
- **Structured I/O** — the CLI emits one JSON object per call; the MCP tools
  return structured results. No string scraping.
- **Deterministic failures** — typed errors map to fixed exit codes / MCP
  `code`s so the agent can branch on the failure mode. See
  [Errors & exit codes](errors.md).
- **The one-page guide** — over MCP, fetch the whole anchor-model + op
  vocabulary as a tool call with `word_read(command="guide")` (also the
  `wordlive://guide` resource).

For the full tool list and op vocabulary see [MCP server](mcp.md); for the
command surface see the [CLI reference](cli.md).
