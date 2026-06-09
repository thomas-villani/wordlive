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

## Claude Code

Two complementary moves — install the skill so Claude Code knows the surface,
and (optionally) register the MCP server so it can also call the tools directly:

```bash
# Skill — lands in ./.agents/skills/wordlive-cli/SKILL.md (commit it to share).
wordlive install-skill            # CLI skill (default)
wordlive install-skill --both     # also the wordlive-python skill

# MCP — merges a server entry into the project's ./.mcp.json (uses uvx, no install).
wordlive install-mcp --client claude-code
```

Restart Claude Code to pick up the `.mcp.json` change. The skill is committed
per-project; `--system` installs it to `~/.agents/skills/` for every project.

## Claude Desktop

Claude Desktop is MCP-only. Easiest first:

1. **One-click bundle.** Download `wordlive.mcpb` (built from
   [`mcpb/`](https://github.com/thomas-villani/wordlive/tree/main/mcpb)) and drop
   it onto **Settings → Extensions**.
2. **`install-mcp`.** Register the server in Claude Desktop's
   `claude_desktop_config.json` in one command (it launches the server with
   `uvx`, so there's no separate install):

   ```bash
   wordlive install-mcp                  # → Claude Desktop's config (default)
   ```
3. **By hand.** `pip install "wordlive[mcp,snapshot]"`, then add:

   ```json
   { "mcpServers": { "wordlive": { "command": "wordlive-mcp" } } }
   ```

Restart Claude Desktop afterwards. The `snapshot` extra adds the
`word_snapshot` vision tool (renders a page to an image the model can *see*).

## Cursor & other coding agents

Any agent that reads `.agents/skills/**/SKILL.md` (or that you can paste context
into) gets the skill the same way:

```bash
wordlive install-skill            # ./.agents/skills/wordlive-cli/SKILL.md
wordlive install-skill --python   # the Python-API skill instead
```

If your editor's agent speaks MCP, point its config at the server — print the
snippet for any client and paste it where that client expects `mcpServers`:

```bash
wordlive install-mcp --print
```

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
