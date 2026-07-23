# wordlive

[![PyPI](https://img.shields.io/pypi/v/wordlive.svg)](https://pypi.org/project/wordlive/)
[![Python versions](https://img.shields.io/pypi/pyversions/wordlive.svg)](https://pypi.org/project/wordlive/)
[![License: MIT](https://img.shields.io/pypi/l/wordlive.svg)](https://github.com/thomas-villani/wordlive/blob/main/LICENSE)
[![CI](https://github.com/thomas-villani/wordlive/actions/workflows/ci.yml/badge.svg)](https://github.com/thomas-villani/wordlive/actions/workflows/ci.yml)
[![Docs](https://github.com/thomas-villani/wordlive/actions/workflows/docs.yml/badge.svg)](https://thomas-villani.github.io/wordlive/)

**Let an AI assistant — or your own scripts — read and edit the Word document
you already have open.** Politely: your cursor and scroll position never move,
edits target the document's *structure* (headings, bookmarks, tables — never
the live cursor), and every change is a single Ctrl-Z. Windows-only.

For developers: wordlive drives a *running* Word instance over COM —
"`xlwings`, but for Word" — with a Python API, a JSON-in/JSON-out CLI, and an
MCP server that all share one anchor-based addressing scheme.

<!-- TODO(demo): docs/assets/regularize.gif — messy doc → `wordlive regularize`
     → clean → Ctrl-Z restores. Choreography: examples/demos/demo_regularize.py -->
<!-- TODO(demo): docs/assets/agent-drive.gif — an LLM building a styled document
     and leaving comments, live. Choreography: examples/demos/demo_agent_showcase.py -->

## Pick your path

**🖱️ "I use Claude Desktop — no code, please."**
Install [`uv`](https://docs.astral.sh/uv/getting-started/installation/) (one
small installer), download **`wordlive.mcpb`** from the
[latest release](https://github.com/thomas-villani/wordlive/releases/latest),
and drag it onto Claude Desktop → **Settings → Extensions**. Open a document in
Word and try: *"Lint my document and fix what's safe."* Details under
[MCP server](#mcp-server-claude-desktop--other-agents) below.

**🤖 "I'm wiring up an agent (CLI · MCP · skills)."**
`uv tool install wordlive`, then `wordlive llm-help` prints the whole LLM-ready
guide to stdout — no install step in the loop, no Word needed to read it.
Copy-paste setup for Claude Code, Cursor, and friends lives in
[Agents & LLM tools](https://thomas-villani.github.io/wordlive/agents/).

**🐍 "I write Python."**
`pip install wordlive` (or `uv add wordlive`) → the
[quickstart below](#python) and the
[Python API reference](https://thomas-villani.github.io/wordlive/python-api/).

## Install

```
pip install wordlive

# Add to a python project
uv add wordlive

# Or as a `uv` tool
uv tool install wordlive
```

(Requires Python 3.10+ and `pywin32` on Windows.)

Rendering pages to PNG (`snapshot`) needs the optional `snapshot` extra, which
pulls in PyMuPDF; the MCP server is the `mcp` extra:

```
pip install "wordlive[snapshot]"
uv add "wordlive[mcp,snapshot]"
```

## Tidy an entire document in one command

Normalizing a document's formatting is one of the most tedious jobs in office
work: the paragraph someone left double-spaced, the heading hand-set in Arial
16 instead of the style, the stray space before a period. Done by hand it's an
afternoon of scrolling and clicking. wordlive turns it into two calls:

```
wordlive lint          # pure read — severity-ranked findings, each with an anchor id
wordlive regularize    # apply the *fixable* ones — one atomic undo (--dry-run to preview)
```

Every finding says what's off, where, and how it would fix it:

```json
{
  "rule": "space-before-punctuation",
  "severity": "info",
  "anchor_id": "para:4",
  "message": "Whitespace before punctuation.",
  "fixable": true,
  "observed": "space before , . ; : )",
  "expected": "no space before punctuation"
}
```

It's a power tool for both audiences: an LLM agent can `lint`, reason over the
structured findings, and apply exactly the fixes it wants; a human can run
`wordlive regularize` and reclaim the afternoon. `lint` never touches the
document, `regularize` never moves your cursor, and **one Ctrl-Z reverts the
whole pass**. House styles are configurable via `--profile`. The guided tour is
in [Linting & regularizing](https://thomas-villani.github.io/wordlive/linting/).

## Python

```python
import wordlive as wl

with wl.attach() as word:
    doc = word.documents.active

    # Reads
    outline = doc.outline()
    bookmarks = doc.bookmarks.list()

    # See it the way a vision model would — render a section to PNG
    # (needs `wordlive[snapshot]`):
    png = doc.heading("Introduction").snapshot()[0].png

    # Polite writes — preserves the user's cursor and view, atomic Ctrl-Z.
    with doc.edit("Update address block"):
        doc.bookmarks["Address"].set_text("123 Main St")
        doc.content_controls["Signatory"].set_text("Jane Doe")
        doc.heading("Introduction").insert_paragraph_after("New context paragraph.")

    # Hand back a deliverable (the Python API is ungated):
    doc.export_pdf("report.pdf")
```

## CLI

JSON in, JSON out — designed to drop straight into an LLM tool-use loop. A
taste of the surface:

```
wordlive status                                   # is Word up? which document?
wordlive outline                                  # heading structure → heading:N ids
wordlive read section "Introduction"              # the body under a heading
wordlive find-paragraph --text "roughly remembered text"    # fuzzy locate → ranked para:N
wordlive replace --anchor-id heading:3 --text "Updated section text"
printf '# Plan\n\n- scope it\n- staff it\n' | wordlive insert-markdown --anchor-id end --markdown -
wordlive table records 1                          # body rows as {header: value} dicts
wordlive table update-row --table 1 --key Travel --values '{"Cost":"$450"}'
wordlive comment add --anchor-id heading:3 --text "Please expand this." --author Bot
wordlive track on                                 # record edits as revisions; `revisions` reads them back
wordlive lint                                     # audit formatting/structure
wordlive regularize                               # …fix the safe findings, one Ctrl-Z
wordlive snapshot --page 2 --out p2.png           # render a page for a vision model
wordlive exec --script ops.json                   # batch N ops in a single undo
wordlive --save-dir C:\out export-pdf C:\out\report.pdf   # saving is gated by a whitelist
```

That's a sample, not the surface: tables, images, charts, equations, citations
& bibliographies, themes, floating shapes, headers/footers, lists, sections,
breaks, document properties & variables, checkpoint/diff, proofing stats, and
more are covered verb-by-verb in the
[CLI reference](https://thomas-villani.github.io/wordlive/cli/) — or run
`wordlive llm-help` for the one-page agent guide.

Anchor ids are the addressing scheme everywhere: `heading:N`, `para:N`,
`bookmark:NAME`, `cc:NAME`, `table:N:R:C`, `range:START-END`, `start`, `end`, …
Positional ids renumber under edits; `wordlive pin heading:3 --name methods`
mints a durable `pin:methods`.

An `exec` script batches many ops into one Ctrl-Z:

```json
{
  "label": "Update report",
  "ops": [
    {"op": "write_bookmark", "name": "Address", "text": "123 Main St"},
    {"op": "insert_paragraph", "anchor_id": "heading:3", "text": "New risk paragraph."},
    {"op": "apply_style", "anchor_id": "heading:3", "name": "Heading 2"}
  ]
}
```

Exit codes: `0` ok, `1` other, `2` anchor-not-found, `3` Word-busy, `4`
Word-not-running, `5` ambiguous-match (`replace --find` hit several), `6`
Excel-not-available (`insert-chart`).

## Agent skills

> Setting up a specific tool (Claude Code, Claude Desktop, Cursor, …)? The
> [Agents & LLM tools](https://thomas-villani.github.io/wordlive/agents/) guide
> has copy-paste setup per client.

wordlive ships **two** LLM-facing skills (`SKILL.md`): `wordlive-cli` (the
command-line workflow) and `wordlive-python` (the `import wordlive as wl` API).
Each covers the anchor model, every verb, and the exit-code / exception contract.

An agent that hits `wordlive --help` is pointed straight at `wordlive llm-help`,
which prints the whole guide to stdout in one shot — no install step, no Word:

```
wordlive llm-help                 # the CLI guide
wordlive llm-help --python        # the Python-API guide
```

Or drop the skill files into a project or your home directory so coding tools
discover them on their own (CLI skill by default; `--python` for just Python,
`--both` for both):

```
wordlive install-skill            # ./.agents/skills/wordlive-cli/SKILL.md
wordlive install-skill --both     # also drops wordlive-python/SKILL.md
wordlive install-skill --system   # into ~/.agents/skills/ instead
```

## MCP server (Claude Desktop & other agents)

Prefer MCP? `wordlive` ships a server so Claude Desktop and other MCP clients
can drive your open document directly. Three ways to set it up, easiest first:

**1. One-click bundle.** Download `wordlive.mcpb` from the
[latest release](https://github.com/thomas-villani/wordlive/releases/latest)
and drop it onto Claude Desktop → **Settings → Extensions**. It needs
[`uv`](https://docs.astral.sh/uv/getting-started/installation/) on your PATH —
the bundle resolves the published package at load time, so there's nothing else
to install. (The bundle's source lives in
[`mcpb/`](https://github.com/thomas-villani/wordlive/tree/main/mcpb).)

**2. `install-mcp`.** Register the server in your client's config in one command
(it uses `uvx`, so there's no separate install step):

```
wordlive install-mcp                      # → Claude Desktop's config
wordlive install-mcp --client claude-code # → ./.mcp.json
wordlive install-mcp --print              # just print the JSON snippet
```

**3. By hand.** `pip install "wordlive[mcp,snapshot]"` (the `snapshot` extra adds
the vision tool), then add to `claude_desktop_config.json`:

```json
{ "mcpServers": { "wordlive": { "command": "wordlive-mcp" } } }
```

Or, if you prefer `uvx`:

```json
{ "mcpServers": { "wordlive": { "command": "uvx wordlive[mcp,snapshot]" } } }
```

It exposes four dispatch tools — `word_read`, `word_write`, `word_exec`, and
`word_snapshot` (which returns a rendered page as an image). The full op
vocabulary and anchor model are in the one-page guide, fetchable as a tool call
with `word_read(command="guide")` (also the `wordlive://guide` resource). Word
must be running on the same Windows machine. See
[docs/mcp.md](https://thomas-villani.github.io/wordlive/mcp/).

## Examples

Runnable, out-of-the-box scripts live in
[`examples/`](https://github.com/thomas-villani/wordlive/tree/main/examples) — Python
(using the library) and PowerShell (driving the CLI). Each attaches to the
document you already have open; the read-only and append-only ones are safe to
try on a real document. (`examples/demos/` holds the scripted walk-throughs
behind the demo GIFs — they run against a throwaway document.)

```bash
python examples/python/read_outline.py            # read-only: print the outline
python examples/python/append_note.py "Reviewed." # append one paragraph (atomic, polite)
```

```powershell
.\examples\powershell\Show-Outline.ps1
.\examples\powershell\Invoke-WordliveWithRetry.ps1 write bookmark Address --text "123 Main St"
```

## Design

- **Politeness first** — operations preserve the user's `Selection`, view, and
  scroll. The user keeps editing alongside you.
- **Semantic anchors over `Selection`** — operations target bookmarks, content
  controls, or headings — never the live cursor unless you ask.
- **Atomic undo** — every `doc.edit()` opens a Word `UndoRecord`, so a single
  Ctrl-Z reverts the whole block.
- **Escape hatch** — every wrapper exposes `.com` for the raw COM object;
  you're never blocked by missing coverage.

See [`spec.md`](https://github.com/thomas-villani/wordlive/blob/main/spec.md) for the full design.

## Contributing & security

Contributions are welcome — see
[`CONTRIBUTING.md`](https://github.com/thomas-villani/wordlive/blob/main/CONTRIBUTING.md)
for the dev setup (`uv`), the design invariants, and the test / lint / docs
gates. To report a security issue, please use private disclosure as described in
[`SECURITY.md`](https://github.com/thomas-villani/wordlive/blob/main/SECURITY.md)
rather than a public issue.
