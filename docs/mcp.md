# MCP server

`wordlive` ships an [MCP](https://modelcontextprotocol.io) server so MCP clients —
**Claude Desktop**, Cursor, and other agent hosts — can read and edit the Word
document you have open right now, including *seeing* rendered pages.

It talks to the same running Word instance the CLI and Python API do, over COM, on
Windows. Edits stay polite (your cursor, selection, and scroll are preserved; each
write is a single Ctrl-Z), and failures come back with a stable error `code` and a
`retryable` hint.

## Install

Three ways, easiest first. All three end with the same `word_*` tools in your
client — pick whichever fits.

### 1. One-click bundle (`.mcpb`)

The repo ships an **MCP bundle** in
[`mcpb/`](https://github.com/thomas-villani/wordlive/tree/main/mcpb) — a single
`wordlive.mcpb` file you drag onto Claude Desktop → **Settings → Extensions** to
install in one click, no JSON editing. The bundle carries no code: it declares a
dependency on the published `wordlive[mcp,snapshot]` package and Claude Desktop's
`uv` runtime resolves it on first run. Needs [`uv`](https://docs.astral.sh/uv/)
on `PATH`. See [`mcpb/README.md`](https://github.com/thomas-villani/wordlive/blob/main/mcpb/README.md)
to download or rebuild it.

### 2. `wordlive install-mcp`

Let the CLI write the config entry for you (it launches the server with `uvx`,
so there's no separate install step):

```
wordlive install-mcp                       # merge into Claude Desktop's config
wordlive install-mcp --client claude-code  # write a project-local ./.mcp.json
wordlive install-mcp --print               # just print the JSON snippet
wordlive install-mcp --directory .         # dev: run a local checkout via uv run
```

It merges an `mcpServers.wordlive` entry (using `uvx --from "wordlive[mcp,snapshot]"
wordlive-mcp`) into the target config, refusing to clobber an existing entry
without `--force`. Offline — it never touches Word. Restart the client afterward.

### 3. By hand

Install the optional extra, then add the entry yourself:

```
pip install "wordlive[mcp,snapshot]"   # snapshot extra adds the vision tool
uv tool install "wordlive[mcp,snapshot]"
```

```json
{
  "mcpServers": {
    "wordlive": {
      "command": "wordlive-mcp"
    }
  }
}
```

If `wordlive-mcp` isn't on Claude Desktop's `PATH`, point at the interpreter in
your environment explicitly (note the doubled backslashes in JSON):

```json
{
  "mcpServers": {
    "wordlive": {
      "command": "C:\\Users\\you\\project\\.venv\\Scripts\\python.exe",
      "args": ["-m", "wordlive.mcp"]
    }
  }
}
```

Claude Desktop's config lives at *Settings → Developer → Edit Config*. Restart
Claude Desktop, open a `.docx` in Word, and the `word_*` tools appear.

## Run

```
wordlive-mcp            # console script (stdio transport)
python -m wordlive.mcp  # equivalent
```

The server speaks MCP over **stdio** — the transport Claude Desktop spawns. Word
must already be running on the same machine (wordlive *attaches*; it never launches
or closes Word).

## Tools

A small set of **dispatch tools** keeps the client's tool list (and its context
cost) lean. Every tool takes an optional `doc` (target a document by name; default
is the active one).

| Tool | What it does |
| --- | --- |
| `word_read` | Every read, dispatched on `command`: `guide`, `status`, `outline`, `paragraphs`, `find`, `read_bookmark`, `read_cc`, `read_section`, `table_list`, `table_read`, `styles`, `comments`, `sections`. |
| `word_write` | Every single atomic-undo edit, dispatched on `command`: `insert`, `insert_break`, `append`, `prepend`, `replace`, `write_bookmark`, `write_cc`, `apply_style`, `format_paragraph`, `format_run`, `set_shading`, `set_borders`, `add_tab_stop`, `add_style`, `set_style`, `list`, `comment`, `table` (`action` = `set_cell`/`add_row`/`delete_row`/`create`/`delete`), `header`, `footer`, `track`, `insert_image`. `append`/`prepend` add a new paragraph (optional `style`); pass `paragraph: false` to continue the adjacent paragraph inline (no `style`). `format_run` styles characters (bold/italic/colour/highlight/…); `set_shading`/`set_borders`/`add_tab_stop` add cell/range shading, borders, and tab stops; `add_style`/`set_style` create and configure styles (the border line style is the `line_style` param, to avoid colliding with `style`). |
| `word_exec` | Apply a batch of `ops` as a **single** atomic undo — the power tool for multi-step intents. |
| `word_snapshot` | Render page(s) to PNG so the model can *see* the layout. Returns image content. Needs the `snapshot` extra. |

The anchor model (`heading:N`, `para:N`, `bookmark:NAME`, `cc:NAME`,
`table:N:R:C`, `range:START-END`, `header:S:WHICH` / `footer:S:WHICH`, `start` /
`end`) and the full `word_exec` op vocabulary are documented in the one-page
guide. Fetch it as a tool call with **`word_read(command="guide")`** (it needs
neither Word nor a document) — the most reliable path, since not every MCP
client surfaces resources. The same text is also the **`wordlive://guide`**
resource and `wordlive llm-help`. See also [CLI](cli.md) for each op's fields.

`heading:N` / `para:N` are **positional** and renumber when a structural edit
shifts the document, so re-read `outline` / `paragraphs` after an insert before
reusing ids; `bookmark:NAME` / `cc:NAME` are name-based and survive edits.

### Batches

`word_exec` mirrors the CLI's [`exec`](cli.md):

```json
{
  "ops": [
    {"op": "write_bookmark",   "name": "Address",      "text": "123 Main St"},
    {"op": "insert_paragraph", "anchor_id": "heading:2", "text": "New section.", "style": "Body Text"},
    {"op": "find_replace",     "find": "Q3", "text": "Q4", "all": true}
  ],
  "tracked": false
}
```

It stops at the first failing op and reports its `index`, `op`, and error `type`;
the successful prefix still rolls back as one undo step. A field an op doesn't
use (a typo, or `style` on an inline append) is reported in a `warnings` array on
the result rather than silently dropped.

## Errors

A failed tool call comes back flagged as an error whose message is a JSON object:

```json
{"error": "bookmark not found: 'Addr'", "code": "anchor_not_found", "retryable": false, "type": "AnchorNotFoundError"}
```

`code` / `retryable` mirror the CLI's [exit-code taxonomy](errors.md):

| `code` | Meaning | `retryable` |
| --- | --- | --- |
| `anchor_not_found` | anchor missing, or `find` matched zero | no — re-read first |
| `style_not_found` | a named style isn't defined in the document | no — read `styles` first |
| `ambiguous_match` | `find` matched several | yes — pass `occurrence`/`all` |
| `replace_verification` | a `find_replace` target couldn't be verified (e.g. a whole-doc match inside a table) | no — scope to the cell anchor (`table:N:R:C`) |
| `word_busy` | a modal dialog is open | **yes** — back off and retry |
| `word_not_running` | Word isn't running | no — until Word is opened |
| `document_not_found` | no document by that name | no |
| `error` | bad input / other | no — fix the request |

## How it works

All Word access funnels through a single dedicated, COM-initialised worker thread,
so COM stays on one apartment and concurrent tool calls serialise — matching
Word's single-threaded reality. Each call attaches to the running instance fresh,
exactly like the CLI. The server is in-process: it calls the wordlive Python API
directly rather than shelling out, which is also how `word_snapshot` returns native
image content.
