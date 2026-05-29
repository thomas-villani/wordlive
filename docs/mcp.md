# MCP server

`wordlive` ships an [MCP](https://modelcontextprotocol.io) server so MCP clients —
**Claude Desktop**, Cursor, and other agent hosts — can read and edit the Word
document you have open right now, including *seeing* rendered pages.

It talks to the same running Word instance the CLI and Python API do, over COM, on
Windows. Edits stay polite (your cursor, selection, and scroll are preserved; each
write is a single Ctrl-Z), and failures come back with a stable error `code` and a
`retryable` hint.

## Install

The server is an optional extra:

```
pip install "wordlive[mcp]"
# or, with the snapshot (vision) tool too:
pip install "wordlive[mcp,snapshot]"

# uv
uv tool install "wordlive[mcp,snapshot]"
```

## Run

```
wordlive-mcp            # console script (stdio transport)
python -m wordlive.mcp  # equivalent
```

The server speaks MCP over **stdio** — the transport Claude Desktop spawns. Word
must already be running on the same machine (wordlive *attaches*; it never launches
or closes Word).

## Register with Claude Desktop

Add an entry to `claude_desktop_config.json` (Claude Desktop → Settings →
Developer → Edit Config):

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

Restart Claude Desktop, open a `.docx` in Word, and the `word_*` tools appear.

## Tools

A small set of **dispatch tools** keeps the client's tool list (and its context
cost) lean. Every tool takes an optional `doc` (target a document by name; default
is the active one).

| Tool | What it does |
| --- | --- |
| `word_read` | Every read, dispatched on `command`: `status`, `outline`, `paragraphs`, `find`, `read_bookmark`, `read_cc`, `read_section`, `table_list`, `table_read`, `styles`, `comments`, `sections`. |
| `word_write` | Every single atomic-undo edit, dispatched on `command`: `insert`, `insert_break`, `append`, `prepend`, `replace`, `write_bookmark`, `write_cc`, `apply_style`, `format_paragraph`, `list`, `comment`, `table` (`action` = `set_cell`/`add_row`/`delete_row`/`create`/`delete`), `header`, `footer`, `track`, `insert_image`. |
| `word_exec` | Apply a batch of `ops` as a **single** atomic undo — the power tool for multi-step intents. |
| `word_snapshot` | Render page(s) to PNG so the model can *see* the layout. Returns image content. Needs the `snapshot` extra. |

The anchor model (`heading:N`, `para:N`, `bookmark:NAME`, `cc:NAME`,
`table:N:R:C`, `range:START-END`, `header:S:WHICH` / `footer:S:WHICH`, `start` /
`end`) and the full `word_exec` op vocabulary are documented in the
**`wordlive://guide`** resource the server exposes — the same one-page guide as
`wordlive llm-help`. See also [CLI](cli.md) for each op's fields.

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
the successful prefix still rolls back as one undo step.

## Errors

A failed tool call comes back flagged as an error whose message is a JSON object:

```json
{"error": "bookmark not found: 'Addr'", "code": "anchor_not_found", "retryable": false, "type": "AnchorNotFoundError"}
```

`code` / `retryable` mirror the CLI's [exit-code taxonomy](errors.md):

| `code` | Meaning | `retryable` |
| --- | --- | --- |
| `anchor_not_found` | anchor or style missing, or `find` matched zero | no — re-read first |
| `ambiguous_match` | `find` matched several | yes — pass `occurrence`/`all` |
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
