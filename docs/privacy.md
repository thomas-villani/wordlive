# Privacy policy

_Last updated: 2026-08-04. Applies to the `wordlive` Python package, the
`wordlive` CLI, the `wordlive-mcp` MCP server, and the `wordlive.mcpb` bundle._

## The short version

**wordlive collects nothing, sends nothing, and stores nothing about you.**

It is a local automation library. It talks to a copy of Microsoft Word already
running on your own machine, over Windows COM — an operating-system mechanism
that never leaves the machine. There is no wordlive account, no wordlive server,
no telemetry, no analytics, no crash reporting, and no update check. The package
declares no HTTP client as a dependency and makes no network calls.

## What data wordlive touches

Only the documents you point it at, and only while a command is running:

- **Document content and structure** — text, styles, tables, comments, headers
  and footers, revisions, and document properties of the open document, read or
  modified as your command asks.
- **Rendered pages** — the `snapshot` feature asks Word to export the document
  to a temporary PDF, rasterises the requested page(s) to PNG locally, and
  deletes the temporary file. Nothing is uploaded.
- **File paths you supply** — e.g. an image to insert, or a save-as target.

All of it stays in memory or in files on your machine, under paths you choose.
wordlive writes no logs, caches, or usage records of its own.

## Third parties

wordlive has none. It integrates with software you already run:

- **Microsoft Word** handles your document, and its own behaviour — cloud
  autosave to OneDrive/SharePoint, connected experiences, telemetry — is
  governed by [Microsoft's privacy statement](https://privacy.microsoft.com/privacystatement)
  and your organisation's Office configuration, not by wordlive. If your
  document lives in OneDrive, Word syncs it whether or not wordlive is involved.
- **Your MCP client or AI assistant** (Claude Desktop, Cursor, an agent you
  wrote) is what decides which wordlive commands to run, and it is the thing
  that may transmit document content to a model provider. That transmission is
  the client's, under the client's privacy policy. wordlive returns results to
  whatever process invoked it and does nothing further with them.

Put plainly: if you ask an AI assistant to edit your document, the assistant
sees the parts of the document it reads. Choose your assistant accordingly.

## Retention

None. wordlive keeps no data after a command exits. Temporary files it creates
(the snapshot PDF, for example) are deleted when it is done with them.

## Children

wordlive is a developer tool, not directed at children, and collects no
personal information from anyone.

## Changes

Material changes to this policy will be noted in
[CHANGELOG.md](https://github.com/thomas-villani/wordlive/blob/main/CHANGELOG.md)
and reflected in the "last updated" date above.

## Contact

Questions, or a security concern? Open an issue at
[github.com/thomas-villani/wordlive/issues](https://github.com/thomas-villani/wordlive/issues),
or see [SECURITY.md](https://github.com/thomas-villani/wordlive/blob/main/SECURITY.md)
for private disclosure.
