# Security Policy

## Reporting a vulnerability

**Please do not report security vulnerabilities through public GitHub issues.**

Use GitHub's private vulnerability reporting instead: go to the repository's
**Security** tab → **Report a vulnerability** (this opens a private advisory only
the maintainers can see). If you can't use that, email the maintainer at
**thomas.villani@gmail.com** with `wordlive security` in the subject.

Please include:

- a description of the issue and the impact you think it has,
- the version of wordlive (and Word, if relevant),
- step-by-step reproduction — ideally a minimal CLI / MCP / script example,
- any proof-of-concept, logs, or crash output.

We aim to acknowledge a report within a few days and to keep you updated as we
work on a fix. We practice coordinated disclosure: please give us a reasonable
window to release a fix before any public disclosure, and we'll credit you
(unless you'd rather stay anonymous).

## Supported versions

wordlive is pre-1.0 and ships from a single line of development. Security fixes
land on the **latest released version**; please upgrade to the most recent
release before reporting, in case the issue is already fixed.

| Version          | Supported          |
| ---------------- | ------------------ |
| Latest release   | :white_check_mark: |
| Anything older   | :x:                |

## Threat model — what wordlive does and doesn't defend

wordlive automates a running Word instance with **the privileges of the user who
launched it**. Understanding which surface you're using matters:

- **The Python API is trusted and intentionally ungated.** It does exactly what
  your code tells it — including overwriting files via `save_as` / `export_pdf`
  or driving Word destructively. Treat it like any library you call directly:
  the trust boundary is your own code. This is by design, not a vulnerability.

- **The CLI and MCP surfaces are the gated, prompt-injection-aware boundary.**
  These are the surfaces an LLM can drive from untrusted document content, so
  their filesystem-touching inputs run through a **default-deny policy**:

  - **Saving** (`save` / `save-as` / `export-pdf`) writes **only** inside
    directories whitelisted with `--save-dir` (repeatable) / `WORDLIVE_SAVE_DIRS`.
    With none configured, saving is **off**. The target is resolved first (so
    `..` / symlinks can't escape) and then required to be inside the whitelist.
  - **Image sources** (`insert-image --path`) **reject non-local forms** — UNC
    paths (`\\host\share\…`), `file://`, and URLs — *before* any filesystem
    probe, because a UNC `is_file()` check alone would authenticate to a remote
    SMB server and can leak NTLM credentials (and URLs were an SSRF /
    local-file-disclosure vector). An optional `--image-dir` /
    `WORDLIVE_IMAGE_DIRS` allowlist further restricts which local directories a
    path may come from.

  Refusals raise `PathNotAllowedError` (CLI exit code `1`, MCP
  `code: "path_not_allowed"`). See
  [Errors & exit codes](https://thomas-villani.github.io/wordlive/errors/).

### In scope

- A way for **prompt-injected CLI / MCP input** to escape the path whitelist,
  read or write files outside the configured directories, exfiltrate data via
  the image-source path, or otherwise act beyond the documented gated contract.
- COM-error handling that could be turned into a more serious failure.
- Anything that lets untrusted document content trigger an action the gated
  surface is supposed to forbid.

### Out of scope / by design

- The **Python API being ungated** — it's a trusted library surface (see above).
- Destructive Word automation you explicitly asked for (the CLI/MCP `save`
  surface is off unless you opt in with a whitelist).
- Issues that require the attacker to already control the machine or the Word
  session.
- Vulnerabilities in Word, `pywin32`, or other dependencies — report those
  upstream (we'll happily bump a pin once a fix is available).

When in doubt, report it privately and let us make the call.
