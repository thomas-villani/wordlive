# Where to list & post wordlive

Roughly in order of value. The two demo GIFs should exist before anything in
"Launch posts" — those channels send traffic straight at the README.

## 1. Official MCP registry — **automated, nothing to do by hand**

`server.json` at the repo root registers wordlive as
**`io.github.thomas-villani/wordlive`**, and the `mcp-registry` job in
`.github/workflows/release.yml` republishes it on every `v*` tag: it waits for
PyPI, recomputes the `.mcpb` SHA-256, repoints the download URL at the new tag,
authenticates with GitHub OIDC (no secret), and publishes.

- [x] `server.json` written and schema-valid
- [x] `mcp-name: io.github.thomas-villani/wordlive` marker in `README.md`
      (this is how PyPI ownership is proven — it must be in the *published*
      description, so it only counts from the next release onward)
- [x] release workflow job
- [ ] **Cut a release** (the marker is not in 0.19.0's README, so the first
      publish has to be ≥ the next tag), then confirm:
      `curl "https://registry.modelcontextprotocol.io/v0.1/servers?search=wordlive"`

Aggregators (PulseMCP, Glama, mcp.so, Smithery, …) are *expected to scrape* the
official registry rather than take manual submissions, so this one listing is
what feeds most of the directory ecosystem. Check each after a release and only
submit by hand where the entry hasn't appeared.

> Note: `modelcontextprotocol/servers` **no longer accepts community-server
> PRs** — its README now points people at the registry instead. Skip it.

## 2. Anthropic's Claude connector / desktop-extension directory

The in-app directory is the jackpot for the no-code audience. Local MCPB
connectors go through the *desktop extension* submission form (not the remote
MCP one). Reviewers gate hard on privacy documentation.

- [x] `manifest_version` ≥ 0.2 (we're on 0.4)
- [x] `privacy_policies` array in `mcpb/manifest.json`
- [x] "Privacy Policy" section in `README.md` and `mcpb/README.md`
- [x] HTTPS privacy policy page — <https://thomas-villani.github.io/wordlive/privacy/>
- [ ] Submit via the desktop-extension form in the Claude developer docs
- [ ] Expect a slow review (community reports: two weeks to several months)

## 3. Other directories & lists (cheap, evergreen)

- [ ] punkpeye/awesome-mcp-servers — PR; largest MCP awesome list
- [ ] vinta/awesome-python — PR under office/docx processing (sits nicely next
      to python-docx as the "live Word" complement)
- [ ] PulseMCP / mcp.so / Glama / Smithery — only if the registry scrape hasn't
      picked wordlive up on its own

## 4. Launch posts (after the GIFs are embedded)

- [ ] **Show HN** — "Show HN: Wordlive – let AI politely edit the Word doc you
      have open (xlwings for Word)". Tue–Thu morning US-East; plan to be in the
      comments all day. The politeness / atomic-undo design section is what HN
      will actually engage with.
- [ ] **r/ClaudeAI** — MCPB angle: "drag this into Claude Desktop and ask it to
      tidy your document." Lead with the regularize GIF.
- [ ] **r/Python** — the "xlwings, but for Word" framing; check the sub's
      current self-promo/showcase rules first.
- [ ] **LinkedIn** — the office-worker angle: "AI that fixes your Word
      formatting, and one Ctrl-Z undoes it." GIFs autoplay in-feed, which is
      why this one is unusually worth it.
- [ ] Optional: r/Office365, the MCP/Anthropic Discords, X.

## Suggested sequence

Merge the registry PR → cut a release (this is what activates the registry
listing) → record the two GIFs → swap the README placeholders → submit to the
Anthropic directory and the awesome-lists that week → Show HN + Reddit +
LinkedIn together the following week, so launch traffic lands on the friendly
README.
