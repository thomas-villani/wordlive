"""Document export — the read mirror of `insert_markdown`.

`doc.to_markdown()` / `doc.to_html()` serialise a running document (or any
anchor's range) to clean Markdown / HTML, and `doc.read(budget=N)` produces a
token-budgeted, anchor-addressable digest of a whole document. All three share
one COM document-walk (`_walk_blocks`, in this module) that produces a flat list
of `Block` nodes; the emitters and the budgeted elide are **pure functions over
that node list**, so they are unit-testable off-Windows and Markdown/HTML provably
agree on structure.

The split mirrors the import side: `_markdown.py` is the pure block classifier and
`_anchors.insert_markdown` does the COM materialisation. Here `_walk_blocks` is the
COM half (lives lower in this file, populated in a later slice) and the `render_*`
functions are the pure half.

The node model is deliberately **flat** — a list of block-level `Block`s, each
carrying its inline `Span`s; a table is the only nested shape (`TableNode`). Every
node carries its `anchor_id` so the budgeted read can keep every `para:N` /
`heading:N` / `table:N` / `image:N` addressable after eliding.

Export is **lossy by design**, like the constrained-subset import: it round-trips
the dialect import speaks (`#`/`##`/`###`, `-`/`*` bullets, `1.` numbers,
`**bold**`/`*italic*`/`***both***`, blank-separated paragraphs) and additionally
emits richer reads import can't yet consume (deeper headings, GFM pipe tables,
`![alt](image:N)`, `[text](url)`).

This module adds no third-party dependency — both emitters write strings directly,
and the token estimate is a cheap char count.
"""

from __future__ import annotations

import re
from dataclasses import dataclass

# Block kinds (mirrors `_markdown`'s flat-constant style).
HEADING = "heading"
PARAGRAPH = "paragraph"
BULLET = "bullet"
NUMBER = "number"
TABLE = "table"
IMAGE = "image"  # a block-level (standalone-paragraph) image
BLANK = "blank"

# Token estimate: ~4 characters per token. Exposed as a constant so the budgeted
# read's live tuning pass can adjust it without touching logic.
_CHARS_PER_TOKEN = 4

# Characters that always carry Markdown meaning in inline content and so are
# always backslash-escaped (ported from all2md's `_escape_markdown`). `#` and `_`
# are handled context-sensitively below; `|` is escaped only inside table cells.
_ALWAYS_ESCAPE = frozenset("\\`*{}[]")


@dataclass(frozen=True)
class Span:
    """One inline span of a block plus its character formatting.

    `bold`/`italic`/`underline` are plain booleans (the read side knows the
    effective value, unlike the tri-state write `Run`). `href` makes the span a
    hyperlink (`[text](href)`); `image_ref` makes it an inline image
    (`![text](image_ref)`, with `text` the alt text) and wins over the rest.
    """

    text: str
    bold: bool = False
    italic: bool = False
    underline: bool = False
    href: str | None = None
    image_ref: str | None = None


@dataclass(frozen=True)
class TableNode:
    """A table as a rectangular grid of plain-text cells.

    Row 1 is the header. `alignments` is one of ``left``/``center``/``right``/
    ``None`` per column (``None`` → an undirected ``---`` separator). Cell text is
    plain in v1 (intra-cell emphasis is a documented deferral); the renderer
    escapes pipes and flattens newlines per GFM.
    """

    anchor_id: str
    header: tuple[str, ...]
    rows: tuple[tuple[str, ...], ...]
    alignments: tuple[str | None, ...]


@dataclass(frozen=True)
class Block:
    """One block-level node in document order.

    `kind` is one of the module constants. `level` is the heading level (1–9) for
    a heading; `list_level` is the 1-based nesting depth for a bullet/number item.
    `spans` carries inline content for text/heading/list blocks. `table` holds the
    `TableNode` for a table block; `image_ref`/`image_alt` describe an `IMAGE`
    block. `anchor_id` is the node's stable address (`heading:N`, `para:N`,
    `table:N`, `image:N`) — kept on every node so the budgeted read stays
    addressable after eliding.
    """

    kind: str
    anchor_id: str | None = None
    level: int | None = None
    spans: tuple[Span, ...] = ()
    list_level: int = 1
    table: TableNode | None = None
    image_ref: str | None = None
    image_alt: str | None = None
    # Word/paragraph count carried for the budgeted read's elision accounting;
    # unused by the plain emitters.
    word_count: int = 0


def _est_tokens(s: str) -> int:
    """A cheap char-based token estimate (~4 chars/token), floored at 1."""
    return max(1, len(s) // _CHARS_PER_TOKEN)


def _escape_inline(text: str) -> str:
    """Backslash-escape Markdown-special characters with context awareness.

    Always escapes ``\\`` ` ``*{}[]``; escapes ``#`` only at the very start (where
    it would start a heading) and ``_`` only at word boundaries (so ``snake_case``
    survives). Ported from all2md. The escapes ``\\*`` and ``\\\\`` round-trip back
    through `_runs.parse_markup` on re-import.
    """
    out: list[str] = []
    last = len(text) - 1
    for i, ch in enumerate(text):
        if ch in _ALWAYS_ESCAPE:
            out.append("\\")
            out.append(ch)
        elif ch == "#":
            out.append("\\#" if i == 0 else ch)
        elif ch == "_":
            mid_word = (i > 0 and text[i - 1].isalnum()) and (i < last and text[i + 1].isalnum())
            out.append(ch if mid_word else "\\_")
        else:
            out.append(ch)
    return "".join(out)


def _escape_table_cell(text: str) -> str:
    """Escape a GFM table cell: inline-escape, then guard pipes and flatten lines."""
    return _escape_inline(text).replace("|", "\\|").replace("\r", " ").replace("\n", " ")


def _render_spans(spans: tuple[Span, ...]) -> str:
    """Render inline spans to Markdown (emphasis, links, inline images)."""
    out: list[str] = []
    for s in spans:
        if s.image_ref is not None:
            out.append(f"![{_escape_inline(s.text)}]({s.image_ref})")
            continue
        text = _escape_inline(s.text)
        if s.href is not None:
            text = f"[{text}]({s.href})"
        # Underline has no Markdown markup in wordlive's dialect (write side only
        # reaches it via structured runs), so it is intentionally dropped here.
        if s.bold and s.italic:
            text = f"***{text}***"
        elif s.bold:
            text = f"**{text}**"
        elif s.italic:
            text = f"*{text}*"
        out.append(text)
    return "".join(out)


_ALIGN_MARKER = {"left": ":--", "center": ":-:", "right": "--:"}


def _render_table(t: TableNode) -> str:
    """Render a `TableNode` as a GFM pipe table (header + alignment row + body)."""
    width = len(t.header)

    def row(cells: tuple[str, ...]) -> str:
        padded = list(cells) + [""] * (width - len(cells))
        return "| " + " | ".join(_escape_table_cell(c) for c in padded[:width]) + " |"

    sep = "| " + " | ".join(_ALIGN_MARKER.get(a or "", "---") for a in t.alignments) + " |"
    lines = [row(t.header), sep]
    lines.extend(row(r) for r in t.rows)
    return "\n".join(lines)


def _render_block_md(b: Block) -> str:
    """Render one block to its Markdown chunk (may be multi-line for tables)."""
    if b.kind == HEADING:
        level = b.level or 1
        return "#" * level + " " + _render_spans(b.spans)
    if b.kind == PARAGRAPH:
        return _render_spans(b.spans)
    if b.kind == BULLET:
        return "  " * (b.list_level - 1) + "- " + _render_spans(b.spans)
    if b.kind == NUMBER:
        # Always "1." — consecutive ordered items renumber under GFM, and the
        # import side accepts any leading integer, so the subset round-trips.
        return "  " * (b.list_level - 1) + "1. " + _render_spans(b.spans)
    if b.kind == TABLE and b.table is not None:
        return _render_table(b.table)
    if b.kind == IMAGE:
        return f"![{_escape_inline(b.image_alt or '')}]({b.image_ref or ''})"
    return ""  # BLANK or unknown → nothing


_LIST_KINDS = frozenset({BULLET, NUMBER})


def _collapse(text: str) -> str:
    """Normalise line endings, collapse 3+ blank lines to one, strip trailing ws."""
    text = text.replace("\r\n", "\n").replace("\r", "\n")
    text = re.sub(r"\n{3,}", "\n\n", text)
    return text.rstrip()


def render_markdown(blocks: list[Block]) -> str:
    """Render a flat `Block` list to a Markdown document.

    Consecutive list items are joined tight (single newline); every other
    block-level element is separated by a blank line. Empty/`BLANK` blocks are
    dropped, and runs of blank lines collapse to one.
    """
    chunks: list[tuple[str, str]] = []  # (kind, rendered)
    for b in blocks:
        rendered = _render_block_md(b)
        if not rendered and b.kind != IMAGE:
            continue
        chunks.append((b.kind, rendered))

    parts: list[str] = []
    prev_kind: str | None = None
    for kind, rendered in chunks:
        if prev_kind is not None:
            tight = kind in _LIST_KINDS and prev_kind in _LIST_KINDS
            parts.append("\n" if tight else "\n\n")
        parts.append(rendered)
        prev_kind = kind
    return _collapse("".join(parts))
