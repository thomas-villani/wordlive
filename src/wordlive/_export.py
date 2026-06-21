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
from dataclasses import dataclass, replace
from typing import TYPE_CHECKING, Any

if TYPE_CHECKING:
    from ._anchors import Anchor
    from ._document import Document

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
    """Render inline spans to Markdown (emphasis, links, inline images).

    Surrounding whitespace is hoisted *outside* the emphasis/link markers — a
    span coalesced from `Range.Words` carries trailing spaces (``"bold "``), and
    ``**bold **`` is not valid CommonMark emphasis (the closing ``**`` must follow
    non-whitespace).
    """
    out: list[str] = []
    for s in spans:
        if s.image_ref is not None:
            out.append(f"![{_escape_inline(s.text)}]({s.image_ref})")
            continue
        decorated = s.href is not None or s.bold or s.italic
        if not decorated:
            out.append(_escape_inline(s.text))
            continue
        # Split leading/trailing whitespace off so the markers wrap only the core.
        core = s.text.strip()
        if not core:
            out.append(s.text)  # whitespace-only span — nothing to decorate
            continue
        lead = s.text[: len(s.text) - len(s.text.lstrip())]
        trail = s.text[len(s.text.rstrip()) :]
        text = _escape_inline(core)
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
        out.append(f"{lead}{text}{trail}")
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


# ---------------------------------------------------------------------------
# The COM document-walk — the one impure pass the emitters share.
# ---------------------------------------------------------------------------

# WdListType values (live-probed 2026-06-21): 0 = no list, 2 = bullet,
# 6 = picture bullet; 1/3/4/5 are the numbered variants.
_BULLET_LIST_TYPES = frozenset({2, 6})
_NUMBER_LIST_TYPES = frozenset({1, 3, 4, 5})

# Word reports OutlineLevel 10 for every non-heading paragraph; 1–9 are headings.
_BODY_OUTLINE_LEVEL = 10
_HEADING_STYLE = re.compile(r"Heading (\d+)")

# Structural marks that terminate a paragraph / cell — never real content.
_PARA_MARKS = "\r\n\x07\x0c"

# The placeholder a font flag carries when it varies across a run (wdUndefined).
_WD_UNDEFINED = 9999999


def _font_bool(value: Any) -> bool:
    """Coerce a Word ``Font.Bold``/``Italic``/``Underline`` read to a plain bool.

    Word returns ``-1`` (True) / ``0`` (False); ``9999999`` (wdUndefined) means
    the property *varies* within the word — treated as unset so a mixed word
    degrades to plain text rather than guessing.
    """
    return value == -1


def _href_of(hyperlink: Any) -> str | None:
    """The Markdown target for a hyperlink — external address or ``#bookmark``."""
    addr = str(hyperlink.Address or "")
    if addr:
        return addr
    sub = str(hyperlink.SubAddress or "")
    return f"#{sub}" if sub else None


def _coalesce_spans(raw: list[Span]) -> tuple[Span, ...]:
    """Merge adjacent spans that share formatting; inline images stay standalone."""
    out: list[Span] = []
    for s in raw:
        if (
            out
            and s.image_ref is None
            and out[-1].image_ref is None
            and (s.bold, s.italic, s.underline, s.href)
            == (out[-1].bold, out[-1].italic, out[-1].underline, out[-1].href)
        ):
            out[-1] = replace(out[-1], text=out[-1].text + s.text)
        else:
            out.append(s)
    return tuple(out)


def _paragraph_spans(
    para_range: Any, image_by_start: dict[int, tuple[int, str]]
) -> tuple[Span, ...]:
    """Extract a paragraph's inline content as formatted `Span`s.

    Walks `Range.Words`, reading per-word bold/italic/underline and tagging each
    word that falls inside a hyperlink's range (live-probed: a link's display
    words report a `Start` within the link `Range`, the hidden field code aside)
    or sits at an inline shape's position (an `![alt](image:N)`). Adjacent words
    with identical formatting coalesce into one span.
    """
    links: list[tuple[int, int, str]] = []
    for hl in para_range.Hyperlinks:
        href = _href_of(hl)
        if href is None:
            continue
        r = hl.Range
        links.append((int(r.Start), int(r.End), href))

    raw: list[Span] = []
    for word in para_range.Words:
        text = str(word.Text or "")
        stripped = text.rstrip(_PARA_MARKS)
        if not stripped:
            continue  # a structural-mark-only word (paragraph/cell/page mark)
        start = int(word.Start)
        if start in image_by_start:
            idx, alt = image_by_start[start]
            raw.append(Span(text=alt, image_ref=f"image:{idx}"))
            continue
        font = word.Font
        href = next((h for (s, e, h) in links if s <= start < e), None)
        raw.append(
            Span(
                text=stripped,
                bold=_font_bool(font.Bold),
                italic=_font_bool(font.Italic),
                underline=_font_bool(font.Underline),
                href=href,
            )
        )
    return _coalesce_spans(raw)


def _heading_level(outline_level: int, style: str | None) -> int | None:
    """The heading level for a paragraph, or ``None`` if it is body text.

    Prefers Word's `OutlineLevel` (1–9 = the anchor model's heading space); falls
    back to a ``Heading N`` style name for a renamed-outline paragraph.
    """
    if outline_level < _BODY_OUTLINE_LEVEL:
        return outline_level
    if style and (m := _HEADING_STYLE.fullmatch(style)):
        return int(m.group(1))
    return None


def _table_block(doc: Document, index: int) -> Block:
    """Build a `TABLE` `Block` from the 1-based table `index`."""
    table = doc.tables[index]
    grid = table.grid()
    header = tuple(grid[0]) if grid else ()
    rows = tuple(tuple(r) for r in grid[1:])
    alignments: list[str | None] = []
    for c in range(1, len(header) + 1):
        try:
            align = int(table.com.Cell(1, c).Range.ParagraphFormat.Alignment)
        except Exception:
            align = 0
        alignments.append({1: "center", 2: "right"}.get(align))
    node = TableNode(
        anchor_id=f"table:{index}",
        header=header,
        rows=rows,
        alignments=tuple(alignments),
    )
    return Block(TABLE, anchor_id=f"table:{index}", table=node)


def _list_kind(list_type: int) -> str | None:
    if list_type in _BULLET_LIST_TYPES:
        return BULLET
    if list_type in _NUMBER_LIST_TYPES:
        return NUMBER
    return None


def walk_blocks(doc: Document, within: str | Anchor | None = None) -> list[Block]:
    """Walk a document (or one anchor's range) into a flat `Block` list.

    The single COM pass shared by `to_markdown` / `to_html` / the budgeted
    `read`. `within` is an anchor id or `Anchor` whose range scopes the walk
    (``None`` = the whole document); a table is emitted as one unit at the
    position of its first cell-paragraph (in-table paragraphs are skipped via a
    range-interval test, not a per-paragraph COM call).
    """
    doc_com = doc.com
    lo: int | None = None
    hi: int | None = None
    if within is not None:
        rng = doc.anchor_by_id(within).com if isinstance(within, str) else within.com
        lo, hi = int(rng.Start), int(rng.End)

    # Precompute table intervals (start, end, 1-based index), and inline-shape
    # positions → (image:N index, alt text), so the per-paragraph loop makes no
    # extra COM calls for membership / image lookup.
    tables: list[tuple[int, int, int]] = []
    for i, t in enumerate(doc_com.Tables, start=1):
        tr = t.Range
        tables.append((int(tr.Start), int(tr.End), i))
    image_by_start: dict[int, tuple[int, str]] = {}
    shapes = doc_com.InlineShapes
    for i in range(1, int(shapes.Count) + 1):
        sh = shapes.Item(i)
        try:
            image_by_start[int(sh.Range.Start)] = (i, str(sh.AlternativeText or ""))
        except Exception:
            continue

    blocks: list[Block] = []
    seen_tables: set[int] = set()
    for idx, para in enumerate(doc_com.Paragraphs, start=1):
        pr = para.Range
        start = int(pr.Start)
        if lo is not None and hi is not None and not (lo <= start < hi):
            continue
        in_table = next((ti for (ts, te, ti) in tables if ts <= start < te), None)
        if in_table is not None:
            if in_table not in seen_tables:
                seen_tables.add(in_table)
                blocks.append(_table_block(doc, in_table))
            continue
        try:
            outline_level = int(para.OutlineLevel)
        except Exception:
            outline_level = _BODY_OUTLINE_LEVEL
        try:
            style: str | None = str(pr.Style.NameLocal)
        except Exception:
            style = None
        spans = _paragraph_spans(pr, image_by_start)
        text = "".join(s.text for s in spans)
        words = len(text.split())

        level = _heading_level(outline_level, style)
        if level is not None:
            blocks.append(
                Block(
                    HEADING, anchor_id=f"heading:{idx}", level=level, spans=spans, word_count=words
                )
            )
            continue
        kind = _list_kind(int(pr.ListFormat.ListType))
        if kind is not None:
            blocks.append(
                Block(
                    kind,
                    anchor_id=f"para:{idx}",
                    list_level=max(1, int(pr.ListFormat.ListLevelNumber)),
                    spans=spans,
                    word_count=words,
                )
            )
            continue
        blocks.append(Block(PARAGRAPH, anchor_id=f"para:{idx}", spans=spans, word_count=words))
    return blocks
