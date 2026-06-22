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


def _md_target(target: str) -> str:
    """A Markdown link/image destination safe for the bare ``(...)`` form.

    A target with spaces or parentheses (SharePoint URLs, ``#bookmark`` subadresses)
    breaks an unwrapped ``](url)``; CommonMark's angle-bracket form ``<url>`` allows
    both, so wrap when needed and escape the only chars it can't carry literally
    (``<`` / ``>`` / line breaks)."""
    if not target:
        return target
    if set(target) & {" ", "(", ")", "<", ">", "\n", "\r"}:
        inner = (
            target.replace("<", "%3C").replace(">", "%3E").replace("\r", "").replace("\n", "%0A")
        )
        return f"<{inner}>"
    return target


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
            out.append(f"![{_escape_inline(s.text)}]({_md_target(s.image_ref)})")
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
            text = f"[{text}]({_md_target(s.href)})"
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
        return f"![{_escape_inline(b.image_alt or '')}]({_md_target(b.image_ref or '')})"
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
# HTML emitter — the second pure renderer over the shared node list.
# ---------------------------------------------------------------------------


def _html_escape(text: str) -> str:
    return text.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def _html_attr(text: str) -> str:
    return _html_escape(text).replace('"', "&quot;")


def _render_spans_html(spans: tuple[Span, ...]) -> str:
    out: list[str] = []
    for s in spans:
        if s.image_ref is not None:
            out.append(f'<img src="{_html_attr(s.image_ref)}" alt="{_html_attr(s.text)}">')
            continue
        text = _html_escape(s.text)
        if s.underline:  # HTML keeps underline, unlike the Markdown dialect
            text = f"<u>{text}</u>"
        if s.italic:
            text = f"<em>{text}</em>"
        if s.bold:
            text = f"<strong>{text}</strong>"
        if s.href is not None:
            text = f'<a href="{_html_attr(s.href)}">{text}</a>'
        out.append(text)
    return "".join(out)


def _render_table_html(t: TableNode) -> str:
    width = len(t.header)

    def cells(row: tuple[str, ...], tag: str) -> str:
        padded = list(row) + [""] * (width - len(row))
        return "".join(f"<{tag}>{_html_escape(c)}</{tag}>" for c in padded[:width])

    head = f"<thead><tr>{cells(t.header, 'th')}</tr></thead>"
    body = "".join(f"<tr>{cells(r, 'td')}</tr>" for r in t.rows)
    return f"<table>{head}<tbody>{body}</tbody></table>"


def _render_list_html(items: list[Block]) -> str:
    """Render a run of consecutive list items as nested ``<ul>``/``<ol>``.

    A deeper item nests inside the currently-open ``<li>`` (proper HTML nesting).
    Mixed bullet/number at the *same* depth keep the level's first tag — switching
    list type mid-level is a documented v1 simplification.
    """
    out: list[str] = []
    stack: list[str] = []
    for b in items:
        level = max(1, b.list_level)
        tag = "ul" if b.kind == BULLET else "ol"
        if level > len(stack):
            while len(stack) < level:
                out.append(f"<{tag}>")
                stack.append(tag)
                if len(stack) < level:
                    out.append("<li>")  # descend through a skipped level
        else:
            out.append("</li>")
            while len(stack) > level:
                out.append(f"</{stack.pop()}>")
                out.append("</li>")
        out.append(f"<li>{_render_spans_html(b.spans)}")
    if stack:
        out.append("</li>")
    while stack:
        out.append(f"</{stack.pop()}>")
        if stack:
            out.append("</li>")
    return "".join(out)


def _render_block_html(b: Block) -> str:
    if b.kind == HEADING:
        level = b.level or 1
        return f"<h{level}>{_render_spans_html(b.spans)}</h{level}>"
    if b.kind == PARAGRAPH:
        inner = _render_spans_html(b.spans)
        return f"<p>{inner}</p>" if inner else ""
    if b.kind == TABLE and b.table is not None:
        return _render_table_html(b.table)
    if b.kind == IMAGE:
        src = _html_attr(b.image_ref or "")
        return f'<p><img src="{src}" alt="{_html_attr(b.image_alt or "")}"></p>'
    return ""


def render_html(blocks: list[Block]) -> str:
    """Render a flat `Block` list to an HTML fragment (shares the node list with
    `render_markdown`, so the two provably agree on structure)."""
    out: list[str] = []
    i = 0
    n = len(blocks)
    while i < n:
        if blocks[i].kind in _LIST_KINDS:
            j = i
            while j < n and blocks[j].kind in _LIST_KINDS:
                j += 1
            out.append(_render_list_html(blocks[i:j]))
            i = j
            continue
        rendered = _render_block_html(blocks[i])
        if rendered:
            out.append(rendered)
        i += 1
    return "\n".join(out)


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
    """Coerce a Word ``Font.Bold``/``Italic`` read to a plain bool.

    Word returns ``-1`` (True) / ``0`` (False); ``9999999`` (wdUndefined) means
    the property *varies* within the word — treated as unset so a mixed word
    degrades to plain text rather than guessing.
    """
    return value == -1


def _font_underline(value: Any) -> bool:
    """Coerce a Word ``Font.Underline`` read to a plain bool.

    Unlike ``Bold``/``Italic``, ``Underline`` is a ``WdUnderline`` enum — *not* a
    tri-state boolean: ``0`` is no underline, ``9999999`` (wdUndefined) means it
    varies within the word, and every other value (1 = single, 3 = double, …) is
    underlined. So it must not go through `_font_bool` (which tests ``== -1`` and
    would report every underline style as *off*).
    """
    try:
        v = int(value)
    except (TypeError, ValueError):
        return False
    return v not in (0, _WD_UNDEFINED)


def _href_of(hyperlink: Any) -> str | None:
    """The Markdown target for a hyperlink — external address or ``#bookmark``."""
    addr = str(hyperlink.Address or "")
    if addr:
        return addr
    sub = str(hyperlink.SubAddress or "")
    return f"#{sub}" if sub else None


def _coalesce_spans(raw: list[Span]) -> tuple[Span, ...]:
    """Merge adjacent spans that share formatting; inline images stay standalone.

    Adjacent words of one hyperlink also merge **by href alone**, even if their
    incidental bold/italic/underline reads differ: a hyperlink is stored as a
    field, so its first display word's range overlaps the hidden field code and
    reads ragged formatting — without href-merging a multi-word link splits into
    `[one](url) [two](url) …` instead of one `[one two three](url)`.
    """
    out: list[Span] = []
    for s in raw:
        if out and s.image_ref is None and out[-1].image_ref is None:
            prev = out[-1]
            same_fmt = (s.bold, s.italic, s.underline, s.href) == (
                prev.bold,
                prev.italic,
                prev.underline,
                prev.href,
            )
            same_link = s.href is not None and s.href == prev.href
            if same_fmt or same_link:
                out[-1] = replace(prev, text=prev.text + s.text)
                continue
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
        end = int(word.End)
        # Tag by range *overlap*, not a point test on `start`: a hyperlink's
        # display range and a word's range don't always share an edge (field
        # codes, trailing-mark trimming), so a point test can split a multi-word
        # link or drop its last word.
        href = next((h for (s, e, h) in links if start < e and end > s), None)
        raw.append(
            Span(
                text=stripped,
                bold=_font_bool(font.Bold),
                italic=_font_bool(font.Italic),
                underline=_font_underline(font.Underline),
                href=href,
            )
        )
    return _coalesce_spans(raw)


def _heading_level(
    outline_level: int, style: str | None, heading_styles: dict[str, int] | None = None
) -> int | None:
    """The heading level for a paragraph, or ``None`` if it is body text.

    Prefers Word's `OutlineLevel` (1–9 = the anchor model's heading space); falls
    back to a heading *style* name. `heading_styles` maps this document's
    localized built-in heading names (``"Heading 1"``, ``"Überschrift 1"``, …) to
    their level, so the fallback works on non-English Word; the English
    ``Heading N`` regex is a last resort for renamed-outline paragraphs.
    """
    if outline_level < _BODY_OUTLINE_LEVEL:
        return outline_level
    if style:
        if heading_styles and style in heading_styles:
            return heading_styles[style]
        if m := _HEADING_STYLE.fullmatch(style):
            return int(m.group(1))
    return None


def _builtin_heading_styles(doc_com: Any) -> dict[str, int]:
    """Map this document's localized built-in heading style names to their level.

    `wdStyleHeadingN` is the magic constant ``-(N + 1)``; reading each style's
    `NameLocal` yields the language-specific name Word stamps on paragraphs, so a
    name-based heading fallback stays correct on a non-English install."""
    out: dict[str, int] = {}
    for n in range(1, 10):
        try:
            out[str(doc_com.Styles(-(n + 1)).NameLocal)] = n
        except Exception:
            continue
    return out


def _table_block(doc: Document, index: int, image_by_start: dict[int, tuple[int, str]]) -> Block:
    """Build a `TABLE` `Block` from the 1-based table `index`.

    Cell text is plain (intra-cell emphasis is a documented deferral), but any
    inline image inside a cell is appended as an `![alt](image:N)` token so the
    `image:N` anchor stays addressable — the digest's "every image stays
    referenceable" invariant must hold inside tables too, not just in body text.
    """
    table = doc.tables[index]
    grid = table.grid()
    images = sorted(image_by_start.items())
    if images:
        # Walk the same *physical* cells `grid()` did (merged-safe) and append
        # any inline-image refs whose start falls inside the cell's range.
        tcom = table.com
        for r in range(1, len(grid) + 1):
            cells = tcom.Rows(r).Cells
            for c in range(1, len(grid[r - 1]) + 1):
                try:
                    rng = cells.Item(c).Range
                    cs, ce = int(rng.Start), int(rng.End)
                except Exception:
                    continue
                refs = [f"![{alt}](image:{idx})" for pos, (idx, alt) in images if cs <= pos < ce]
                if refs:
                    cur = grid[r - 1][c - 1]
                    joined = " ".join(refs)
                    grid[r - 1][c - 1] = f"{cur} {joined}".strip() if cur else joined
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
    heading_styles = _builtin_heading_styles(doc_com)

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
                blocks.append(_table_block(doc, in_table, image_by_start))
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

        level = _heading_level(outline_level, style, heading_styles)
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


# ---------------------------------------------------------------------------
# Budgeted digest — the elided whole-document read (doc.read(budget=N)).
# ---------------------------------------------------------------------------

# How much verbatim body budget a section gets, scaled by its enclosing heading
# depth — a flat document spends evenly, a deep one stays top-heavy. Tunable; the
# live tuning pass adjusts these (and `_CHARS_PER_TOKEN`) without touching logic.
_DEPTH_WEIGHTS = {1: 1.0, 2: 0.7, 3: 0.4}
_DEEP_WEIGHT = 0.2
# When a section's budget share won't fit even its first block, a bounded lead
# snippet (first sentence, capped at this many words) is kept so no section
# vanishes — this is what lets `budget` bound the output regardless of how many
# sections a document has.
_LEAD_WORDS = 20


def _depth_weight(level: int) -> float:
    return _DEPTH_WEIGHTS.get(level, _DEEP_WEIGHT)


def _lead_snippet(line: str) -> tuple[str, bool]:
    """A bounded lead of a rendered body line — its first sentence, ≤ `_LEAD_WORDS`.

    Returns the snippet and whether anything was dropped (a later sentence exists
    or the first sentence was word-capped) so the caller only marks elided text
    when text was actually elided.
    """
    parts = re.split(r"(?<=[.!?])\s+", line, maxsplit=1)
    first = parts[0]
    words = first.split()
    if len(words) > _LEAD_WORDS:
        return " ".join(words[:_LEAD_WORDS]) + " …", True
    return first, len(parts) > 1


def _plain_lead_words(b: Block) -> int:
    """Words of `b`'s first sentence, counted off the same plain-text source as
    `word_count` — so `word_count − this` is a consistent "words shown" figure."""
    plain = "".join(s.text for s in b.spans)
    first = re.split(r"(?<=[.!?])\s+", plain, maxsplit=1)[0]
    return min(len(first.split()), _LEAD_WORDS)


def _anchor_comment(anchor_id: str | None) -> str:
    return f"  <!-- {anchor_id} -->" if anchor_id else ""


def _heading_digest_line(b: Block) -> str:
    return "#" * (b.level or 1) + " " + _render_spans(b.spans) + _anchor_comment(b.anchor_id)


def _table_stub(b: Block) -> str:
    """A one-line table shape: dimensions + header labels, kept addressable."""
    t = b.table
    if t is None:
        return ""
    rows = len(t.rows) + 1  # body rows + the header row
    cols = len(t.header)
    labels = ", ".join(h for h in t.header if h)
    suffix = f": {labels}" if labels else ""
    return f"> {b.anchor_id} — {rows} rows × {cols} cols{suffix}"


def _has_image(b: Block) -> bool:
    return any(s.image_ref is not None for s in b.spans)


def _render_body_segment(
    body: list[Block], share: float, *, suppress: bool, budget_left: list[float]
) -> str:
    """Render a section's body to its token `share`, eliding the overflow.

    Keeps a leading run verbatim until `share` is spent (plus any image-bearing
    block, so `image:N` stays addressable); when even the section's first block
    overflows, a bounded lead snippet of it is kept (so the section never
    vanishes) with a `…(para:N, M more words)…` truncation marker. Each run of
    fully-dropped blocks collapses to one `…(para:A–para:B, N words elided)…`
    marker. `suppress` (a depth-capped section) elides everything into markers.

    `budget_left` is a one-element ``[tokens]`` cell shared across every section
    (the global remaining body budget). Both verbatim keeps and lead snippets
    draw it down, and the lead-snippet path only fires while it's positive — so
    the digest cannot grow without bound on a document with very many small
    sections (the snippet count is capped by the budget, not the section count).
    """
    out: list[str] = []
    pending: list[Block] = []
    used = 0.0
    kept = 0

    def flush() -> None:
        if not pending:
            return
        first, last = pending[0].anchor_id, pending[-1].anchor_id
        rng = first if first == last else f"{first}–{last}"
        words = sum(b.word_count for b in pending)
        out.append(f"…({rng}, {words} words elided)…")
        pending.clear()

    for b in body:
        line = _render_block_md(b)
        if not line:
            continue
        cost = _est_tokens(line)
        if not suppress and (_has_image(b) or used + cost <= share):
            flush()
            out.append(line)
            used += cost
            budget_left[0] -= cost
            kept += 1
        elif not suppress and kept == 0 and not pending and budget_left[0] > 0:
            # The section's first block overflows: keep a bounded lead snippet so
            # the section stays useful — but only while the global body budget
            # isn't spent, so `budget` actually bounds the whole digest.
            snippet, truncated = _lead_snippet(line)
            out.append(snippet)
            scost = _est_tokens(snippet)
            used += scost
            budget_left[0] -= scost
            kept += 1
            # Only mark elided text when the snippet actually dropped some, and
            # count the remainder off the same plain-text source as word_count so
            # a fully-shown paragraph never emits a spurious "N more words".
            rest = b.word_count - _plain_lead_words(b) if truncated else 0
            if rest > 0:
                out.append(f"…({b.anchor_id}, {rest} more words)…")
        else:
            pending.append(b)
    flush()
    return "\n\n".join(out)


def build_digest(blocks: list[Block], *, budget: int = 6000, depth: int | None = None) -> str:
    """Render a `Block` list to a token-budgeted, anchor-addressable digest.

    Headings are always emitted verbatim (the navigation spine, each tagged with
    its `<!-- heading:N -->` anchor) and tables become one-line shape stubs; the
    remaining budget is spread across each section's body, weighted by heading
    depth, and the overflow is elided to markers that still name the `para:` range.
    `depth` caps how deep a section keeps any body (deeper sections collapse to a
    single elision marker). The whole document stays addressable after eliding.

    `budget` bounds the **body** content: the heading spine and table stubs are
    the fixed navigation backbone (always kept, so a caller can see the document's
    shape), and `budget` governs how much paragraph text rides on top — both the
    verbatim keeps and the per-section lead snippets draw from one shared pool, so
    the body can't grow without bound as section count rises. (A document with
    more headings than `budget` allows is dominated by its spine by design.)
    """
    # Partition into ordered segments, tracking each body run's enclosing level.
    segments: list[dict[str, Any]] = []
    body: list[Block] = []
    level = 1

    def flush_body() -> None:
        nonlocal body
        if body:
            segments.append({"type": "body", "blocks": body, "level": level})
            body = []

    for b in blocks:
        if b.kind == HEADING:
            flush_body()
            level = b.level or 1
            segments.append({"type": "heading", "block": b})
        elif b.kind == TABLE:
            flush_body()
            segments.append({"type": "table", "block": b})
        else:
            body.append(b)
    flush_body()

    # Reserve the fixed cost (headings + table stubs) off the top.
    fixed = sum(
        _est_tokens(_heading_digest_line(s["block"]))
        if s["type"] == "heading"
        else _est_tokens(_table_stub(s["block"]))
        for s in segments
        if s["type"] in ("heading", "table")
    )
    body_budget = max(0, budget - fixed)

    # Spread the rest across body segments by depth weight.
    body_segs = [s for s in segments if s["type"] == "body"]
    total_weight = sum(_depth_weight(s["level"]) for s in body_segs) or 1.0
    for s in body_segs:
        s["share"] = body_budget * _depth_weight(s["level"]) / total_weight

    lines: list[str] = []
    budget_left = [float(body_budget)]  # shared pool drawn by keeps + snippets
    for s in segments:
        if s["type"] == "heading":
            lines.append(_heading_digest_line(s["block"]))
        elif s["type"] == "table":
            lines.append(_table_stub(s["block"]))
        else:
            suppress = depth is not None and s["level"] > depth
            rendered = _render_body_segment(
                s["blocks"], s["share"], suppress=suppress, budget_left=budget_left
            )
            if rendered:
                lines.append(rendered)
    return _collapse("\n\n".join(line for line in lines if line))
