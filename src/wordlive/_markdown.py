"""Block-level Markdown — a deliberately *tiny* subset mapped to Word structure.

`insert_markdown` lets an agent hand wordlive a chunk of Markdown and get back
real Word paragraphs, headings, and lists in one op. This module does the
COM-free part: classify each source line into a `Block`. The actual insertion
(and the inline `**bold**`/`*italic*` parsing, which is reused as-is from
`_runs.parse_markup`) lives in `_anchors.insert_markdown`.

The dialect is **a documented subset, not CommonMark** — Word is not Markdown.
Supported, on its own line:

- ``# `` / ``## `` / ``### `` → `Heading 1` / `Heading 2` / `Heading 3`
  (1–3 hashes only; ``#### `` and deeper fall through to a plain paragraph).
- ``- `` / ``* `` → a bulleted list item.
- ``1. `` (one-or-more digits, a dot, a space) → a numbered list item.
- a blank line separates paragraphs; consecutive plain lines join into one
  `Normal` paragraph (a soft wrap, CommonMark-style).
- inline spans (`**bold**`, `*italic*`, `***both***`) are left intact in the
  block text — `insert_block` runs them through `_runs.parse_markup` downstream.

Explicitly **out of scope** in v1 (deferred until asked): code fences, nested /
mixed lists, block quotes, tables, links, images. Anything that isn't a heading,
list item, or blank line is plain paragraph text — so a stray ``> quote`` or
```` ``` ```` fence lands as literal text rather than being silently dropped.

This module is pure and COM-free so the parser is unit-testable off-Windows,
mirroring `_runs.py`.
"""

from __future__ import annotations

import re
from dataclasses import dataclass

# Block kinds. `heading` carries a 1–3 `level`; the rest leave it `None`.
HEADING = "heading"
BULLET = "bullet"
NUMBER = "number"
NORMAL = "normal"

# Marker patterns, matched after stripping leading whitespace (the subset is
# flat — there are no nested list levels in v1, so indentation is insignificant).
_HEADING = re.compile(r"(#{1,3}) +(.*)$")
_BULLET = re.compile(r"[-*] +(.*)$")
_NUMBER = re.compile(r"\d+\. +(.*)$")


@dataclass(frozen=True)
class Block:
    """One source-derived block in document order.

    `kind` is `heading`/`bullet`/`number`/`normal`; `text` is the block's content
    with inline markup left intact for the run parser. `level` is the heading
    level (1–3) for a heading, `None` otherwise.
    """

    kind: str
    text: str
    level: int | None = None


def parse_markdown(md: str) -> list[Block]:
    """Classify `md` into a flat list of `Block`s in document order.

    A blank line flushes the current plain-text paragraph; consecutive plain
    lines join with a single space. Headings and list items each flush the
    pending paragraph and emit their own block. Returns ``[]`` for empty input.
    """
    blocks: list[Block] = []
    para: list[str] = []

    def flush() -> None:
        if para:
            blocks.append(Block(NORMAL, " ".join(para)))
            para.clear()

    for raw in md.splitlines():
        line = raw.strip()
        if not line:
            flush()
            continue
        if m := _HEADING.fullmatch(line):
            flush()
            blocks.append(Block(HEADING, m.group(2), level=len(m.group(1))))
        elif m := _BULLET.fullmatch(line):
            flush()
            blocks.append(Block(BULLET, m.group(1)))
        elif m := _NUMBER.fullmatch(line):
            flush()
            blocks.append(Block(NUMBER, m.group(1)))
        else:
            para.append(line)
    flush()
    return blocks
