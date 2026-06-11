"""Block-Markdown parser — pure, COM-free (like test_runs.py).

Covers `_markdown.parse_markdown` (line → Block classification) and
`_anchors._markdown_segments` (grouping Blocks into insert_block segments).
Inline-span correctness lives in test_runs.py; visible Word formatting in smoke.
"""

from __future__ import annotations

from wordlive._anchors import _markdown_segments
from wordlive._markdown import BULLET, HEADING, NORMAL, NUMBER, Block, parse_markdown


def _kinds(md: str) -> list[tuple[str, str, int | None]]:
    return [(b.kind, b.text, b.level) for b in parse_markdown(md)]


class TestParseMarkdown:
    def test_empty_is_no_blocks(self) -> None:
        assert parse_markdown("") == []
        assert parse_markdown("\n\n   \n") == []

    def test_headings_one_to_three(self) -> None:
        assert _kinds("# A\n## B\n### C") == [
            (HEADING, "A", 1),
            (HEADING, "B", 2),
            (HEADING, "C", 3),
        ]

    def test_four_hashes_is_plain_text(self) -> None:
        # Subset: only 1–3 hashes are headings; deeper falls through to a paragraph.
        assert _kinds("#### D") == [(NORMAL, "#### D", None)]

    def test_bullets_dash_and_star(self) -> None:
        assert _kinds("- one\n* two") == [(BULLET, "one", None), (BULLET, "two", None)]

    def test_numbered(self) -> None:
        assert _kinds("1. first\n2. second\n10. tenth") == [
            (NUMBER, "first", None),
            (NUMBER, "second", None),
            (NUMBER, "tenth", None),
        ]

    def test_blank_line_separates_paragraphs(self) -> None:
        assert _kinds("alpha\n\nbeta") == [(NORMAL, "alpha", None), (NORMAL, "beta", None)]

    def test_consecutive_plain_lines_join(self) -> None:
        assert _kinds("soft\nwrap\nlines") == [(NORMAL, "soft wrap lines", None)]

    def test_inline_markup_left_intact_for_run_parser(self) -> None:
        assert _kinds("**bold** lead") == [(NORMAL, "**bold** lead", None)]

    def test_leading_whitespace_is_insignificant(self) -> None:
        # Flat dialect — an indented marker is still a top-level item, not a nest.
        assert _kinds("  - indented") == [(BULLET, "indented", None)]

    def test_mixed_document_order_preserved(self) -> None:
        md = "# Title\n\nIntro para.\n\n- a\n- b\n\n1. x\n\nClosing."
        assert _kinds(md) == [
            (HEADING, "Title", 1),
            (NORMAL, "Intro para.", None),
            (BULLET, "a", None),
            (BULLET, "b", None),
            (NUMBER, "x", None),
            (NORMAL, "Closing.", None),
        ]

    def test_list_immediately_after_heading_flushes_no_phantom_paragraph(self) -> None:
        assert _kinds("# H\n- item") == [(HEADING, "H", 1), (BULLET, "item", None)]


class TestMarkdownSegments:
    def test_non_list_run_is_one_segment_no_list_type(self) -> None:
        blocks = [Block(HEADING, "H", 1), Block(NORMAL, "body")]
        segs = _markdown_segments(blocks)
        assert len(segs) == 1
        items, list_type = segs[0]
        assert list_type is None
        assert items == [{"text": "H", "style": "Heading 1"}, {"text": "body", "style": "Normal"}]

    def test_bullet_run_carries_bulleted_list_type(self) -> None:
        segs = _markdown_segments([Block(BULLET, "a"), Block(BULLET, "b")])
        assert len(segs) == 1
        items, list_type = segs[0]
        assert list_type == "bulleted"
        assert all(i["style"] == "List Bullet" for i in items)

    def test_number_run_carries_numbered_list_type(self) -> None:
        segs = _markdown_segments([Block(NUMBER, "a"), Block(NUMBER, "b")])
        items, list_type = segs[0]
        assert list_type == "numbered"
        assert all(i["style"] == "List Number" for i in items)

    def test_different_list_kinds_split_into_separate_segments(self) -> None:
        segs = _markdown_segments([Block(BULLET, "a"), Block(NUMBER, "b")])
        assert [s[1] for s in segs] == ["bulleted", "numbered"]

    def test_mixed_doc_groups_runs(self) -> None:
        blocks = [
            Block(HEADING, "H", 1),
            Block(NORMAL, "p"),
            Block(BULLET, "a"),
            Block(BULLET, "b"),
            Block(NORMAL, "q"),
        ]
        segs = _markdown_segments(blocks)
        assert [s[1] for s in segs] == [None, "bulleted", None]
