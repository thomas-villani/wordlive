"""Export emitters — pure, COM-free (like test_markdown.py / test_runs.py).

Covers the Markdown renderer and its helpers in `_export`: node-list →
Markdown, context-aware escaping, GFM pipe tables, nested lists, inline
emphasis/links/images, and blank-line collapse. The COM document-walk
(`_walk_blocks`) and round-trip fidelity are exercised in later slices / smoke.
"""

from __future__ import annotations

from wordlive._export import (
    BULLET,
    HEADING,
    IMAGE,
    NUMBER,
    PARAGRAPH,
    TABLE,
    Block,
    Span,
    TableNode,
    _escape_inline,
    _escape_table_cell,
    _est_tokens,
    render_markdown,
)


def _t(*spans: Span) -> tuple[Span, ...]:
    return spans


class TestEscaping:
    def test_always_escaped_chars(self) -> None:
        assert _escape_inline(r"a*b`c[d]e{f}g\h") == r"a\*b\`c\[d\]e\{f\}g\\h"

    def test_hash_only_at_line_start(self) -> None:
        assert _escape_inline("# title") == "\\# title"
        assert _escape_inline("C# is fine") == "C# is fine"

    def test_underscore_mid_word_kept(self) -> None:
        assert _escape_inline("snake_case") == "snake_case"

    def test_underscore_at_boundary_escaped(self) -> None:
        assert _escape_inline("_emph_") == "\\_emph\\_"
        assert _escape_inline("a _ b") == "a \\_ b"

    def test_table_cell_escapes_pipe_and_flattens(self) -> None:
        assert _escape_table_cell("a | b\nc") == "a \\| b c"


class TestEstTokens:
    def test_four_chars_per_token_floor_one(self) -> None:
        assert _est_tokens("") == 1
        assert _est_tokens("abcd") == 1
        assert _est_tokens("a" * 40) == 10


class TestRenderInline:
    def test_bold_italic_both(self) -> None:
        md = render_markdown(
            [
                Block(
                    PARAGRAPH,
                    spans=_t(
                        Span("plain "),
                        Span("b", bold=True),
                        Span(" "),
                        Span("i", italic=True),
                        Span(" "),
                        Span("bi", bold=True, italic=True),
                    ),
                )
            ]
        )
        assert md == "plain **b** *i* ***bi***"

    def test_underline_dropped_in_markdown(self) -> None:
        md = render_markdown([Block(PARAGRAPH, spans=_t(Span("u", underline=True)))])
        assert md == "u"

    def test_link_and_inline_image(self) -> None:
        md = render_markdown(
            [
                Block(
                    PARAGRAPH,
                    spans=_t(
                        Span(
                            "see ",
                        ),
                        Span("here", href="https://x.test"),
                        Span(" and "),
                        Span("a chart", image_ref="image:3"),
                    ),
                )
            ]
        )
        assert md == "see [here](https://x.test) and ![a chart](image:3)"

    def test_emphasis_wraps_escaped_text(self) -> None:
        # Inner literal asterisk is escaped; the emphasis markers are not.
        md = render_markdown([Block(PARAGRAPH, spans=_t(Span("a*b", bold=True)))])
        assert md == r"**a\*b**"


class TestRenderBlocks:
    def test_heading_levels(self) -> None:
        blocks = [Block(HEADING, level=lvl, spans=_t(Span(f"H{lvl}"))) for lvl in (1, 3, 6)]
        assert render_markdown(blocks) == "# H1\n\n### H3\n\n###### H6"

    def test_tight_lists_and_nesting(self) -> None:
        blocks = [
            Block(BULLET, spans=_t(Span("a"))),
            Block(BULLET, list_level=2, spans=_t(Span("a.i"))),
            Block(BULLET, spans=_t(Span("b"))),
        ]
        assert render_markdown(blocks) == "- a\n  - a.i\n- b"

    def test_numbered_emits_one_dot(self) -> None:
        blocks = [Block(NUMBER, spans=_t(Span("x"))), Block(NUMBER, spans=_t(Span("y")))]
        assert render_markdown(blocks) == "1. x\n1. y"

    def test_blank_line_between_paragraphs(self) -> None:
        blocks = [Block(PARAGRAPH, spans=_t(Span("one"))), Block(PARAGRAPH, spans=_t(Span("two")))]
        assert render_markdown(blocks) == "one\n\ntwo"

    def test_block_image(self) -> None:
        blocks = [Block(IMAGE, image_ref="image:1", image_alt="logo")]
        assert render_markdown(blocks) == "![logo](image:1)"

    def test_empty_blocks_dropped(self) -> None:
        blocks = [
            Block(PARAGRAPH, spans=_t(Span("a"))),
            Block(PARAGRAPH, spans=()),
            Block(PARAGRAPH, spans=_t(Span("b"))),
        ]
        assert render_markdown(blocks) == "a\n\nb"


class TestRenderTable:
    def test_gfm_pipe_table_with_alignment(self) -> None:
        t = TableNode(
            anchor_id="table:1",
            header=("Name", "Qty"),
            rows=(("Widget", "3"), ("Gad|get", "10")),
            alignments=("left", "right"),
        )
        md = render_markdown([Block(TABLE, anchor_id="table:1", table=t)])
        assert md == ("| Name | Qty |\n| :-- | --: |\n| Widget | 3 |\n| Gad\\|get | 10 |")

    def test_ragged_row_padded(self) -> None:
        t = TableNode(
            anchor_id="table:2",
            header=("A", "B", "C"),
            rows=(("x",),),
            alignments=(None, None, None),
        )
        md = render_markdown([Block(TABLE, anchor_id="table:2", table=t)])
        assert md == "| A | B | C |\n| --- | --- | --- |\n| x |  |  |"


class TestCollapse:
    def test_collapses_excess_blank_lines(self) -> None:
        # Two image blocks separated by a blank line stay one blank line apart.
        blocks = [
            Block(IMAGE, image_ref="image:1", image_alt="a"),
            Block(IMAGE, image_ref="image:2", image_alt="b"),
        ]
        assert render_markdown(blocks) == "![a](image:1)\n\n![b](image:2)"
