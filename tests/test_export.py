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
    _font_bool,
    _font_underline,
    build_digest,
    render_html,
    render_markdown,
)


class TestFontCoercion:
    def test_bold_italic_tristate(self) -> None:
        # Bold/Italic are tri-state: -1 True, 0 False, 9999999 (varies) → unset.
        assert _font_bool(-1) is True
        assert _font_bool(0) is False
        assert _font_bool(9999999) is False

    def test_underline_is_an_enum_not_a_bool(self) -> None:
        # WdUnderline: 0 = none, 1 = single, 3 = double, … ; 9999999 = varies.
        # The bug was routing this through `_font_bool` (== -1), which reported
        # every real underline as off.
        assert _font_underline(0) is False
        assert _font_underline(1) is True  # single
        assert _font_underline(3) is True  # double
        assert _font_underline(9999999) is False  # varies → unset
        assert _font_underline(None) is False


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


class TestRenderHtml:
    def test_heading_and_paragraph(self) -> None:
        blocks = [
            Block(HEADING, level=2, spans=_t(Span("Title"))),
            Block(PARAGRAPH, spans=_t(Span("a "), Span("b", bold=True))),
        ]
        assert render_html(blocks) == "<h2>Title</h2>\n<p>a <strong>b</strong></p>"

    def test_emphasis_underline_link_image(self) -> None:
        spans = _t(
            Span("u", underline=True),
            Span("i", italic=True),
            Span("here", href="https://x.test"),
            Span("logo", image_ref="image:1"),
        )
        html = render_html([Block(PARAGRAPH, spans=spans)])
        assert "<u>u</u>" in html
        assert "<em>i</em>" in html
        assert '<a href="https://x.test">here</a>' in html
        assert '<img src="image:1" alt="logo">' in html

    def test_html_escaping(self) -> None:
        html = render_html([Block(PARAGRAPH, spans=_t(Span("a < b & c")))])
        assert html == "<p>a &lt; b &amp; c</p>"

    def test_nested_lists(self) -> None:
        blocks = [
            Block(BULLET, spans=_t(Span("a"))),
            Block(BULLET, list_level=2, spans=_t(Span("a.i"))),
            Block(BULLET, spans=_t(Span("b"))),
        ]
        assert render_html(blocks) == "<ul><li>a<ul><li>a.i</li></ul></li><li>b</li></ul>"

    def test_ordered_list(self) -> None:
        blocks = [Block(NUMBER, spans=_t(Span("x"))), Block(NUMBER, spans=_t(Span("y")))]
        assert render_html(blocks) == "<ol><li>x</li><li>y</li></ol>"

    def test_table(self) -> None:
        t = TableNode(
            anchor_id="table:1",
            header=("A", "B"),
            rows=(("1", "2"),),
            alignments=(None, None),
        )
        html = render_html([Block(TABLE, anchor_id="table:1", table=t)])
        assert html == (
            "<table><thead><tr><th>A</th><th>B</th></tr></thead>"
            "<tbody><tr><td>1</td><td>2</td></tr></tbody></table>"
        )


def _para(text: str, anchor: str) -> Block:
    return Block(PARAGRAPH, anchor_id=anchor, spans=_t(Span(text)), word_count=len(text.split()))


def _heading(text: str, level: int, anchor: str) -> Block:
    return Block(HEADING, anchor_id=anchor, level=level, spans=_t(Span(text)))


class TestDigest:
    def test_headings_always_verbatim_with_anchor_tags(self) -> None:
        blocks = [
            _heading("Alpha", 1, "heading:1"),
            _para("filler " * 50, "para:2"),
            _heading("Beta", 1, "heading:3"),
        ]
        out = build_digest(blocks, budget=1)  # impossibly tight
        assert "# Alpha  <!-- heading:1 -->" in out
        assert "# Beta  <!-- heading:3 -->" in out

    def test_table_becomes_one_line_stub(self) -> None:
        t = TableNode("table:2", ("Q", "Rev"), (("1", "2"), ("3", "4")), (None, None))
        out = build_digest([Block(TABLE, anchor_id="table:2", table=t)], budget=500)
        assert "> table:2 — 3 rows × 2 cols: Q, Rev" in out
        assert "| Q |" not in out  # not the full table

    def test_body_overflow_elided_with_range_marker(self) -> None:
        blocks = [
            _heading("H", 1, "heading:1"),
            _para("lead " * 40, "para:2"),
            _para("more " * 40, "para:3"),
            _para("tail " * 40, "para:4"),
        ]
        out = build_digest(blocks, budget=80)
        # The floor keeps para:2; para:3-4 elide into one marker naming the range.
        assert "…(para:3–para:4, 80 words elided)…" in out

    def test_image_block_always_kept_addressable(self) -> None:
        blocks = [
            _heading("H", 1, "heading:1"),
            _para("filler " * 60, "para:2"),
            Block(PARAGRAPH, anchor_id="para:3", spans=_t(Span("chart", image_ref="image:1"))),
        ]
        out = build_digest(blocks, budget=20)
        assert "![chart](image:1)" in out  # survives elision

    def test_every_anchor_present(self) -> None:
        t = TableNode("table:4", ("A",), (("x",),), (None,))
        blocks = [
            _heading("One", 1, "heading:1"),
            _para("a " * 50, "para:2"),
            _heading("Two", 2, "heading:3"),
            Block(TABLE, anchor_id="table:4", table=t),
            _para("b " * 50, "para:5"),
        ]
        out = build_digest(blocks, budget=60)
        for anchor in ("heading:1", "heading:3", "table:4"):
            assert anchor in out
        # para:5 is either verbatim or named in an elision marker — addressable either way.
        assert "para:5" in out or "b b" in out

    def test_lead_snippets_capped_by_shared_budget(self) -> None:
        # Many small sections whose first body block always overflows its tiny
        # per-section share, so each can only appear via a lead snippet. Pre-fix,
        # *every* section emitted one regardless of budget (count == num); the
        # shared budget pool now caps the snippets well below one-per-section.
        num = 40
        blocks: list[Block] = []
        for i in range(1, num + 1):
            blocks.append(_heading(f"S{i}", 1, f"heading:{i}"))
            blocks.append(
                _para("alpha beta gamma delta epsilon zeta eta theta " * 5, f"para:{i}00")
            )
        out = build_digest(blocks, budget=400)
        fired = out.count("more words)…")
        assert fired < num  # not one-per-section (the pre-fix blowup)
        assert fired >= 1  # the pool still funds a few, so sections aren't all blank

    def test_tighter_budget_yields_smaller_output(self) -> None:
        blocks = [_heading("H", 1, "heading:1")] + [
            _para(f"section {i} " * 30, f"para:{i}") for i in range(2, 12)
        ]
        small = build_digest(blocks, budget=80)
        large = build_digest(blocks, budget=100_000)
        assert _est_tokens(small) < _est_tokens(large)

    def test_depth_weighting_favours_shallow_sections(self) -> None:
        big = "word " * 40
        blocks = [
            _heading("A", 1, "heading:1"),
            _para(big, "para:2"),
            _para(big, "para:3"),
            _para(big, "para:4"),
            _heading("B", 3, "heading:5"),
            _para(big, "para:6"),
            _para(big, "para:7"),
            _para(big, "para:8"),
        ]
        out = build_digest(blocks, budget=260)
        a_part, b_part = out.split("# B")
        # A (level 1, weight 1.0) keeps more verbatim body than B (level 3, weight 0.4).
        assert len(a_part) > len(b_part)

    def test_depth_cap_collapses_deep_sections(self) -> None:
        blocks = [
            _heading("Top", 1, "heading:1"),
            _para("kept " * 5, "para:2"),
            _heading("Deep", 2, "heading:3"),
            _para("hidden " * 5, "para:4"),
        ]
        out = build_digest(blocks, budget=10_000, depth=1)
        assert "# Deep" in out  # heading still shown (navigation)
        assert "hidden hidden" not in out  # its body is collapsed
        assert "para:4" in out  # but named in an elision marker
