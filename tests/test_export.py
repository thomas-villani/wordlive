"""Export emitters — pure, COM-free (like test_markdown.py / test_runs.py).

Covers the Markdown renderer and its helpers in `_export`: node-list →
Markdown, context-aware escaping, GFM pipe tables, nested lists, inline
emphasis/links/images, and blank-line collapse. The COM document-walk
(`_walk_blocks`) and round-trip fidelity are exercised in later slices / smoke.
"""

from __future__ import annotations

import pytest

from wordlive._export import (
    _ALWAYS_ESCAPE,
    BULLET,
    HEADING,
    IMAGE,
    NUMBER,
    PARAGRAPH,
    TABLE,
    Block,
    Span,
    TableNode,
    _code_span,
    _escape_inline,
    _escape_table_cell,
    _est_tokens,
    _font_bool,
    _font_monospace,
    _font_underline,
    _heading_level,
    _md_target,
    _render_body_segment,
    _render_spans,
    build_digest,
    render_html,
    render_markdown,
)
from wordlive._runs import _ESCAPABLE, parse_markup, runs_to_text


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


class TestEscapeRoundTrip:
    """`_escape_inline` (write-out) and `parse_markup` (read-in) must be inverses.

    When they drifted, every `to_markdown` -> `insert_markdown` cycle added one
    backslash per special character, without bound — silently corrupting any
    read-modify-write of a document containing code, paths, or brackets.
    """

    CORPUS = [
        "use `attach()` now",
        "call `_app.py` and `x[0]`",
        "a*b`c[d]e{f}g\\h",
        "snake_case stays",
        "_emph_ at boundaries",
        "# title",
        "C# is fine",
        "literal \\* star",
        "C:\\Users\\thomas",
        "{braces} and [brackets]",
        "",
        "plain text",
    ]

    @pytest.mark.parametrize("src", CORPUS)
    def test_escape_then_parse_is_identity(self, src: str) -> None:
        assert runs_to_text(parse_markup(_escape_inline(src))) == src

    @pytest.mark.parametrize("src", CORPUS)
    def test_round_trip_is_a_fixed_point_not_a_ratchet(self, src: str) -> None:
        # Three cycles must not grow the text by even one backslash.
        text = src
        for _ in range(3):
            text = runs_to_text(parse_markup(_escape_inline(text)))
        assert text == src

    def test_escapable_set_matches_the_emitter(self) -> None:
        # The lockstep invariant, asserted directly rather than inferred.
        emitted = set(_ALWAYS_ESCAPE) | {"#", "_"}
        assert set(_ESCAPABLE) == emitted

    def test_escaped_backslash_leaves_following_delimiter_live(self) -> None:
        # `\\*x*` is a literal backslash followed by an italic `x`.
        runs = parse_markup("\\\\*x*")
        assert [(r.text, r.italic) for r in runs] == [("\\", None), ("x", True)]


class TestMonospaceDetection:
    @pytest.mark.parametrize(
        "name",
        ["Consolas", "consolas", "Courier New", "Cascadia Mono", "DejaVu Sans Mono", "Roboto Mono"],
    )
    def test_monospace_families_detected(self, name: str) -> None:
        assert _font_monospace(name) is True

    @pytest.mark.parametrize("name", ["Calibri", "Times New Roman", "Monotype Corsiva", "", None])
    def test_proportional_families_rejected(self, name: object) -> None:
        # "Monotype Corsiva" is a script face — substring-matching "mono" would
        # wrongly backtick it. Empty string is Word's "varies within the word".
        assert _font_monospace(name) is False


class TestCodeSpanRendering:
    def test_code_span_wrapped_in_backticks(self) -> None:
        md = render_markdown([Block(PARAGRAPH, spans=_t(Span("attach()", code=True)))])
        assert md == "`attach()`"

    def test_code_content_is_not_backslash_escaped(self) -> None:
        # Escapes carry no meaning inside a code span; escaping would corrupt it.
        md = render_markdown([Block(PARAGRAPH, spans=_t(Span("a_b[0]*", code=True)))])
        assert md == "`a_b[0]*`"

    def test_fence_grows_to_contain_backticks(self) -> None:
        assert _code_span("a`b") == "``a`b``"
        assert _code_span("a``b") == "```a``b```"

    def test_backtick_edges_get_a_space_pad(self) -> None:
        assert _code_span("`") == "`` ` ``"

    def test_code_combines_with_emphasis(self) -> None:
        md = render_markdown([Block(PARAGRAPH, spans=_t(Span("x", code=True, bold=True)))])
        assert md == "**`x`**"

    @pytest.mark.parametrize("text", ["attach()", "a`b", "`", "x[0]", "a_b", "C:\\Users"])
    def test_code_span_round_trips(self, text: str) -> None:
        md = _render_spans(_t(Span(text, code=True)))
        runs = parse_markup(md)
        assert runs_to_text(runs) == text
        assert all(r.code for r in runs)


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

    def test_link_target_with_space_and_paren_angle_wrapped(self) -> None:
        # EXPORT-7: a URL with a space / ')' breaks the bare (url) form — wrap it.
        md = render_markdown(
            [
                Block(
                    PARAGRAPH,
                    spans=_t(Span("doc", href="https://s.test/My Files/a(1).docx")),
                )
            ]
        )
        assert md == "[doc](<https://s.test/My Files/a(1).docx>)"

    def test_clean_target_left_bare(self) -> None:
        md = render_markdown([Block(PARAGRAPH, spans=_t(Span("x", href="https://x.test/a")))])
        assert md == "[x](https://x.test/a)"

    def test_image_target_with_space_wrapped(self) -> None:
        blocks = [Block(IMAGE, image_ref="my image.png", image_alt="logo")]
        assert render_markdown(blocks) == "![logo](<my image.png>)"


class TestMdTarget:
    def test_bare_when_safe(self) -> None:
        assert _md_target("https://x.test/a#b") == "https://x.test/a#b"

    def test_wraps_space_and_parens(self) -> None:
        assert _md_target("a (b).png") == "<a (b).png>"

    def test_escapes_angle_brackets(self) -> None:
        assert _md_target("a<b>c d") == "<a%3Cb%3Ec d>"

    def test_empty_passthrough(self) -> None:
        assert _md_target("") == ""


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


class TestLeadSnippetAccounting:
    def test_fully_shown_snippet_emits_no_more_words_marker(self) -> None:
        # EXPORT-6: a single short sentence shown in full must NOT emit a spurious
        # "N more words" marker (share=0 forces the lead-snippet path).
        b = Block(
            PARAGRAPH,
            anchor_id="para:2",
            spans=_t(Span("alpha beta gamma delta")),
            word_count=4,
        )
        out = _render_body_segment([b], 0.0, suppress=False, budget_left=[100.0])
        assert "alpha beta gamma delta" in out
        assert "more words" not in out

    def test_truncated_snippet_still_marks_remainder(self) -> None:
        # A long multi-sentence paragraph still gets a marker, counted off the
        # same plain-text source as word_count.
        text = "one two three four five. " + "tail " * 30
        b = Block(
            PARAGRAPH,
            anchor_id="para:2",
            spans=_t(Span(text)),
            word_count=len(text.split()),
        )
        out = _render_body_segment([b], 0.0, suppress=False, budget_left=[100.0])
        assert "…(para:2, " in out and "more words)…" in out


class TestHeadingLevel:
    def test_outline_level_preferred(self) -> None:
        assert _heading_level(2, "Body Text") == 2

    def test_english_style_fallback(self) -> None:
        assert _heading_level(10, "Heading 3") == 3

    def test_localized_style_map_used(self) -> None:
        # EXPORT-8: a non-English heading style name resolves via the per-document
        # localized map, where the English regex would miss it.
        styles = {"Überschrift 1": 1, "Überschrift 2": 2}
        assert _heading_level(10, "Überschrift 2", styles) == 2

    def test_body_text_is_none(self) -> None:
        assert _heading_level(10, "Normal", {"Heading 1": 1}) is None
