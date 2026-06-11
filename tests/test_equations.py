"""Unit tests for the equation surface (off-Windows, against the fake COM).

The *execution* of an insert (native UnicodeMath build / OMML InsertXML) needs
live paragraph geometry and MSXML, so it lives in the smoke suite. Here we cover
everything that doesn't: input validation, the read collection (`doc.equations`),
anchor resolution (`equation:N`), the ops/MCP request builders, and the
LaTeX→MathML hop (pure-Python, runs anywhere).
"""

from __future__ import annotations

import pytest

import wordlive
from wordlive import _equations
from wordlive.exceptions import AnchorNotFoundError, EquationError, OpError

# ---------------------------------------------------------------------------
# insert_equation — input validation (raises before any COM call)
# ---------------------------------------------------------------------------


def test_insert_equation_requires_an_input_dialect(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        with pytest.raises(EquationError, match="exactly one of"):
            doc.start.insert_equation()


def test_insert_equation_rejects_multiple_dialects(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        with pytest.raises(EquationError, match="exactly one of"):
            doc.start.insert_equation(unicodemath="a", latex="b")


def test_insert_equation_rejects_bad_where(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        with pytest.raises(ValueError, match="where must be"):
            doc.start.insert_equation(unicodemath="a^2", where="sideways")


# ---------------------------------------------------------------------------
# doc.equations — the read collection (fake seeds one display equation)
# ---------------------------------------------------------------------------


def test_equations_collection_length_and_iteration(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        assert len(doc.equations) == 1
        ids = [eq.anchor_id for eq in doc.equations]
    assert ids == ["equation:1"]


def test_equations_list_shape(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        rows = doc.equations.list()
    assert rows == [
        {
            "index": 1,
            "anchor_id": "equation:1",
            "type": "display",
            "linear": "𝐸=𝑚𝑐2",  # built-up text with the CR markers stripped
            "para": "para:2",  # the seeded range (21–27) sits in the body paragraph
        }
    ]


def test_equation_anchor_type_and_linear(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        eq = doc.equations[1]
        assert eq.type == "display"
        assert eq.linear == "𝐸=𝑚𝑐2"
        assert eq.anchor_id == "equation:1"


def test_equation_anchor_set_text_raises(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        with pytest.raises(OpError, match="no plain text"):
            doc.equations[1].set_text("nope")


def test_equations_index_out_of_range(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        with pytest.raises(AnchorNotFoundError):
            _ = doc.equations[2]


# ---------------------------------------------------------------------------
# anchor_by_id — equation:N resolution
# ---------------------------------------------------------------------------


def test_anchor_by_id_resolves_equation(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        eq = doc.anchor_by_id("equation:1")
        assert isinstance(eq, wordlive.EquationAnchor)
        assert eq.index == 1


def test_anchor_by_id_equation_bad_index(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        with pytest.raises(AnchorNotFoundError):
            doc.anchor_by_id("equation:banana")


# ---------------------------------------------------------------------------
# exec-op + MCP request builders (no COM)
# ---------------------------------------------------------------------------


def test_insert_equation_in_op_registries():
    from wordlive._ops import OP_OPTIONAL_FIELDS, OP_REQUIRED_FIELDS

    assert OP_REQUIRED_FIELDS["insert_equation"] == ("anchor_id",)
    optional = OP_OPTIONAL_FIELDS["insert_equation"]
    for field in ("unicodemath", "latex", "mathml", "display", "before"):
        assert field in optional


def test_mcp_build_write_op_insert_equation():
    from wordlive.mcp.server import _build_write_op

    op = _build_write_op(
        "insert_equation",
        {"anchor_id": "heading:1", "latex": r"\frac12", "display": False, "before": True},
    )
    assert op == {
        "op": "insert_equation",
        "anchor_id": "heading:1",
        "before": True,
        "latex": r"\frac12",
        "display": False,
    }


def test_mcp_build_write_op_insert_equation_requires_one_dialect():
    from wordlive.mcp.server import _build_write_op

    with pytest.raises(OpError, match="exactly one"):
        _build_write_op("insert_equation", {"anchor_id": "heading:1"})
    with pytest.raises(OpError, match="exactly one"):
        _build_write_op("insert_equation", {"anchor_id": "heading:1", "latex": "a", "mathml": "b"})


# ---------------------------------------------------------------------------
# LaTeX → MathML hop (pure-Python; behaviour depends on the optional backend)
# ---------------------------------------------------------------------------


def test_latex_to_mathml_with_or_without_backend():
    try:
        import latex2mathml  # noqa: F401
    except ImportError:
        with pytest.raises(EquationError, match="latex2mathml"):
            _equations.latex_to_mathml(r"\frac{1}{2}")
    else:
        out = _equations.latex_to_mathml(r"\frac{1}{2}")
        assert "<math" in out and "mfrac" in out


def test_equation_package_splices_before_sectpr():
    template = "<w:document><w:body><w:p/><w:sectPr/></w:body></w:document>"
    pkg = _equations.equation_package(template, "<m:oMath/>", display=True)
    assert "<m:oMathPara><m:oMath/></m:oMathPara>" in pkg
    # The injected paragraph precedes the section break it anchors against.
    assert pkg.index("oMathPara") < pkg.index("<w:sectPr")


def test_equation_package_inline_omits_omathpara():
    template = "<w:body><w:sectPr/></w:body>"
    pkg = _equations.equation_package(template, "<m:oMath/>", display=False)
    assert "oMathPara" not in pkg
    assert "<w:p><m:oMath/></w:p>" in pkg
