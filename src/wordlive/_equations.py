"""Equation helpers — turn a math source into Office Math (OMML) and back.

Word's native math object model (`Range.OMaths`) speaks two things well:

- **UnicodeMath** — a linear string (``x=(-b±√(b^2-4ac))/(2a)``) you type into a
  math zone and *build up* into the 2-D professional form. This is native and
  dependency-free: `insert_equation(unicodemath=…)` writes the string, wraps it
  in an `OMaths.Add`, and calls `BuildUp()`. No XML involved.

- **OMML** (Office MathML, the ``<m:oMath>`` dialect) — Word's on-disk math
  markup. Anything that isn't UnicodeMath reaches Word as OMML through
  `Range.InsertXML`.

The bridge for **MathML** and **LaTeX** is Office's *own* XSLT, shipped beside
every Word install (``MML2OMML.XSL`` / ``OMML2MML.XSL``):

    LaTeX --[latex2mathml]--> MathML --[MML2OMML.XSL]--> OMML --[InsertXML]--> Word

Only the first hop needs a third-party library (`latex2mathml`, the optional
``latex`` extra); MathML onward is pure Office. The reverse (`omml_to_mathml`,
via ``OMML2MML.XSL``) is the read side — it serialises an existing equation back
to MathML without mutating the document.

The XSLT transforms run through MSXML (`MSXML2.DOMDocument`), which is present on
every Windows box, so they need no extra dependency. All COM/Win32 imports are
lazy, keeping the module importable off-Windows for the fake-COM test suite.
"""

from __future__ import annotations

import os
from typing import Any

from .exceptions import EquationError

# The Office math transforms live next to winword.exe. Filenames are stable
# across Office versions; only the install directory varies, so we probe the
# common roots (and let the caller fall back to COM's own resolution if needed).
_MML2OMML = "MML2OMML.XSL"  # MathML -> OMML (insert side)
_OMML2MML = "OMML2MML.XSL"  # OMML -> MathML (read side)

_xsl_cache: dict[str, str] = {}


def _office_roots() -> list[str]:
    """Candidate Office install directories that may hold the math transforms."""
    roots: list[str] = []
    for env in ("ProgramFiles", "ProgramFiles(x86)", "ProgramW6432"):
        base = os.environ.get(env)
        if not base:
            continue
        office = os.path.join(base, "Microsoft Office")
        # Click-to-Run nests the binaries under root\OfficeNN; MSI installs them
        # directly under OfficeNN. Probe both layouts for any Office generation.
        for mid in ("root", ""):
            for ver in ("Office16", "Office15", "Office14"):
                roots.append(os.path.join(office, mid, ver) if mid else os.path.join(office, ver))
    return roots


def _load_xsl(filename: str) -> str:
    """Return the text of an Office math XSLT, raising `EquationError` if absent.

    The result is cached: the stylesheet is a few hundred KB and never changes
    within a run, so we read it once. Read with ``utf-8-sig`` to drop the BOM the
    Office stylesheets carry (MSXML's `loadXML` rejects a leading BOM).
    """
    if filename in _xsl_cache:
        return _xsl_cache[filename]
    for root in _office_roots():
        path = os.path.join(root, filename)
        if os.path.isfile(path):
            with open(path, encoding="utf-8-sig") as f:
                text = f.read()
            _xsl_cache[filename] = text
            return text
    raise EquationError(
        f"could not find Office's {filename} math transform; a standard Microsoft "
        "Word install ships it beside winword.exe. MathML/LaTeX equations need it "
        "(UnicodeMath input does not)."
    )


def _transform(source_xml: str, xsl_text: str, *, what: str) -> str:
    """Run an XSLT over `source_xml` via MSXML, returning the result string.

    MSXML (`MSXML2.DOMDocument.6.0`) is the always-present Windows XML engine, so
    this adds no dependency. `loadXML` is synchronous for in-memory strings, so
    the reserved-word `async` property never needs setting. A parse/transform
    failure becomes a clean `EquationError`.
    """
    import win32com.client as wc  # type: ignore[import-not-found]  # lazy: Windows-only

    src = wc.Dispatch("MSXML2.DOMDocument.6.0")
    if not src.loadXML(source_xml):
        raise EquationError(
            f"{what}: input is not well-formed XML ({src.parseError.reason.strip()})"
        )
    xsl = wc.Dispatch("MSXML2.DOMDocument.6.0")
    if not xsl.loadXML(xsl_text):
        raise EquationError(f"{what}: the Office transform failed to load")
    try:
        return str(src.transformNode(xsl))
    except Exception as e:  # noqa: BLE001 — surfaced as clean bad-input
        raise EquationError(f"{what}: XSLT transform failed ({e})") from e


def _strip_xml_decl(xml: str) -> str:
    """Drop a leading ``<?xml …?>`` declaration so a fragment can be spliced."""
    if xml.startswith("<?xml"):
        _, _, rest = xml.partition("?>")
        return rest
    return xml


def latex_to_mathml(latex: str) -> str:
    """Convert a LaTeX math string to MathML via the optional `latex2mathml` extra.

    `latex2mathml` is pure-Python and pulled in by the ``latex`` extra; when it
    isn't installed this raises `EquationError` with the install hint, so the
    core (UnicodeMath + MathML) stays dependency-free. A LaTeX string the
    converter can't parse also surfaces as `EquationError`.
    """
    try:
        from latex2mathml.converter import convert  # type: ignore[import-not-found]
    except ImportError as e:
        raise EquationError(
            "LaTeX equations need the latex2mathml backend; install the extra with "
            '`pip install "wordlive[latex]"` (or `uv add "wordlive[latex]"`). '
            "UnicodeMath and MathML inputs work without it."
        ) from e
    try:
        return str(convert(latex))
    except EquationError:
        raise
    except Exception as e:  # noqa: BLE001 — any converter failure is bad input
        raise EquationError(f"could not convert LaTeX to MathML: {e}") from e


def mathml_to_omml(mathml: str) -> str:
    """Convert a MathML string to an OMML ``<m:oMath>`` fragment via Office's XSLT.

    Returns the bare OMML element (XML declaration stripped) ready to splice into
    a WordprocessingML package for `Range.InsertXML`. Raises `EquationError` for
    malformed MathML or a missing transform.
    """
    omml = _transform(mathml, _load_xsl(_MML2OMML), what="MathML→OMML")
    return _strip_xml_decl(omml).strip()


def omml_to_mathml(package_or_omml_xml: str) -> str:
    """Convert OMML (or a package containing it) to MathML via Office's XSLT.

    The read side: hand it an equation range's ``WordOpenXML`` (or a bare
    ``<m:oMath>``) and get presentation MathML back, without touching the
    document. Raises `EquationError` if the transform is missing or fails.
    """
    mathml = _transform(package_or_omml_xml, _load_xsl(_OMML2MML), what="OMML→MathML")
    return _strip_xml_decl(mathml).strip()


def equation_package(template_xml: str, omml_inner: str, *, display: bool) -> str:
    """Splice an OMML fragment into a real `WordOpenXML` template for InsertXML.

    `Range.InsertXML` rejects a bare ``<m:oMath>`` fragment (and even a minimal
    hand-built package) — it validates the full package skeleton (content types,
    relationships). So we take a *live* `Range.WordOpenXML` as the template and
    inject one math paragraph just before the body's ``<w:sectPr>``. A `display`
    equation is wrapped in ``<m:oMathPara>`` (its own centred line); an inline
    one drops the bare ``<m:oMath>`` into the paragraph.
    """
    body = f"<m:oMathPara>{omml_inner}</m:oMathPara>" if display else omml_inner
    para = f"<w:p>{body}</w:p>"
    if "<w:sectPr" not in template_xml:
        raise EquationError("equation template is missing a section break to anchor against")
    return template_xml.replace("<w:sectPr", para + "<w:sectPr", 1)


def omath_in_range(doc_com: Any, start: int) -> Any | None:
    """Return the document's `OMath` whose range contains `start`, or `None`.

    Used after a native `OMaths.Add` to recover the freshly created zone (Word
    has no "return the OMath I just added" call), and to resolve `equation:N`
    consistently with document order.
    """
    omaths = doc_com.OMaths
    for i in range(1, int(omaths.Count) + 1):
        zone = omaths.Item(i)
        rng = zone.Range
        if int(rng.Start) <= start < int(rng.End):
            return zone
    return None
