"""`equation:N` anchors and the equation collection (see also `_equations`)."""

from __future__ import annotations

from collections.abc import Iterator
from typing import TYPE_CHECKING, Any

from .. import _com, _equations
from ..exceptions import AnchorNotFoundError, OpError

if TYPE_CHECKING:
    from .._document import Document

from ._base import Anchor


class EquationAnchor(Anchor):
    """A mathematical equation located by 1-based index — `equation:N`.

    Mirrors Word's own `OMaths(N)` ordering (document order). The anchor resolves
    to the equation's range, so `mathml` round-trips it back to MathML (via
    Office's own transform, without mutating the document) and `linear` reads its
    UnicodeMath form. `type` is ``"display"`` or ``"inline"``. Create equations
    with [`Anchor.insert_equation`][wordlive.Anchor.insert_equation]; discover
    them via [`doc.equations`][wordlive.Document.equations]. An equation isn't
    plain text, so `set_text` raises — delete and re-insert to change it.
    """

    kind = "equation"

    def __init__(self, doc: Document, index: int) -> None:
        super().__init__(doc, name=f"equation:{index}")
        self._index = index

    @property
    def index(self) -> int:
        return self._index

    @property
    def anchor_id(self) -> str:
        return f"equation:{self._index}"

    def _omath(self) -> Any:
        omaths = self._doc.com.OMaths
        n = int(omaths.Count)
        if not (1 <= self._index <= n):
            raise AnchorNotFoundError("equation", f"equation:{self._index}")
        return omaths.Item(self._index)

    def _range(self) -> Any:
        return self._omath().Range

    @property
    def type(self) -> str:
        """``"display"`` (its own centred line) or ``"inline"`` (in the text flow)."""
        with _com.translate_com_errors():
            # WdOMathType: wdOMathDisplay == 1, wdOMathInline == 0.
            return "display" if int(self._omath().Type) == 1 else "inline"

    @property
    def mathml(self) -> str:
        """The equation as MathML — a non-mutating read via Office's OMML→MathML transform."""
        with _com.translate_com_errors():
            package = str(self._omath().Range.WordOpenXML)
        return _equations.omml_to_mathml(package)

    @property
    def linear(self) -> str:
        """The equation's text in Word's built-up linear form (a compact preview).

        Reads the zone's text with the internal structure markers collapsed — a
        readable approximation of the math, not a precise round-trip. For
        fidelity use [`mathml`][wordlive.EquationAnchor.mathml].
        """
        with _com.translate_com_errors():
            raw = str(self._omath().Range.Text or "")
        return raw.replace("\r", "").replace("\x0b", "").strip()

    def set_text(self, text: str) -> None:
        raise OpError(
            "an equation anchor has no plain text to set; delete it and "
            "insert_equation(...) again to change it"
        )


class EquationCollection:
    """Read-only, iterable view over the document's equations (`doc.equations`).

    Index an equation by 1-based position (`doc.equations[2]`) to get an
    [`EquationAnchor`][wordlive.EquationAnchor] (`equation:N`), then `mathml` /
    `linear` to read it. `list()` summarises every equation — id, type, a linear
    preview, and the `para:N` it sits in. Positions match Word's own `OMaths(n)`
    ordering. The write mirror is any anchor's
    [`insert_equation`][wordlive.Anchor.insert_equation].
    """

    def __init__(self, doc: Document) -> None:
        self._doc = doc

    def __len__(self) -> int:
        with _com.translate_com_errors():
            return int(self._doc.com.OMaths.Count)

    def __getitem__(self, index: int) -> EquationAnchor:
        if isinstance(index, bool) or not isinstance(index, int):
            raise TypeError(f"equation index must be int, got {type(index).__name__}")
        n = len(self)
        if not (1 <= index <= n):
            raise AnchorNotFoundError("equation", str(index))
        return EquationAnchor(self._doc, index)

    def __iter__(self) -> Iterator[EquationAnchor]:
        with _com.translate_com_errors():
            count = int(self._doc.com.OMaths.Count)
        for i in range(1, count + 1):
            yield EquationAnchor(self._doc, i)

    def list(self) -> list[dict[str, Any]]:
        """Every equation as `{index, anchor_id, type, linear, para}`.

        `type` is ``"display"`` / ``"inline"``; `linear` is the built-up text as
        a compact preview (read [`EquationAnchor.mathml`][wordlive.EquationAnchor]
        for fidelity); `para` is the `para:N` the equation sits in (or ``None``).
        Reads no XML, so this is cheap to call over a whole document.
        """
        out: list[dict[str, Any]] = []
        with _com.translate_com_errors():
            omaths = self._doc.com.OMaths
            count = int(omaths.Count)
            for i in range(1, count + 1):
                zone = omaths.Item(i)
                rng = zone.Range
                try:
                    start = int(rng.Start)
                except Exception:
                    start = None
                try:
                    eq_type = "display" if int(zone.Type) == 1 else "inline"
                except Exception:
                    eq_type = "inline"
                linear = str(rng.Text or "").replace("\r", "").replace("\x0b", "").strip()
                para_id: str | None = None
                if start is not None:
                    para = self._doc.paragraphs.at(start)
                    para_id = para.anchor_id if para is not None else None
                out.append(
                    {
                        "index": i,
                        "anchor_id": f"equation:{i}",
                        "type": eq_type,
                        "linear": linear,
                        "para": para_id,
                    }
                )
        return out
