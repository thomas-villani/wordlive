"""`word_snapshot` rendering implementation."""

from __future__ import annotations

from .. import attach
from .._ops import (
    pick_doc,
)
from ..exceptions import OpError
from ._worker import Worker


def _parse_pages(pages: str | None) -> int | tuple[int, int] | None:
    """`"4"` -> 4, `"2-5"` -> (2, 5), None/"" -> None (whole document)."""
    if not pages:
        return None
    s = str(pages).strip()
    if "-" in s:
        a, _, b = s.partition("-")
        try:
            return int(a), int(b)
        except ValueError as e:
            raise OpError(f"invalid pages range: {pages!r}") from e
    try:
        return int(s)
    except ValueError as e:
        raise OpError(f"invalid pages value: {pages!r}") from e


def _snapshot_impl(
    worker: Worker,
    *,
    doc: str | None,
    pages: str | None,
    anchor: str | None,
    dpi: int,
    markup: str,
    max_dim: int | None = None,
) -> list[tuple[int, bytes]]:
    dpi = max(72, min(300, int(dpi)))
    md = max(1, int(max_dim)) if max_dim is not None else None
    pages_arg = _parse_pages(pages)

    def job() -> list[tuple[int, bytes]]:
        with attach() as word:
            d = pick_doc(word, doc)
            if anchor:
                snaps = d.snapshot_anchor(
                    d.anchor_by_id(anchor), dpi=dpi, max_dim=md, markup=markup
                )
            else:
                snaps = d.snapshot(pages=pages_arg, dpi=dpi, max_dim=md, markup=markup)
            return [(s.page, s.png) for s in snaps]

    return worker.run_on_word(job)
