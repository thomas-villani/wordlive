"""The `Anchor` spine: the COM range, its text, and the abstract hooks a
concrete anchor must implement. Every `Anchor` feature mixin builds on this."""

from __future__ import annotations

from abc import ABC, abstractmethod
from typing import TYPE_CHECKING, Any, cast

from .. import _com

if TYPE_CHECKING:
    from .._document import Document

from ._helpers import (
    range_text,
)

if TYPE_CHECKING:
    pass

if TYPE_CHECKING:
    from ._base import Anchor


class AnchorCore(ABC):
    """Core state and primitives shared by every `Anchor` feature mixin."""

    kind: str = "anchor"
    name: str = ""

    @property
    def _as_anchor(self) -> Anchor:
        """`self`, narrowed to the concrete `Anchor`.

        The feature mixins are only ever mixed into `Anchor`, but a type checker
        only sees the mixin, and collaborators are annotated on `Anchor`.
        Runtime-free: `cast` returns its argument unchanged.
        """
        return cast("Anchor", self)

    def __init__(self, doc: Document, name: str) -> None:
        self._doc = doc
        self.name = name

    @property
    def com(self) -> Any:
        """Raw COM range. Subclasses override."""
        return self._range()

    @abstractmethod
    def _range(self) -> Any:
        """Return the COM Range that this anchor refers to. Must be overridden."""

    @property
    def text(self) -> str:
        with _com.translate_com_errors():
            return range_text(self._range())

    @property
    @abstractmethod
    def anchor_id(self) -> str:
        """Stable string identifier for this anchor (e.g. `bookmark:Address`).

        Each anchor kind has its own scheme (`bookmark:`, `cc:`, `heading:`),
        so subclasses must declare theirs explicitly — no useful default
        exists at this level.
        """

    @abstractmethod
    def set_text(self, text: str) -> None:
        """Replace the anchor's text in place. Must be overridden."""

    def delete(self) -> None:
        with _com.translate_com_errors():
            self._range().Delete()

    def __repr__(self) -> str:
        return f"<{type(self).__name__} {self.name!r}>"
