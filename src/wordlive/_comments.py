"""Comments — Word's review-side annotation channel.

A comment points at a range in the document body (its *scope*) without changing
that text. `doc.comments.add(anchor, text)` is exactly the polite, side-channel
edit an agent should prefer over rewriting text directly — "flag this for
review" instead of "change it".

Comments are addressed by 1-based index (`doc.comments[2]`), matching Word's own
`Comments(n)` ordering. `Comment.resolve()` marks one done; the `Done` flag is
Word 2013+, so on older builds `done` reads `False` and `resolve()` raises.
"""

from __future__ import annotations

from collections.abc import Iterator
from typing import TYPE_CHECKING, Any

from . import _com
from .exceptions import AnchorNotFoundError

if TYPE_CHECKING:
    from ._anchors import Anchor
    from ._document import Document


def _clean(raw: Any) -> str:
    """Strip Word's trailing paragraph / cell markers from comment-range text."""
    return str(raw or "").rstrip("\r\n\x07")


class Comment:
    """A single review comment, located by its 1-based document index."""

    def __init__(self, doc: Document, com: Any, index: int) -> None:
        self._doc = doc
        self._com = com
        self._index = index

    @property
    def com(self) -> Any:
        """Raw COM Comment object — escape hatch (replies, ranges, etc.)."""
        return self._com

    @property
    def index(self) -> int:
        return self._index

    @property
    def author(self) -> str:
        with _com.translate_com_errors():
            return str(self._com.Author or "")

    @property
    def text(self) -> str:
        """The comment body."""
        with _com.translate_com_errors():
            return _clean(self._com.Range.Text)

    @property
    def scope_text(self) -> str:
        """The document text the comment is attached to (its anchored range)."""
        with _com.translate_com_errors():
            return _clean(self._com.Scope.Text)

    @property
    def done(self) -> bool:
        """Whether the comment is marked resolved/done. `False` on Word <2013."""
        try:
            return bool(self._com.Done)
        except Exception:
            return False

    def resolve(self) -> None:
        """Mark the comment as done/resolved (Word 2013+)."""
        with _com.translate_com_errors():
            self._com.Done = True

    def reopen(self) -> None:
        """Clear the done/resolved flag (Word 2013+)."""
        with _com.translate_com_errors():
            self._com.Done = False

    def delete(self) -> None:
        """Remove the comment from the document."""
        with _com.translate_com_errors():
            self._com.Delete()

    def to_dict(self) -> dict[str, Any]:
        """`{index, author, text, scope, done}` — the JSON shape `list()` emits."""
        with _com.translate_com_errors():
            return {
                "index": self._index,
                "author": str(self._com.Author or ""),
                "text": _clean(self._com.Range.Text),
                "scope": _clean(self._com.Scope.Text),
                "done": self.done,
            }

    def __repr__(self) -> str:
        return f"<Comment {self._index} by {self.author!r}>"


class CommentCollection:
    """Indexable, iterable view over a document's review comments."""

    def __init__(self, doc: Document) -> None:
        self._doc = doc

    def __len__(self) -> int:
        with _com.translate_com_errors():
            return int(self._doc.com.Comments.Count)

    def __getitem__(self, index: int) -> Comment:
        if isinstance(index, bool) or not isinstance(index, int):
            raise TypeError(f"comment index must be int, got {type(index).__name__}")
        n = len(self)
        if not (1 <= index <= n):
            raise AnchorNotFoundError("comment", str(index))
        with _com.translate_com_errors():
            return Comment(self._doc, self._doc.com.Comments(index), index)

    def __iter__(self) -> Iterator[Comment]:
        with _com.translate_com_errors():
            count = int(self._doc.com.Comments.Count)
        for i in range(1, count + 1):
            with _com.translate_com_errors():
                com = self._doc.com.Comments(i)
            yield Comment(self._doc, com, i)

    def add(self, anchor: Anchor, text: str, *, author: str | None = None) -> Comment:
        """Attach a new comment to `anchor`'s range.

        `anchor` is any wordlive anchor (bookmark, heading, cell, range, …); its
        COM range becomes the comment's scope and the document text is left
        untouched — only an annotation is added. Returns the new `Comment`.
        """
        with _com.translate_com_errors():
            rng = anchor.com
            comments = self._doc.com.Comments
            com = comments.Add(rng, text)
            if author:
                try:
                    com.Author = author
                except Exception:
                    # Some COM builds reject a per-comment Author write; the
                    # comment still lands with the app's default author.
                    pass
            index = int(comments.Count)
        return Comment(self._doc, com, index)

    def list(self) -> list[dict[str, Any]]:
        """All comments as `{index, author, text, scope, done}` dicts."""
        return [c.to_dict() for c in self]
