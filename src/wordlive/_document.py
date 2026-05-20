"""Document wrapper + DocumentCollection."""

from __future__ import annotations

from contextlib import contextmanager
from typing import Any, Iterator, TYPE_CHECKING

from . import _com, _findreplace
from ._anchors import (
    BookmarkCollection,
    ContentControlCollection,
    Heading,
    HeadingCollection,
    RangeAnchor,
    _IndexedHeading,
    paragraph_text,
)
from ._comments import CommentCollection
from ._edit import EditScope
from ._selection import Selection
from ._styles import Style, StyleCollection
from ._tables import TableCollection
from .exceptions import (
    AmbiguousMatchError,
    AnchorNotFoundError,
    DocumentNotFoundError,
)

if TYPE_CHECKING:
    from ._app import Word
    from ._anchors import Anchor


class Document:
    """Wraps a Word Document COM object."""

    def __init__(self, word: "Word", doc: Any) -> None:
        self._word = word
        self._doc = doc

    @property
    def com(self) -> Any:
        return self._doc

    @property
    def name(self) -> str:
        with _com.translate_com_errors():
            return str(self._doc.Name)

    @property
    def path(self) -> str:
        with _com.translate_com_errors():
            return str(self._doc.FullName)

    @property
    def bookmarks(self) -> BookmarkCollection:
        return BookmarkCollection(self)

    @property
    def content_controls(self) -> ContentControlCollection:
        return ContentControlCollection(self)

    @property
    def styles(self) -> StyleCollection:
        return StyleCollection(self)

    @property
    def tables(self) -> TableCollection:
        """Iterable, indexable view over the document's tables.

        Index by 1-based position (`doc.tables[1]`) or `Title`
        (`doc.tables["Budget"]`). Cells are anchors: `doc.tables[1].cell(2, 3)`
        — or `doc.anchor_by_id("table:1:2:3")` — returns a `Cell` that works
        with `set_text`, `apply_style`, and `format_paragraph`.
        """
        return TableCollection(self)

    @property
    def headings(self) -> HeadingCollection:
        """Iterable view over the document's headings.

        Symmetric with `bookmarks`, `content_controls`, and `styles`. Index by
        visible text (`doc.headings["Risks"]`) or 1-based paragraph position
        (`doc.headings[3]`). `Document.heading(name)` remains as sugar for
        `self.headings[name]`.
        """
        return HeadingCollection(self)

    @property
    def comments(self) -> CommentCollection:
        """Iterable, indexable view over the document's review comments.

        `doc.comments.add(anchor, text, author=...)` attaches a comment to any
        anchor's range without changing the text — the polite, side-channel way
        to flag something. Index existing comments by 1-based position
        (`doc.comments[2]`) to `resolve()` or `delete()` them.
        """
        return CommentCollection(self)

    @property
    def selection(self) -> Selection:
        return self._word.selection

    @property
    def track_changes(self) -> bool:
        """Whether Word's Track Changes is currently on for this document."""
        with _com.translate_com_errors():
            return bool(self._doc.TrackRevisions)

    @track_changes.setter
    def track_changes(self, value: bool) -> None:
        with _com.translate_com_errors():
            self._doc.TrackRevisions = bool(value)

    @contextmanager
    def tracked_changes(self) -> Iterator[None]:
        """Turn on Track Changes for the duration of the block, then restore it.

        Every mutation made inside the scope is recorded as a tracked revision
        the user can accept or reject — "make this edit *visibly*." The prior
        `TrackRevisions` setting is restored on exit, so the scope stays polite
        even when the user had tracking off.

        Pairs with `edit()` for an atomic, visibly-tracked batch:

            with doc.tracked_changes(), doc.edit("Suggest rewordings"):
                doc.find_replace("utilise", "use", all=True)
        """
        with _com.translate_com_errors():
            previous = bool(self._doc.TrackRevisions)
            self._doc.TrackRevisions = True
        try:
            yield
        finally:
            with _com.translate_com_errors():
                self._doc.TrackRevisions = previous

    def heading(self, name: str) -> Heading:
        # Lazy lookup — Heading.__init__ doesn't hit COM. _range() validates.
        return Heading(self, name)

    def range(self, start: int, end: int) -> RangeAnchor:
        """Return a `RangeAnchor` over the absolute offsets `[start, end)`.

        Offsets are UTF-16 code units — the coordinates Word uses and that
        `find()` emits as `range:START-END`. Lazy: the offsets aren't validated
        against the document until the anchor is used.
        """
        return RangeAnchor(self, start, end)

    def anchor_by_id(self, anchor_id: str) -> "Anchor":
        """Resolve an `anchor_id` string into an Anchor.

        Recognised forms:
          - `heading:N`       — Nth paragraph in the document (1-based, must be a heading)
          - `bookmark:NAME`   — bookmark by name
          - `cc:NAME`         — content control by Title (or Tag)
          - `table:N:R:C`     — cell at 1-based (row, column) of the Nth table
          - `range:START-END` — arbitrary character span (the form `find()` emits)

        The bare `table:N` form is not an anchor (a whole table is a collection,
        not a single range) — use `doc.tables[N]` instead.

        Raises `AnchorNotFoundError` for unknown schemes or missing anchors.
        """
        if not isinstance(anchor_id, str) or ":" not in anchor_id:
            raise AnchorNotFoundError("anchor", anchor_id)
        kind, _, value = anchor_id.partition(":")
        if kind == "heading":
            try:
                idx = int(value)
            except ValueError as e:
                raise AnchorNotFoundError("heading", anchor_id) from e
            return _IndexedHeading(self, idx)
        if kind == "bookmark":
            return self.bookmarks[value]
        if kind == "cc":
            return self.content_controls[value]
        if kind == "table":
            parts = value.split(":")
            if len(parts) != 3:
                # `table:N` (whole table) isn't a single-range anchor.
                raise AnchorNotFoundError("table cell", anchor_id)
            try:
                t, r, c = (int(p) for p in parts)
            except ValueError as e:
                raise AnchorNotFoundError("table cell", anchor_id) from e
            return self.tables[t].cell(r, c)
        if kind == "range":
            start_str, sep, end_str = value.partition("-")
            if not sep:
                raise AnchorNotFoundError("range", anchor_id)
            try:
                start, end = int(start_str), int(end_str)
            except ValueError as e:
                raise AnchorNotFoundError("range", anchor_id) from e
            try:
                return self.range(start, end)
            except ValueError as e:
                raise AnchorNotFoundError("range", anchor_id) from e
        raise AnchorNotFoundError("anchor", anchor_id)

    def _scope_range(self, scope: "Anchor | None") -> tuple[Any, int]:
        """Return (COM Range, absolute_start_offset) for a find/replace scope.

        Headings expand to their *section* (body under the heading); other
        anchor kinds use their own range. `None` means the whole document.
        """
        with _com.translate_com_errors():
            if scope is None:
                rng = self._doc.Content
            elif isinstance(scope, Heading):
                rng = scope.section_range()
            else:
                rng = scope.com
            return rng, int(rng.Start)

    def find(
        self,
        text: str,
        *,
        scope: "Anchor | None" = None,
    ) -> list[dict[str, Any]]:
        """Locate every fuzzy occurrence of `text` within `scope` (or the whole doc).

        Matching is whitespace- and Unicode-normalized (NFKC, smart quotes,
        dashes, NBSP). Returns a list of `{anchor_id, start, end, text}` where
        offsets are absolute document positions and `text` is the actual
        original substring (not the normalized form).

        `anchor_id` for each match is `range:START-END`, which resolves through
        `anchor_by_id` to a `RangeAnchor` — so a hit can be fed straight back
        into `replace --anchor-id` or `comments.add`. The offsets are live,
        though, so use them before further edits shift the document.
        """
        with _com.translate_com_errors():
            rng, base = self._scope_range(scope)
            haystack = str(rng.Text or "")

        matches = _findreplace.find_matches(haystack, text)
        return [
            {
                "anchor_id": f"range:{base + m.start}-{base + m.end}",
                "start": base + m.start,
                "end": base + m.end,
                "text": m.text,
            }
            for m in matches
        ]

    def find_replace(
        self,
        find: str,
        replace: str,
        *,
        scope: "Anchor | None" = None,
        all: bool = False,
        occurrence: int | None = None,
    ) -> list[dict[str, Any]]:
        """Fuzzy plain-text replace. See `find()` for matching semantics.

        Args:
            find: the text to look for (fuzzy-matched).
            replace: the replacement text.
            scope: optional anchor to restrict the search to. Headings expand
                to their body section.
            all: replace every match.
            occurrence: 1-based index — replace only the Nth match.

        Raises:
            AnchorNotFoundError: zero matches (uses `kind='find'`).
            AmbiguousMatchError: more than one match and neither `all` nor
                `occurrence` was given.

        Returns the list of replacements actually applied, each
        `{anchor_id, start, end, text}` in their pre-replacement coordinates.
        """
        with _com.translate_com_errors():
            rng, base = self._scope_range(scope)
            haystack = str(rng.Text or "")

        matches = _findreplace.find_matches(haystack, find)
        if not matches:
            raise AnchorNotFoundError("find", find)

        match_payloads = [
            {
                "anchor_id": f"range:{base + m.start}-{base + m.end}",
                "start": base + m.start,
                "end": base + m.end,
                "text": m.text,
            }
            for m in matches
        ]

        if occurrence is not None:
            if occurrence < 1 or occurrence > len(matches):
                raise AnchorNotFoundError("find", f"{find} (occurrence {occurrence})")
            to_apply = [match_payloads[occurrence - 1]]
        elif all:
            to_apply = match_payloads
        elif len(matches) == 1:
            to_apply = match_payloads
        else:
            raise AmbiguousMatchError(find, match_payloads)

        with _com.translate_com_errors():
            # Apply in reverse so earlier offsets don't shift.
            for m in reversed(to_apply):
                target = self._doc.Range(m["start"], m["end"])
                target.Text = replace
        return to_apply

    def outline(self) -> list[dict[str, Any]]:
        """Return all heading paragraphs as `[{level, text, anchor_id}, ...]`."""
        out: list[dict[str, Any]] = []
        with _com.translate_com_errors():
            for idx, para in enumerate(self._doc.Paragraphs, start=1):
                try:
                    level = int(para.OutlineLevel)
                except Exception:
                    continue
                if level >= 10:
                    continue
                out.append(
                    {
                        "level": level,
                        "text": paragraph_text(para),
                        "anchor_id": f"heading:{idx}",
                    }
                )
        return out

    @contextmanager
    def edit(self, label: str) -> Iterator[EditScope]:
        """Open an atomic-undo / Selection-preserving edit scope.

        ```
        with doc.edit("Update address"):
            doc.bookmarks["Address"].set_text("…")
        ```
        """
        scope = EditScope(self._word, label)
        with scope:
            yield scope

    def go_to(self, anchor: "Anchor", scroll: bool = True) -> None:
        """Move the user's Selection to the given anchor (rare — most ops preserve it).

        Does NOT open an `UndoRecord` — cursor moves don't belong on the user's
        undo stack. If you want the move to ride along with a batch of edits,
        call this inside a `doc.edit(...)` scope and the surrounding
        `UndoRecord` will still group everything together.
        """
        with _com.translate_com_errors():
            rng = anchor.com
            collapsed = self._doc.Range(int(rng.Start), int(rng.Start))
            collapsed.Select()
            if scroll:
                try:
                    self._word.com.ActiveWindow.ScrollIntoView(collapsed)
                except Exception:
                    pass


class DocumentCollection:
    """Indexable view over open documents."""

    def __init__(self, word: "Word") -> None:
        self._word = word

    @property
    def _com_collection(self) -> Any:
        return self._word.com.Documents

    @property
    def active(self) -> Document:
        with _com.translate_com_errors():
            try:
                doc = self._word.com.ActiveDocument
            except Exception as e:
                raise DocumentNotFoundError("<active>") from e
        return Document(self._word, doc)

    def __getitem__(self, name: str) -> Document:
        with _com.translate_com_errors():
            for doc in self._com_collection:
                if str(doc.Name) == name:
                    return Document(self._word, doc)
        raise DocumentNotFoundError(name)

    def __iter__(self) -> Iterator[Document]:
        with _com.translate_com_errors():
            docs = list(self._com_collection)
        for d in docs:
            yield Document(self._word, d)

    def __len__(self) -> int:
        with _com.translate_com_errors():
            return int(self._com_collection.Count)

    def list(self) -> list[dict[str, Any]]:
        """`[{name, path, is_active}, ...]` — used by `wordlive status`."""
        out: list[dict[str, Any]] = []
        with _com.translate_com_errors():
            active_name: str | None
            try:
                active_name = str(self._word.com.ActiveDocument.Name)
            except Exception:
                active_name = None
            for doc in self._com_collection:
                name = str(doc.Name)
                out.append(
                    {
                        "name": name,
                        "path": str(doc.FullName),
                        "is_active": name == active_name,
                    }
                )
        return out
