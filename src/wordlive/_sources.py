"""Bibliography sources — the document's source store behind `doc.sources`.

Word keeps a per-document repository of bibliography *sources* (the books,
articles, cases you cite). Citations reference a source by its **tag**; the
`BIBLIOGRAPHY` field renders the cited ones as a reference list. wordlive surfaces
the store as `doc.sources` — a collection mirroring `doc.content_controls`.

Adding a source means handing Word a ``<b:Source>`` XML element. `sources.add(...)`
builds that XML from friendly fields (a book/journal-article/… with author, title,
year, …); `sources.add_xml(...)` is the raw escape hatch for callers who already
have the OOXML. Both go through ``doc.com.Bibliography.Sources.Add`` (confirmed
against live Word: it ingests a single ``<b:Source>`` and round-trips the XML
byte-identically).
"""

from __future__ import annotations

import re
import xml.etree.ElementTree as ET
from collections.abc import Iterator
from typing import TYPE_CHECKING, Any
from xml.sax.saxutils import escape

from . import _com
from .exceptions import AnchorNotFoundError, OpError

if TYPE_CHECKING:
    from ._document import Document

# The OOXML bibliography namespace every <b:Source> lives in.
_NS = "http://schemas.openxmlformats.org/officeDocument/2006/bibliography"

# Friendly source-type keyword -> Word's `b:SourceType` literal. Canonical keys
# plus a few forgiving aliases (the names an agent is likely to reach for).
_SOURCE_TYPES: dict[str, str] = {
    "book": "Book",
    "book_section": "BookSection",
    "journal_article": "JournalArticle",
    "journal": "JournalArticle",
    "article_in_periodical": "ArticleInAPeriodical",
    "conference_proceedings": "ConferenceProceedings",
    "conference": "ConferenceProceedings",
    "report": "Report",
    "web_site": "InternetSite",
    "website": "InternetSite",
    "internet_site": "InternetSite",
    "document_from_site": "DocumentFromInternetSite",
    "electronic_source": "ElectronicSource",
    "art": "Art",
    "sound_recording": "SoundRecording",
    "performance": "Performance",
    "film": "Film",
    "interview": "Interview",
    "patent": "Patent",
    "case": "Case",
    "misc": "Misc",
}


def _parse_author(name: str) -> tuple[str, str]:
    """Split one author name into (last, first). Accepts "Last, First" or "First Last"."""
    name = name.strip()
    if "," in name:
        last, _, first = name.partition(",")
        return last.strip(), first.strip()
    parts = name.rsplit(" ", 1)
    if len(parts) == 2:
        first, last = parts
        return last.strip(), first.strip()
    return name, ""


def _parse_authors(author: str | list[str] | None) -> list[tuple[str, str]]:
    if author is None:
        return []
    names = [author] if isinstance(author, str) else list(author)
    return [_parse_author(str(n)) for n in names if str(n).strip()]


def _auto_tag(authors: list[tuple[str, str]], year: Any) -> str:
    """Word's own convention: first-author-last + year, alphanumerics only."""
    base = authors[0][0] if authors else ""
    base = re.sub(r"[^A-Za-z0-9]", "", base)
    yr = re.sub(r"[^A-Za-z0-9]", "", str(year)) if year is not None else ""
    return base + yr


def build_source_xml(
    source_type: str,
    *,
    tag: str,
    authors: list[tuple[str, str]] | None = None,
    fields: dict[str, Any] | None = None,
    extra: dict[str, Any] | None = None,
) -> str:
    """Assemble a ``<b:Source>`` element from a tag, authors, and `b:` fields.

    `fields` keys are `b:` element names (e.g. ``"Title"``); `extra` is an
    additional passthrough of `b:` element names for fields not in the friendly
    set. Text is XML-escaped. Mirrors the exact shape Word round-trips.
    """
    st = _SOURCE_TYPES.get(str(source_type).lower())
    if st is None:
        raise ValueError(
            f"unknown source type {source_type!r}; expected one of {sorted(_SOURCE_TYPES)}"
        )
    parts = [f"<b:SourceType>{st}</b:SourceType>", f"<b:Tag>{escape(tag)}</b:Tag>"]
    for last, first in authors or []:
        person = f"<b:Last>{escape(last)}</b:Last>"
        if first:
            person += f"<b:First>{escape(first)}</b:First>"
        parts.append(
            "<b:Author><b:Author><b:NameList>"
            f"<b:Person>{person}</b:Person>"
            "</b:NameList></b:Author></b:Author>"
        )
    for elem, value in {**(fields or {}), **(extra or {})}.items():
        if value is None or str(value) == "":
            continue
        if not re.fullmatch(r"[A-Za-z][A-Za-z0-9]*", elem):
            raise ValueError(f"invalid bibliography field name {elem!r}")
        parts.append(f"<b:{elem}>{escape(str(value))}</b:{elem}>")
    return f'<b:Source xmlns:b="{_NS}">{"".join(parts)}</b:Source>'


def _tag_from_xml(xml: str) -> str | None:
    try:
        root = ET.fromstring(xml)
    except ET.ParseError as e:
        raise ValueError(f"source XML is not well-formed: {e}") from e
    el = root.find(f"{{{_NS}}}Tag")
    return el.text if el is not None and el.text else None


def _source_by_tag(sources_com: Any, tag: str) -> Any | None:
    """Find a COM `Source` by tag. Returns None if missing."""
    if not tag:
        return None
    for i in range(1, int(sources_com.Count) + 1):
        src = sources_com(i)
        if str(src.Tag) == tag:
            return src
    return None


class Source:
    """One bibliography source in `doc.sources`, addressed by its tag."""

    def __init__(self, doc: Document, tag: str, *, com: Any | None = None) -> None:
        self._doc = doc
        self._tag = tag
        # A freshly added source caches its live COM object; named lookups go
        # through `_source_by_tag` (mirrors ContentControl).
        self._src_com = com

    def _src(self) -> Any:
        if self._src_com is not None:
            return self._src_com
        src = _source_by_tag(self._doc.com.Bibliography.Sources, self._tag)
        if src is None:
            raise AnchorNotFoundError("source", self._tag)
        return src

    @property
    def com(self) -> Any:
        """Raw COM `Source` object — the escape hatch."""
        return self._src()

    @property
    def tag(self) -> str:
        """The source's tag — the id a citation references."""
        return self._tag

    @property
    def cited(self) -> bool:
        """Whether a citation in the document currently references this source."""
        with _com.translate_com_errors():
            return bool(self._src().Cited)

    @property
    def xml(self) -> str:
        """The source's ``<b:Source>`` OOXML."""
        with _com.translate_com_errors():
            return str(self._src().XML)

    def delete(self) -> None:
        """Remove the source from the document's store. Wrap in `doc.edit(...)`."""
        with _com.translate_com_errors():
            self._src().Delete()

    def to_dict(self) -> dict[str, Any]:
        return {"tag": self._tag, "cited": self.cited, "xml": self.xml}

    def __repr__(self) -> str:
        return f"<Source tag={self._tag!r}>"


class SourceCollection:
    """The document's bibliography sources — `doc.sources`.

    Mirrors [`doc.content_controls`][wordlive.Document.content_controls]:
    `add(...)` / `add_xml(...)` create sources, and the collection is
    subscriptable (`doc.sources["Smith2020"]`), iterable, and `in`-testable by
    tag.
    """

    def __init__(self, doc: Document) -> None:
        self._doc = doc

    def _sources(self) -> Any:
        return self._doc.com.Bibliography.Sources

    def add(
        self,
        source_type: str = "book",
        *,
        tag: str | None = None,
        author: str | list[str] | None = None,
        title: str | None = None,
        year: str | int | None = None,
        publisher: str | None = None,
        city: str | None = None,
        journal_name: str | None = None,
        volume: str | None = None,
        issue: str | None = None,
        pages: str | None = None,
        url: str | None = None,
        edition: str | None = None,
        doi: str | None = None,
        **extra: Any,
    ) -> Source:
        """Add a source to the document's store and return it.

        `source_type` is ``"book"`` (default), ``"journal_article"``,
        ``"conference_proceedings"``, ``"report"``, ``"web_site"``, ``"case"``,
        … (see `_SOURCE_TYPES`). `author` is ``"Last, First"`` (or ``"First
        Last"``, or a list of either). `tag` is the id citations reference; if
        omitted it's auto-derived from the first author's surname + year (Word's
        own convention). The remaining keywords map to bibliography fields;
        `**extra` passes any other `b:` element through verbatim (e.g.
        ``Medium="Web"``). Wrap in `doc.edit(...)` for atomic undo. Bad input
        raises `OpError`.
        """
        try:
            authors = _parse_authors(author)
            resolved_tag = str(tag).strip() if tag else _auto_tag(authors, year)
            if not resolved_tag:
                raise ValueError("cannot auto-generate a tag; pass tag= (or author= and year=)")
            fields = {
                "Title": title,
                "Year": year,
                "City": city,
                "Publisher": publisher,
                "JournalName": journal_name,
                "Volume": volume,
                "Issue": issue,
                "Pages": pages,
                "Edition": edition,
                "DOI": doi,
                "URL": url,
            }
            xml = build_source_xml(
                source_type, tag=resolved_tag, authors=authors, fields=fields, extra=extra
            )
            with _com.translate_com_errors():
                self._sources().Add(xml)
            return Source(self._doc, resolved_tag)
        except (ValueError, TypeError) as e:
            raise OpError(str(e)) from e

    def add_xml(self, xml: str) -> Source:
        """Add a source from a raw ``<b:Source>`` element and return it.

        The escape hatch when you already have OOXML. The element must carry a
        ``<b:Tag>``. Wrap in `doc.edit(...)` for atomic undo.
        """
        try:
            tag = _tag_from_xml(xml)
            if not tag:
                raise ValueError("source XML must contain a non-empty <b:Tag>")
            with _com.translate_com_errors():
                self._sources().Add(xml)
            return Source(self._doc, tag)
        except (ValueError, TypeError) as e:
            raise OpError(str(e)) from e

    def list(self) -> list[str]:
        """The tags of every source in the document's store."""
        with _com.translate_com_errors():
            srcs = self._sources()
            return [str(srcs(i).Tag) for i in range(1, int(srcs.Count) + 1)]

    def __getitem__(self, tag: str) -> Source:
        with _com.translate_com_errors():
            if _source_by_tag(self._sources(), tag) is None:
                raise AnchorNotFoundError("source", tag)
        return Source(self._doc, tag)

    def __contains__(self, tag: object) -> bool:
        if not isinstance(tag, str):
            return False
        with _com.translate_com_errors():
            return _source_by_tag(self._sources(), tag) is not None

    def __iter__(self) -> Iterator[Source]:
        for tag in self.list():
            yield Source(self._doc, tag)

    def __len__(self) -> int:
        with _com.translate_com_errors():
            return int(self._sources().Count)
