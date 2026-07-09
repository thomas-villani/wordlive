#!/usr/bin/env python
"""One-shot refactor helper: decompose the god-class ``Document`` into a
``_document/`` package by mixin decomposition. Behavior-preserving.

Part of the pre-v1.0 mega-module refactor (see ``feature-plan.md`` → "Road to
v1.0", item 1). ``cli/commands.py`` and ``mcp/server.py`` were flat function
collections and split cleanly; ``Document`` (and ``Anchor``) are single huge
classes, so the fix is a mixin split: a ``DocumentCore`` spine (construction, the
``com`` handle, ``edit``, ``anchor_by_id``, the collection accessors) plus feature
mixins that inherit it, with a thin

    class Document(EditingMixin, ReadingMixin, StructureMixin, PersistenceMixin):

that multiply-inherits them. Identity and every call site are unchanged.

Usage::

    python scripts/split_document.py        # rewrites src/wordlive/_document.py
                                             # into src/wordlive/_document/*.py
    uv run ruff check --fix src/wordlive/_document/   # prune the maximal imports
    uv run ruff format src/wordlive/_document/
    uv run mypy                              # then the MANUAL step below

MANUAL FOLLOW-UP (why this isn't fully automated): the codebase pervasively types
collaborators (``RevisionCollection``, ``walk_blocks``, ``run_lint``, the anchor
constructors, …) on the *concrete* ``Document``, and there are real cross-mixin
calls (``outline`` → ``pin_outline``). So after the split, ~40 mixin/core methods
that touch that surface need the mypy mixin pattern — annotate ``self: Document``
(with ``if TYPE_CHECKING: from wordlive._document import Document`` in each module).
mypy will name every offending line; walk them and add the annotation. Do
``Document`` and ``Anchor`` together in one reviewable pass.

The block-extraction approach (verbatim line-range slices, maximal import header
pruned by ruff, ``from .X`` → ``from ..X`` depth shift for the new sub-package) is
the same recipe that split commands.py / server.py — reuse it for ``_anchors.py``,
which additionally needs its ~1000 lines of module-level helpers extracted and its
~21 peripheral classes rehomed (keep ``Anchor``'s references to its own subclasses
and to ``Cell``/``Table`` lazy, to avoid an import cycle).
"""

from __future__ import annotations

import re
from pathlib import Path

SRC = Path("src/wordlive/_document.py")
PKG = Path("src/wordlive/_document")


def main() -> None:
    lines = SRC.read_text(encoding="utf-8").splitlines(keepends=True)
    n = len(lines)

    def idx(pred) -> int:
        return next(i for i, line in enumerate(lines) if pred(line))

    future = idx(lambda s: s.startswith("from __future__"))
    markup = idx(lambda s: s.startswith("def _markup_flag"))
    class_doc = idx(lambda s: s.startswith("class Document:"))
    class_coll = idx(lambda s: s.startswith("class DocumentCollection"))

    # shared import header (shift every `from .` one level deeper for the package)
    header = "".join(
        re.sub(r"^(\s*)from \.", r"\1from ..", line) for line in lines[future:markup]
    ).rstrip() + "\n"

    # module-level pieces that live in _core.py: _markup_flag, _resolve_level_band,
    # WatermarkInfo (everything between the imports and `class Document`).
    preamble_core = "".join(lines[markup:class_doc]).rstrip() + "\n"

    # Document class docstring (between `class Document:` and its first method)
    first_method = next(
        i for i in range(class_doc + 1, n) if re.match(r"^    (def|@)", lines[i])
    )
    class_docstring = "".join(lines[class_doc + 1 : first_method]).rstrip("\n")

    # --- method blocks within the Document class body ----------------------
    def_name_re = re.compile(r"^    def\s+(\w+)")
    method_starts: list[int] = []
    for i in range(first_method, class_coll):
        if re.match(r"^    def ", lines[i]):
            j = i  # walk up over the method's decorators
            while j - 1 >= first_method and lines[j - 1].lstrip().startswith("@"):
                j -= 1
            method_starts.append(j)
    method_starts = sorted(set(method_starts))

    banner = re.compile(r"^\s*#")

    def mblock(k: int) -> tuple[str, str]:
        lo = method_starts[k]
        hi = method_starts[k + 1] if k + 1 < len(method_starts) else class_coll
        body = lines[lo:hi]
        while body and (body[-1].strip() == "" or banner.match(body[-1])):
            body.pop()
        name = next(m.group(1) for x in body if (m := def_name_re.match(x)))
        return name, "".join(body).rstrip() + "\n"

    group: dict[str, str] = {}

    def g(grp: str, *names: str) -> None:
        for nm in names:
            group[nm] = grp

    # Anything not listed here falls to DocumentCore (the spine).
    g("persistence", "save", "save_as", "export_pdf", "pin", "_existing_pin_starts",
      "pin_outline", "bibliography_style", "track_changes", "tracked_changes",
      "set_watermark", "remove_watermark", "watermark")
    g("structure", "add_table", "add_toc", "add_index", "add_bibliography",
      "add_table_of_authorities", "group_shapes")
    g("editing", "prepend", "prepend_paragraph", "append", "append_paragraph",
      "delete_paragraph", "find", "find_replace", "update_fields")
    g("reading", "outline", "between", "nearest_heading", "find_paragraphs", "stats",
      "to_markdown", "to_html", "read", "proofing", "lint", "regularize", "checkpoint",
      "changes_since", "diff", "snapshot", "snapshot_anchor")

    core_blocks: list[str] = []
    grp_blocks: dict[str, list[str]] = {k: [] for k in ("persistence", "structure", "editing", "reading")}
    for k in range(len(method_starts)):
        name, text = mblock(k)
        grp = group.get(name, "core")
        (core_blocks if grp == "core" else grp_blocks[grp]).append(text)

    PKG.mkdir(exist_ok=True)

    # _core.py
    (PKG / "_core.py").write_text(
        '"""Document spine: construction, the COM handle, the edit scope, anchor\n'
        'resolution, and the collection accessors the feature mixins build on."""\n\n'
        + header
        + "\n\n"
        + preamble_core
        + "\n\n"
        + "class DocumentCore:\n"
        + '    """Core state and primitives shared by every Document feature mixin."""\n\n'
        + "\n\n".join(core_blocks),
        encoding="utf-8",
    )

    mixins = {
        "persistence": ("PersistenceMixin", "Save/export, pins, watermark, track-changes, bibliography style."),
        "structure": ("StructureMixin", "Structural inserts: tables, TOC/index/bibliography/TOA, shape grouping."),
        "editing": ("EditingMixin", "Text editing: prepend/append/delete, find/replace, field update."),
        "reading": ("ReadingMixin", "Reads/queries/exports: outline, digests, stats, lint, checkpoint, snapshot."),
    }
    for grp, (cls, doc) in mixins.items():
        (PKG / f"_{grp}.py").write_text(
            f'"""{doc}"""\n\n'
            + header
            + "\nfrom ._core import DocumentCore\n\n\n"
            + f"class {cls}(DocumentCore):\n"
            + f'    """{doc}"""\n\n'
            + "\n\n".join(grp_blocks[grp]),
            encoding="utf-8",
        )

    # __init__.py — assemble Document + keep DocumentCollection, re-export
    doc_collection = "".join(lines[class_coll:n]).rstrip() + "\n"
    (PKG / "__init__.py").write_text(
        '"""Document wrapper + DocumentCollection (composed from _document/*)."""\n\n'
        + header
        + "\n"
        + "from ._core import DocumentCore, WatermarkInfo\n"
        + "from ._editing import EditingMixin\n"
        + "from ._persistence import PersistenceMixin\n"
        + "from ._reading import ReadingMixin\n"
        + "from ._structure import StructureMixin\n"
        + "\n"
        + '__all__ = ["Document", "DocumentCollection", "WatermarkInfo"]\n\n\n'
        + "class Document(EditingMixin, ReadingMixin, StructureMixin, PersistenceMixin):\n"
        + class_docstring
        + "\n\n\n"
        + doc_collection,
        encoding="utf-8",
    )

    SRC.unlink()
    print("core methods:", len(core_blocks), "| mixin counts:", {k: len(v) for k, v in grp_blocks.items()})
    print("total methods:", len(method_starts))


if __name__ == "__main__":
    main()
