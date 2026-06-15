"""End-to-end CLI lifecycle against a *live* Word instance.

Unlike `test_smoke.py` (which drives the Python API in-process), this suite shells
out to the real CLI — `python -m wordlive …` as a subprocess — so it exercises the
whole stack a user/LLM actually hits: arg parsing → COM → live Word → JSON on
stdout → deterministic exit codes. One continuous document lifecycle: build via
`exec` + individual verbs, read it back, save to disk, export a PDF, then close,
reopen, and verify the content survived the round-trip.

Excluded from the default run and CI (it needs Word). Run it on a Windows box
with Word open:

    uv run pytest -m e2e

Both markers are set: `smoke` keeps it out of the default `-m "not smoke"` run,
`e2e` lets you select just this lifecycle.
"""

from __future__ import annotations

import contextlib
import json
import subprocess
import sys
from dataclasses import dataclass
from pathlib import Path
from typing import Any

import pytest

import wordlive
from wordlive.exceptions import WordNotRunningError

pytestmark = [pytest.mark.smoke, pytest.mark.e2e]


# ---------------------------------------------------------------------------
# CLI subprocess harness
# ---------------------------------------------------------------------------


@dataclass
class CliResult:
    code: int
    out: str
    err: str

    @property
    def json(self) -> Any:
        """Parse stdout as the one JSON object the CLI emits per invocation."""
        return json.loads(self.out)


def _run_cli(
    args: list[str],
    *,
    doc: str | None = None,
    save_dir: Path | None = None,
    stdin: str | None = None,
) -> CliResult:
    """Invoke `python -m wordlive [globals] <args>` and capture code/out/err.

    Global options (`--doc`, `--save-dir`) belong to the Click *group*, so they
    precede the subcommand — the harness slots them in before `args`.
    """
    cmd = [sys.executable, "-m", "wordlive"]
    if doc is not None:
        cmd += ["--doc", doc]
    if save_dir is not None:
        cmd += ["--save-dir", str(save_dir)]
    cmd += args
    proc = subprocess.run(
        cmd,
        input=stdin,
        capture_output=True,
        text=True,
        timeout=120,
    )
    return CliResult(proc.returncode, proc.stdout, proc.stderr)


# ---------------------------------------------------------------------------
# Live-Word fixtures
# ---------------------------------------------------------------------------


@pytest.fixture(scope="module")
def _word_available() -> None:
    """Skip the whole module unless we can attach to a running Word."""
    try:
        with wordlive.attach():
            pass
    except WordNotRunningError as e:
        pytest.skip(f"Word not running (open Word with a document first): {e}")


@pytest.fixture
def scratch(_word_available: None) -> Any:
    """A fresh blank document driven only by name.

    Yields the document's window name (e.g. ``Document3``) — every CLI call in the
    test targets it with `--doc NAME`, so the suite stays robust even if the user
    has other documents open. Teardown closes whatever documents we created
    (the original blank and any reopened copy) without saving.

    `created` is mutated by the test to register extra window names (e.g. the
    saved ``e2e.docx``) so they're cleaned up too.
    """
    created: list[str] = []
    with wordlive.attach() as word:
        word.com.Documents.Add()
        name = str(word.documents.active.com.Name)
        created.append(name)

    holder = {"name": name, "created": created}
    try:
        yield holder
    finally:
        with contextlib.suppress(Exception), wordlive.attach() as word:
            for doc in list(word.documents):
                if str(doc.com.Name) in created:
                    with contextlib.suppress(Exception):
                        doc.com.Close(SaveChanges=0)


# ---------------------------------------------------------------------------
# The lifecycle
# ---------------------------------------------------------------------------


def test_cli_lifecycle_build_save_reopen(scratch: dict[str, Any], tmp_path: Path) -> None:
    name = scratch["name"]

    # 1. status — our blank document is open and visible to the CLI.
    r = _run_cli(["status"], doc=name)
    assert r.code == 0, r.err
    assert any(row["name"] == name for row in r.json)

    # 2. exec — build the bulk of the document in one atomic-undo batch. This is
    #    the centrepiece: the same op vocabulary the MCP server funnels through.
    batch = {
        "label": "e2e: build report",
        "ops": [
            {"op": "append", "text": "Quarterly Report", "style": "Heading 1"},
            {
                "op": "insert_markdown",
                "anchor_id": "end",
                "markdown": (
                    "# Overview\n\n"
                    "This is the **overview** section.\n\n"
                    "- first point\n- second point"
                ),
            },
            {
                "op": "insert_section",
                "anchor_id": "end",
                "heading": "Financials",
                "body": ["Revenue grew this quarter.", "Costs held flat."],
                "level": 1,
            },
            {
                "op": "create_table",
                "anchor_id": "end",
                "data": [
                    {"Item": "Travel", "Cost": "$400"},
                    {"Item": "Lodging", "Cost": "$600"},
                ],
            },
            {"op": "set_property", "name": "Title", "value": "Q2 Report"},
            {"op": "set_variable", "name": "ClientName", "value": "Acme"},
        ],
    }
    script = tmp_path / "build.json"
    script.write_text(json.dumps(batch), encoding="utf-8")
    r = _run_cli(["exec", "--script", str(script)], doc=name)
    assert r.code == 0, r.err
    assert r.json["ok"] is True
    # create_table reports the new table's 1-based index in `outputs`.
    table_out = next(o for o in r.json["outputs"] if o["op"] == "create_table")
    assert table_out["table"] == 1
    assert (table_out["rows"], table_out["columns"]) == (3, 2)

    # 3. A standalone verb on top of the batch — fuzzy find/replace. This
    #    replaces a *whole paragraph* that sits immediately before the table; the
    #    matched span must stop at the paragraph mark, not swallow it (doing so
    #    would fuse the paragraph into the table's first cell). The table-records
    #    assertion below guards that boundary end-to-end.
    r = _run_cli(
        ["replace", "--find", "Costs held flat.", "--text", "Costs decreased."],
        doc=name,
    )
    assert r.code == 0, r.err
    # `replacements` is the list of applied replacements (one entry here).
    assert len(r.json["replacements"]) == 1

    # 4. A table edit by header name through its own verb.
    r = _run_cli(
        [
            "table",
            "append-record",
            "--table",
            "1",
            "--record",
            json.dumps({"Item": "Meals", "Cost": "$120"}),
        ],
        doc=name,
    )
    assert r.code == 0, r.err

    # 5. TOC at the top, then refresh fields so page numbers populate.
    r = _run_cli(["insert-toc", "--anchor-id", "start", "--levels", "1-2"], doc=name)
    assert r.code == 0, r.err
    r = _run_cli(["update-fields"], doc=name)
    assert r.code == 0, r.err

    # --- Reads: the document reflects everything we built -------------------

    # outline lists the three headings in document order.
    r = _run_cli(["outline"], doc=name)
    assert r.code == 0, r.err
    outline_texts = [it["text"] for it in r.json]
    assert outline_texts == ["Quarterly Report", "Overview", "Financials"]

    # find locates body text the markdown/section ops inserted.
    r = _run_cli(["find", "--text", "overview section"], doc=name)
    assert r.code == 0, r.err
    assert len(r.json) >= 1

    # the replace took: old text gone, new text present.
    assert _run_cli(["find", "--text", "Costs held flat."], doc=name).json == []
    assert len(_run_cli(["find", "--text", "Costs decreased."], doc=name).json) == 1

    # the table reads back as records, including the appended row.
    r = _run_cli(["table", "records", "1"], doc=name)
    assert r.code == 0, r.err
    assert r.json == [
        {"Item": "Travel", "Cost": "$400"},
        {"Item": "Lodging", "Cost": "$600"},
        {"Item": "Meals", "Cost": "$120"},
    ]

    # document variable + built-in property round-trip through their verbs.
    assert _run_cli(["variables", "list"], doc=name).json["ClientName"] == "Acme"
    props = _run_cli(["properties", "list"], doc=name).json
    assert props["builtin"].get("Title") == "Q2 Report"

    # stats reports real, structure-derived counts.
    r = _run_cli(["stats"], doc=name)
    assert r.code == 0, r.err
    assert r.json["tables"] == 1
    assert r.json["headings"] >= 3
    assert r.json["saved"] is False  # never-saved scratch doc

    # --- Persistence: save .docx, export PDF (both gated) -------------------

    docx = tmp_path / "e2e.docx"
    r = _run_cli(["save-as", str(docx)], doc=name, save_dir=tmp_path)
    assert r.code == 0, r.err
    assert Path(r.json["path"]) == docx.resolve()
    assert docx.is_file() and docx.stat().st_size > 0
    # save-as renames the live window to the file; track it for cleanup + reuse.
    scratch["created"].append(docx.name)
    saved_name = docx.name

    # a follow-up edit + plain `save` to the now-existing path.
    r = _run_cli(["append", "--text", "Appendix line."], doc=saved_name)
    assert r.code == 0, r.err
    r = _run_cli(["save"], doc=saved_name, save_dir=tmp_path)
    assert r.code == 0, r.err
    assert r.json["saved"] is True

    pdf = tmp_path / "e2e.pdf"
    r = _run_cli(["export-pdf", str(pdf)], doc=saved_name, save_dir=tmp_path)
    assert r.code == 0, r.err
    assert pdf.is_file()
    assert pdf.read_bytes()[:5] == b"%PDF-"

    # --- Round-trip: close, reopen from disk, verify content survived ------

    with wordlive.attach() as word:
        for d in list(word.documents):
            if str(d.com.Name) == saved_name:
                d.com.Close(SaveChanges=0)
        # reopen the .docx straight off disk
        word.com.Documents.Open(str(docx.resolve()))

    r = _run_cli(["outline"], doc=saved_name)
    assert r.code == 0, r.err
    assert [it["text"] for it in r.json] == ["Quarterly Report", "Overview", "Financials"]
    # the appended line and table data persisted through save → close → reopen.
    assert len(_run_cli(["find", "--text", "Appendix line."], doc=saved_name).json) == 1
    assert _run_cli(["table", "records", "1"], doc=saved_name).json[-1] == {
        "Item": "Meals",
        "Cost": "$120",
    }


def test_cli_exec_batch_reports_failure_and_exit_code(scratch: dict[str, Any]) -> None:
    """A batch that fails partway reports the failing op precisely, and the
    underlying anchor error maps to exit code 2 across the subprocess boundary.

    (Note: wordlive's "atomic undo" groups the successful prefix into a single
    undo step the user can Ctrl-Z — it is *not* a transactional auto-rollback,
    so the first op stays applied. The fixture discards the doc unsaved.)"""
    name = scratch["name"]
    batch = [
        {"op": "append", "text": "PARTIAL_SENTINEL_E2E"},
        {"op": "replace", "anchor_id": "bookmark:DoesNotExist", "text": "x"},
    ]
    r = _run_cli(["exec", "--ops", "-"], doc=name, stdin=json.dumps(batch))
    assert r.code == 2, r.err  # AnchorNotFoundError → EXIT_ANCHOR_NOT_FOUND
    assert r.json["ok"] is False
    assert r.json["ops_run"] == 1  # the prefix applied; it stopped at the bad op
    assert r.json["failure"]["index"] == 1
    assert r.json["failure"]["type"] == "AnchorNotFoundError"


def test_cli_exit_codes_end_to_end(scratch: dict[str, Any]) -> None:
    """Deterministic exit codes survive the subprocess boundary."""
    name = scratch["name"]
    # Missing bookmark → 2.
    r = _run_cli(["read", "bookmark", "NoSuchBookmark"], doc=name)
    assert r.code == 2
    assert "bookmark" in r.err.lower()
    # Unknown document → 1 (DocumentNotFoundError bucket).
    r = _run_cli(["outline"], doc="NoSuchDocument.docx")
    assert r.code == 1


def test_cli_save_gating_denied_outside_whitelist(scratch: dict[str, Any], tmp_path: Path) -> None:
    """save-as to a path outside any --save-dir is refused (exit 1), and writes
    nothing to disk."""
    name = scratch["name"]
    allowed = tmp_path / "allowed"
    allowed.mkdir()
    target = tmp_path / "outside" / "denied.docx"  # not under `allowed`

    r = _run_cli(["save-as", str(target)], doc=name, save_dir=allowed)
    assert r.code == 1
    assert not target.exists()
