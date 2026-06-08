"""delete_paragraph — `doc.delete_paragraph`, the op, and the CLI command.

The `fake_word` document is `"Introduction\rBody text here.\rRisks\r"` (Content.End
== 35), with para:1 at 0–13, para:2 at 13–29, para:3 ("Risks") at 29–35. Deletes
go through `doc.Range(start, end).Delete()`, which the cached range factory lets
us assert against.
"""

from __future__ import annotations

import json

import pytest

import wordlive
from wordlive._ops import run_batch, validate_op
from wordlive.cli.main import EXIT_OK, main
from wordlive.exceptions import AnchorNotFoundError, OpError


def _invoke(args):
    from click.testing import CliRunner

    result = CliRunner().invoke(main, args, catch_exceptions=False)
    return result.exit_code, result.stdout


# --- the Document method -------------------------------------------------------


def test_delete_paragraph_deletes_full_range(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        with doc.edit("del"):
            doc.delete_paragraph("para:1")
    # para:1 is 0–13 and is not the terminal paragraph, so the whole range goes.
    fake_word.ActiveDocument.Range(0, 13).Delete.assert_called_once()


def test_delete_paragraph_clamps_terminal_mark(fake_word):
    # para:3 ends at 35 == Content.End; the undeletable final mark is clamped off,
    # so the delete targets 29–34, not 29–35.
    with wordlive.attach() as word:
        doc = word.documents.active
        with doc.edit("del"):
            doc.delete_paragraph("para:3")
    fake_word.ActiveDocument.Range(29, 34).Delete.assert_called_once()


def test_delete_paragraph_accepts_anchor_object(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        with doc.edit("del"):
            doc.delete_paragraph(doc.paragraphs[2])
    fake_word.ActiveDocument.Range(13, 29).Delete.assert_called_once()


def test_delete_paragraph_unknown_anchor_raises(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        with pytest.raises(AnchorNotFoundError):
            doc.delete_paragraph("para:99")


# --- the op --------------------------------------------------------------------


def test_op_delete_paragraph(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        result, exc = run_batch(doc, [{"op": "delete_paragraph", "anchor_id": "para:1"}], label="x")
    assert exc is None
    assert result["ok"] is True and result["ops_run"] == 1
    fake_word.ActiveDocument.Range(0, 13).Delete.assert_called_once()


def test_op_delete_paragraph_requires_anchor_id():
    with pytest.raises(OpError):
        validate_op({"op": "delete_paragraph"})


# --- CLI -----------------------------------------------------------------------


def test_cli_delete_paragraph(fake_word):
    code, out = _invoke(["--json", "delete-paragraph", "--anchor-id", "para:1"])
    assert code == EXIT_OK
    payload = json.loads(out)
    assert payload == {"ok": True, "anchor_id": "para:1", "deleted": True}
    fake_word.ActiveDocument.Range(0, 13).Delete.assert_called_once()
