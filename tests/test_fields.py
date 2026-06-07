"""Fields — `Anchor.insert_field`, `HeaderFooter.insert_page_number`,
`Document.update_fields`, and their exec ops.

Round-trips against the `fake_word` MagicMock: a range's `.Fields.Add` and the
document's `.Fields.Update` are auto-created MagicMocks, so the calls are
assertable without a conftest fixture. `insert_field` collapses a *duplicate* of
the anchor's range, so assertions read `<range>.Duplicate.Fields.Add`.
"""

from __future__ import annotations

import pytest

import wordlive
from wordlive._ops import run_batch
from wordlive.cli.main import EXIT_OK, main
from wordlive.constants import WdCollapseDirection, WdFieldType
from wordlive.exceptions import OpError


def _addr_add(fake_word):
    """The `Fields.Add` mock for a field inserted at bookmark:Address."""
    return fake_word.ActiveDocument.Bookmarks("Address").Range.Duplicate.Fields.Add


# --- the Anchor method ---------------------------------------------------------


def test_insert_field_page(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        with doc.edit("field"):
            doc.bookmarks["Address"].insert_field("page")
    add = _addr_add(fake_word)
    add.assert_called_once()
    assert add.call_args.args[1] == int(WdFieldType.PAGE)


def test_insert_field_numpages(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        doc.bookmarks["Address"].insert_field("numpages")
    assert _addr_add(fake_word).call_args.args[1] == int(WdFieldType.NUM_PAGES)


def test_insert_field_date(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        doc.bookmarks["Address"].insert_field("date")
    assert _addr_add(fake_word).call_args.args[1] == int(WdFieldType.DATE)


def test_insert_field_raw_code(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        doc.bookmarks["Address"].insert_field("field", text="REF foo")
    add = _addr_add(fake_word)
    assert add.call_args.args[1] == int(WdFieldType.EMPTY)
    assert add.call_args.args[2] == "REF foo"


def test_insert_field_before_collapses_to_start(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        doc.bookmarks["Address"].insert_field("page", where="before")
    collapse = fake_word.ActiveDocument.Bookmarks("Address").Range.Duplicate.Collapse
    assert collapse.call_args.args[0] == int(WdCollapseDirection.START)


def test_insert_field_raw_without_text_raises(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        with pytest.raises(OpError):
            doc.bookmarks["Address"].insert_field("field")


def test_insert_field_bad_kind_raises(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        with pytest.raises(OpError):
            doc.bookmarks["Address"].insert_field("bogus")


def test_insert_field_bad_where_raises(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        with pytest.raises(OpError):
            doc.bookmarks["Address"].insert_field("page", where="sideways")


# --- header/footer page numbers ------------------------------------------------


def test_insert_page_number_on_footer(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        with doc.edit("page number"):
            doc.sections[1].footer().insert_page_number()
    footer_rng = fake_word.ActiveDocument.Sections(1).Footers(1).Range
    add = footer_rng.Duplicate.Fields.Add
    add.assert_called_once()
    assert add.call_args.args[1] == int(WdFieldType.PAGE)


# --- update_fields -------------------------------------------------------------


def test_update_fields_calls_com(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        with doc.edit("update"):
            doc.update_fields()
    fake_word.ActiveDocument.Fields.Update.assert_called_once()


# --- exec ops ------------------------------------------------------------------


def test_exec_insert_field(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        result, exc = run_batch(
            doc,
            [{"op": "insert_field", "anchor_id": "bookmark:Address", "kind": "numpages"}],
            label="test",
        )
    assert exc is None and result["ok"] is True
    assert _addr_add(fake_word).call_args.args[1] == int(WdFieldType.NUM_PAGES)


def test_exec_insert_field_raw_code(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        result, exc = run_batch(
            doc,
            [
                {
                    "op": "insert_field",
                    "anchor_id": "bookmark:Address",
                    "kind": "field",
                    "text": "PAGE",
                }
            ],
            label="test",
        )
    assert exc is None and result["ok"] is True
    add = _addr_add(fake_word)
    assert add.call_args.args[1] == int(WdFieldType.EMPTY)
    assert add.call_args.args[2] == "PAGE"


def test_exec_insert_field_missing_kind_fails(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        result, exc = run_batch(
            doc, [{"op": "insert_field", "anchor_id": "bookmark:Address"}], label="bad"
        )
    assert exc is not None and result["ok"] is False
    assert "kind" in result["failure"]["error"]


def test_exec_update_fields(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        result, exc = run_batch(doc, [{"op": "update_fields"}], label="test")
    assert exc is None and result["ok"] is True
    fake_word.ActiveDocument.Fields.Update.assert_called_once()


# --- CLI -----------------------------------------------------------------------


def _invoke(args):
    from click.testing import CliRunner

    result = CliRunner().invoke(main, args, catch_exceptions=False)
    return result.exit_code, result.stdout


def test_cli_insert_field(fake_word):
    import json

    code, out = _invoke(
        ["--json", "insert-field", "--anchor-id", "bookmark:Address", "--kind", "page"]
    )
    assert code == EXIT_OK
    data = json.loads(out)
    assert data["ok"] is True
    assert data["applied"]["kind"] == "page"
    assert _addr_add(fake_word).call_args.args[1] == int(WdFieldType.PAGE)


def test_cli_update_fields(fake_word):
    import json

    code, out = _invoke(["--json", "update-fields"])
    assert code == EXIT_OK
    assert json.loads(out)["ok"] is True
    fake_word.ActiveDocument.Fields.Update.assert_called_once()
