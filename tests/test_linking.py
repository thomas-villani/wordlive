"""Anchoring & linking — bookmark creation, hyperlinks, cross-references, captions.

Round-trips against the `fake_word` MagicMock. Hyperlinks land via
`doc.Hyperlinks.Add` (an auto MagicMock); cross-references and captions collapse
a *duplicate* of the anchor's range, so assertions read
`<range>.Duplicate.InsertCrossReference` / `.InsertCaption`. `GetCrossReferenceItems`
is wired in conftest to return the seeded bookmark names / heading texts.
"""

from __future__ import annotations

import json

import pytest

import wordlive
from wordlive._ops import run_batch
from wordlive.cli.main import EXIT_OK, main
from wordlive.constants import WdReferenceKind, WdReferenceType
from wordlive.exceptions import AnchorNotFoundError, OpError


def _invoke(args):
    from click.testing import CliRunner

    result = CliRunner().invoke(main, args, catch_exceptions=False)
    return result.exit_code, result.stdout


def _addr_dup(fake_word):
    """The Duplicate range of bookmark:Address (the insert target for xref/caption)."""
    return fake_word.ActiveDocument.Bookmarks("Address").Range.Duplicate


# --- bookmark creation ---------------------------------------------------------


def test_bookmark_add_creates(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        with doc.edit("bm"):
            bm = doc.bookmarks.add("Ref1", "heading:1")
        assert isinstance(bm, wordlive.Bookmark)
        assert "Ref1" in doc.bookmarks


def test_bookmark_add_accepts_anchor_object(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        doc.bookmarks.add("Ref2", doc.anchor_by_id("heading:1"))
        assert "Ref2" in doc.bookmarks


@pytest.mark.parametrize("bad", ["has space", "1leading", "no-dashes", "", "x" * 41])
def test_bookmark_add_invalid_name_raises(fake_word, bad):
    with wordlive.attach() as word:
        doc = word.documents.active
        with pytest.raises(OpError):
            doc.bookmarks.add(bad, "heading:1")


def test_exec_add_bookmark(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        result, exc = run_batch(
            doc,
            [{"op": "add_bookmark", "name": "Ref3", "anchor_id": "heading:1"}],
            label="t",
        )
    assert exc is None and result["ok"] is True
    assert result["outputs"][0]["bookmark"] == "Ref3"


def test_cli_bookmark_add(fake_word):
    code, out = _invoke(["--json", "bookmark", "add", "Ref4", "--anchor-id", "heading:1"])
    assert code == EXIT_OK
    assert json.loads(out)["bookmark"] == "Ref4"


def test_cli_bookmark_add_bad_name(fake_word):
    code, _ = _invoke(["--json", "bookmark", "add", "bad name", "--anchor-id", "heading:1"])
    assert code != EXIT_OK


# --- hyperlinks ----------------------------------------------------------------


def test_link_to_url(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        with doc.edit("link"):
            doc.bookmarks["Address"].link_to(address="https://example.com")
    add = fake_word.ActiveDocument.Hyperlinks.Add
    add.assert_called_once()
    # (Anchor, Address, SubAddress, ScreenTip)
    assert add.call_args.args[1] == "https://example.com"
    assert add.call_args.args[2] == ""


def test_link_to_bookmark(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        doc.bookmarks["Address"].link_to(bookmark="Target")
    add = fake_word.ActiveDocument.Hyperlinks.Add
    assert add.call_args.args[1] == ""  # no external address
    assert add.call_args.args[2] == "Target"  # SubAddress = bookmark


def test_link_to_with_text(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        doc.bookmarks["Address"].link_to(address="https://x", text="click here")
    add = fake_word.ActiveDocument.Hyperlinks.Add
    assert add.call_args.args[4] == "click here"  # TextToDisplay


def test_link_to_requires_exactly_one_target(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        with pytest.raises(OpError):
            doc.bookmarks["Address"].link_to()
        with pytest.raises(OpError):
            doc.bookmarks["Address"].link_to(address="x", bookmark="y")


def test_exec_add_hyperlink_requires_one_target(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        result, exc = run_batch(
            doc, [{"op": "add_hyperlink", "anchor_id": "bookmark:Address"}], label="bad"
        )
    assert exc is not None and result["ok"] is False


def test_cli_link(fake_word):
    code, out = _invoke(
        ["--json", "link", "--anchor-id", "bookmark:Address", "--url", "https://example.com"]
    )
    assert code == EXIT_OK
    assert json.loads(out)["applied"]["url"] == "https://example.com"


def test_cli_link_both_targets_errors(fake_word):
    code, _ = _invoke(
        ["--json", "link", "--anchor-id", "bookmark:Address", "--url", "u", "--bookmark", "b"]
    )
    assert code != EXIT_OK


# --- cross-references ----------------------------------------------------------


def test_cross_reference_to_heading(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        with doc.edit("xref"):
            doc.bookmarks["Address"].insert_cross_reference("heading:1")
    args = _addr_dup(fake_word).InsertCrossReference.call_args.args
    # (ReferenceType, ReferenceKind, ReferenceItem, hyperlink)
    assert args[0] == int(WdReferenceType.HEADING)
    assert args[1] == int(WdReferenceKind.CONTENT_TEXT)
    assert args[2] == 1  # first heading
    assert args[3] is True


def test_cross_reference_to_bookmark_by_name(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        doc.bookmarks["Address"].insert_cross_reference("bookmark:Address", kind="page")
    args = _addr_dup(fake_word).InsertCrossReference.call_args.args
    assert args[0] == int(WdReferenceType.BOOKMARK)
    assert args[1] == int(WdReferenceKind.PAGE_NUMBER)
    assert args[2] == "Address"  # bookmarks reference by NAME, not index


def test_cross_reference_to_footnote_number(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        doc.bookmarks["Address"].insert_cross_reference("footnote:1", kind="number")
    args = _addr_dup(fake_word).InsertCrossReference.call_args.args
    assert args[0] == int(WdReferenceType.FOOTNOTE)
    assert args[1] == int(WdReferenceKind.FOOTNOTE_NUMBER)
    assert args[2] == 1


def test_cross_reference_to_footnote_text_falls_back_to_number(fake_word):
    # wdContentText is invalid for a note mark, so kind="text" → its number.
    with wordlive.attach() as word:
        doc = word.documents.active
        doc.bookmarks["Address"].insert_cross_reference("footnote:1", kind="text")
    args = _addr_dup(fake_word).InsertCrossReference.call_args.args
    assert args[1] == int(WdReferenceKind.FOOTNOTE_NUMBER)


def test_cross_reference_bad_target_raises_anchor_not_found(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        with pytest.raises(AnchorNotFoundError):
            doc.bookmarks["Address"].insert_cross_reference("bookmark:Missing")


def test_cross_reference_bad_kind_raises_operror(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        with pytest.raises(OpError):
            doc.bookmarks["Address"].insert_cross_reference("heading:1", kind="bogus")


def test_exec_insert_cross_reference(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        result, exc = run_batch(
            doc,
            [
                {
                    "op": "insert_cross_reference",
                    "anchor_id": "bookmark:Address",
                    "target": "heading:1",
                    "kind": "page",
                }
            ],
            label="t",
        )
    assert exc is None and result["ok"] is True
    assert _addr_dup(fake_word).InsertCrossReference.call_args.args[1] == int(
        WdReferenceKind.PAGE_NUMBER
    )


def test_exec_insert_cross_reference_bad_target_fails(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        result, exc = run_batch(
            doc,
            [
                {
                    "op": "insert_cross_reference",
                    "anchor_id": "bookmark:Address",
                    "target": "bookmark:Nope",
                }
            ],
            label="bad",
        )
    assert exc is not None and result["ok"] is False
    assert result["failure"]["type"] == "AnchorNotFoundError"


def test_cli_cross_ref(fake_word):
    code, out = _invoke(
        [
            "--json",
            "cross-ref",
            "--anchor-id",
            "bookmark:Address",
            "--target",
            "heading:1",
            "--kind",
            "page",
        ]
    )
    assert code == EXIT_OK
    assert json.loads(out)["applied"]["target"] == "heading:1"


# --- captions ------------------------------------------------------------------


def test_insert_caption_defaults(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        with doc.edit("cap"):
            doc.bookmarks["Address"].insert_caption()
    args = _addr_dup(fake_word).InsertCaption.call_args.args
    # (Label, Title, Position=below(1), ExcludeLabel=False)
    assert args[0] == "Figure"
    assert args[1] == ""
    assert args[2] == 1
    assert args[3] is False


def test_insert_caption_label_and_text(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        doc.bookmarks["Address"].insert_caption("Table", text="Quarterly costs")
    args = _addr_dup(fake_word).InsertCaption.call_args.args
    assert args[0] == "Table"
    assert args[1] == "Quarterly costs"


def test_exec_insert_caption(fake_word):
    with wordlive.attach() as word:
        doc = word.documents.active
        result, exc = run_batch(
            doc,
            [{"op": "insert_caption", "anchor_id": "bookmark:Address", "label": "Figure"}],
            label="t",
        )
    assert exc is None and result["ok"] is True


def test_cli_caption(fake_word):
    code, out = _invoke(
        ["--json", "caption", "--anchor-id", "bookmark:Address", "--label", "Table", "--text", "T"]
    )
    assert code == EXIT_OK
    data = json.loads(out)
    assert data["applied"]["label"] == "Table"
