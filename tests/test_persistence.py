"""Persistence (save / save-as / export-pdf), the path policy, and image hardening.

The Python API is ungated; the CLI / MCP surfaces gate writes behind a
default-deny directory whitelist and reject non-local image-source paths.
Also covers the verb-first bookmark / section CLI consolidation.
"""

from __future__ import annotations

import json

import pytest
from click.testing import CliRunner

from wordlive import attach
from wordlive._paths import PathPolicy, reject_nonlocal_image_path
from wordlive.cli.main import EXIT_OK, EXIT_OTHER, main
from wordlive.exceptions import OpError, PathNotAllowedError
from wordlive.mcp._worker import InlineWorker
from wordlive.mcp.server import _write_impl

_BS = chr(92)  # a single backslash, unambiguous through every quoting layer
_UNC = _BS + _BS + "host" + _BS + "share" + _BS + "a.png"


def _invoke(args: list[str], *, input: str | None = None) -> tuple[int, str, str]:
    runner = CliRunner()
    result = runner.invoke(main, args, input=input, catch_exceptions=False)
    return result.exit_code, result.stdout, result.stderr


def _active():
    """The Document wrapper for the fake_word active document."""
    with attach() as word:
        return word.documents.active


# ---------------------------------------------------------------------------
# PathPolicy
# ---------------------------------------------------------------------------


class TestPathPolicy:
    def test_deny_all_by_default(self, tmp_path):
        with pytest.raises(PathNotAllowedError, match="saving is disabled"):
            PathPolicy().resolve_save_target(tmp_path / "x.docx")

    def test_inside_whitelist(self, tmp_path):
        pol = PathPolicy(save_dirs=[tmp_path])
        assert pol.resolve_save_target(tmp_path / "out.docx") == (tmp_path / "out.docx").resolve()

    def test_escape_via_dotdot_blocked(self, tmp_path):
        allowed = tmp_path / "allowed"
        allowed.mkdir()
        pol = PathPolicy(save_dirs=[allowed])
        with pytest.raises(PathNotAllowedError, match="outside"):
            pol.resolve_save_target(allowed / ".." / "evil.docx")

    def test_outside_whitelist(self, tmp_path):
        pol = PathPolicy(save_dirs=[tmp_path / "a"])
        with pytest.raises(PathNotAllowedError, match="outside"):
            pol.resolve_save_target(tmp_path / "b" / "x.docx")

    def test_saving_enabled_flag(self, tmp_path):
        assert PathPolicy().saving_enabled is False
        assert PathPolicy(save_dirs=[tmp_path]).saving_enabled is True

    def test_from_env(self, monkeypatch, tmp_path):
        monkeypatch.setenv("WORDLIVE_SAVE_DIRS", str(tmp_path))
        pol = PathPolicy.from_env()
        assert pol.saving_enabled
        assert pol.resolve_save_target(tmp_path / "x.docx") == (tmp_path / "x.docx").resolve()

    def test_from_env_merges_extra(self, monkeypatch, tmp_path):
        monkeypatch.delenv("WORDLIVE_SAVE_DIRS", raising=False)
        pol = PathPolicy.from_env(extra_save=[tmp_path])
        assert pol.resolve_save_target(tmp_path / "x.docx")


@pytest.mark.parametrize(
    "bad",
    [_UNC, "//host/share/a.png", "http://example.com/a.png", "https://x/a.png", "file:///etc/x"],
)
def test_reject_nonlocal_image_path(bad):
    with pytest.raises(PathNotAllowedError):
        reject_nonlocal_image_path(bad)


@pytest.mark.parametrize("ok", [r"C:\img\a.png", "rel/a.png", "/abs/unix/a.png", "a.png"])
def test_local_image_paths_pass(ok):
    reject_nonlocal_image_path(ok)  # must not raise


class TestImageScreening:
    def test_allowlist_inside_ok(self, tmp_path):
        PathPolicy(image_dirs=[tmp_path]).screen_image_path(tmp_path / "a.png")

    def test_allowlist_outside_blocked(self, tmp_path):
        outside = tmp_path / "other.png"
        pol = PathPolicy(image_dirs=[tmp_path / "imgs"])
        with pytest.raises(PathNotAllowedError, match="outside"):
            pol.screen_image_path(outside)

    def test_no_allowlist_accepts_local(self, tmp_path):
        PathPolicy().screen_image_path(tmp_path / "a.png")  # local always ok

    def test_op_scan_skips_base64(self):
        PathPolicy().screen_op_image_paths([{"op": "insert_image", "base64": "AAAA"}])

    def test_op_scan_rejects_unc_path(self):
        with pytest.raises(PathNotAllowedError):
            PathPolicy().screen_op_image_paths([{"op": "insert_image", "path": _UNC}])

    def test_op_scan_ignores_other_ops(self):
        PathPolicy().screen_op_image_paths([{"op": "append", "text": "hi"}])


# ---------------------------------------------------------------------------
# Document API (ungated) — against the fake document
# ---------------------------------------------------------------------------


class TestDocumentPersistence:
    def test_saved_property(self, fake_word):
        doc = _active()
        doc.com.Saved = True
        assert doc.saved is True
        doc.com.Saved = False
        assert doc.saved is False

    def test_save_writes_when_path_exists(self, fake_word, tmp_path):
        doc = _active()
        doc.com.Path = str(tmp_path)
        doc.com.FullName = str(tmp_path / "Test.docx")
        assert doc.save() == str(tmp_path / "Test.docx")
        doc.com.Save.assert_called_once()

    def test_save_refuses_never_saved(self, fake_word):
        doc = _active()
        doc.com.Path = ""
        with pytest.raises(OpError, match="never been saved"):
            doc.save()

    def test_save_as_calls_saveas2(self, fake_word, tmp_path):
        doc = _active()
        written = doc.save_as(tmp_path / "out.docx")
        assert written == str((tmp_path / "out.docx").resolve())
        doc.com.SaveAs2.assert_called_once()
        # docx format constant (wdFormatDocumentDefault == 16)
        assert doc.com.SaveAs2.call_args.kwargs["FileFormat"] == 16

    def test_save_as_refuses_overwrite(self, fake_word, tmp_path):
        existing = tmp_path / "out.docx"
        existing.write_text("x")
        doc = _active()
        with pytest.raises(OpError, match="overwrite"):
            doc.save_as(existing)
        doc.save_as(existing, overwrite=True)  # explicit override is allowed

    def test_save_as_rejects_pdf_format(self, fake_word, tmp_path):
        doc = _active()
        with pytest.raises(OpError, match="export_pdf"):
            doc.save_as(tmp_path / "out.pdf", fmt="pdf")

    def test_export_pdf_calls_com(self, fake_word, tmp_path):
        doc = _active()
        written = doc.export_pdf(tmp_path / "out.pdf")
        assert written == str((tmp_path / "out.pdf").resolve())
        doc.com.ExportAsFixedFormat.assert_called_once()


# ---------------------------------------------------------------------------
# CLI — gated save / save-as / export-pdf and image screening
# ---------------------------------------------------------------------------


class TestPersistenceCli:
    def test_save_as_deny_all(self, fake_word, tmp_path):
        code, _, err = _invoke(["save-as", str(tmp_path / "x.docx")])
        assert code == EXIT_OTHER
        assert "saving is disabled" in err

    def test_save_as_whitelisted(self, fake_word, tmp_path):
        code, out, _ = _invoke(["--save-dir", str(tmp_path), "save-as", str(tmp_path / "x.docx")])
        assert code == EXIT_OK
        data = json.loads(out)
        assert data["ok"] is True
        assert data["format"] == "docx"

    def test_save_as_outside_whitelist(self, fake_word, tmp_path):
        allowed = tmp_path / "ok"
        allowed.mkdir()
        code, _, err = _invoke(["--save-dir", str(allowed), "save-as", str(tmp_path / "x.docx")])
        assert code == EXIT_OTHER
        assert "outside" in err

    def test_export_pdf_whitelisted(self, fake_word, tmp_path):
        code, out, _ = _invoke(["--save-dir", str(tmp_path), "export-pdf", str(tmp_path / "x.pdf")])
        assert code == EXIT_OK
        assert json.loads(out)["ok"] is True

    def test_save_env_whitelist(self, fake_word, tmp_path, monkeypatch):
        monkeypatch.setenv("WORDLIVE_SAVE_DIRS", str(tmp_path))
        code, out, _ = _invoke(["save-as", str(tmp_path / "x.docx")])
        assert code == EXIT_OK

    def test_insert_image_rejects_unc(self, fake_word):
        code, _, err = _invoke(
            ["insert-image", "--anchor-id", "end", "--wrap", "inline", "--path", _UNC]
        )
        assert code == EXIT_OTHER
        assert "UNC" in err or "non-local" in err

    def test_exec_rejects_unc_image(self, fake_word):
        ops = json.dumps(
            {"ops": [{"op": "insert_image", "anchor_id": "end", "wrap": "inline", "path": _UNC}]}
        )
        code, _, err = _invoke(["exec", "--ops", "-"], input=ops)
        assert code == EXIT_OTHER
        assert "UNC" in err or "non-local" in err


# ---------------------------------------------------------------------------
# MCP — save commands honour the policy
# ---------------------------------------------------------------------------


class TestPersistenceMcp:
    def test_save_as_deny_all(self, fake_word):
        with pytest.raises(PathNotAllowedError):
            _write_impl(InlineWorker(), "save_as", {"path": "C:/whatever/x.docx"})

    def test_save_as_whitelisted(self, fake_word, tmp_path):
        pol = PathPolicy(save_dirs=[tmp_path])
        out = _write_impl(InlineWorker(), "save_as", {"path": str(tmp_path / "x.docx")}, policy=pol)
        assert out["ok"] is True
        assert out["command"] == "save_as"

    def test_export_pdf_whitelisted(self, fake_word, tmp_path):
        pol = PathPolicy(save_dirs=[tmp_path])
        out = _write_impl(
            InlineWorker(), "export_pdf", {"path": str(tmp_path / "x.pdf")}, policy=pol
        )
        assert out["ok"] is True
        assert out["command"] == "export_pdf"


# ---------------------------------------------------------------------------
# Bookmark / section CLI consolidation (verb-first)
# ---------------------------------------------------------------------------


class TestBookmarkConsolidation:
    def test_write_bookmark_create(self, fake_word):
        code, out, _ = _invoke(
            ["write", "bookmark", "NewMark", "--create", "--anchor-id", "heading:1"]
        )
        assert code == EXIT_OK
        data = json.loads(out)
        assert data["created"] is True
        assert data["bookmark"] == "NewMark"

    def test_write_bookmark_text_still_works(self, fake_word):
        code, out, _ = _invoke(["write", "bookmark", "Address", "--text", "123 Main"])
        assert code == EXIT_OK
        assert json.loads(out)["anchor"]["kind"] == "bookmark"

    def test_write_bookmark_create_requires_anchor(self, fake_word):
        code, _, err = _invoke(["write", "bookmark", "X", "--create"])
        assert code == 2  # click UsageError
        assert "anchor-id" in err

    def test_write_bookmark_create_and_text_conflict(self, fake_word):
        code, _, err = _invoke(
            ["write", "bookmark", "X", "--create", "--anchor-id", "heading:1", "--text", "y"]
        )
        assert code == 2
        assert "mutually exclusive" in err

    def test_write_bookmark_needs_a_mode(self, fake_word):
        code, _, err = _invoke(["write", "bookmark", "Address"])
        assert code == 2

    def test_read_bookmark_list(self, fake_word):
        code, out, _ = _invoke(["read", "bookmark", "--list"])
        assert code == EXIT_OK
        assert "Address" in json.loads(out)

    def test_read_bookmark_name_and_list_conflict(self, fake_word):
        code, _, err = _invoke(["read", "bookmark", "Address", "--list"])
        assert code == 2

    def test_bookmark_add_alias_still_works(self, fake_word):
        code, out, _ = _invoke(["bookmark", "add", "NewMark", "--anchor-id", "heading:1"])
        assert code == EXIT_OK
        assert json.loads(out)["bookmark"] == "NewMark"


class TestSectionConsolidation:
    def test_sections_top_level(self, fake_word):
        code, out, _ = _invoke(["sections"])
        assert code == EXIT_OK
        assert isinstance(json.loads(out), list)

    def test_section_list_alias_still_works(self, fake_word):
        code, out, _ = _invoke(["section", "list"])
        assert code == EXIT_OK
        assert isinstance(json.loads(out), list)
