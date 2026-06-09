"""Filesystem path policy for the gated CLI / MCP surfaces.

wordlive's Python API is trusted and **ungated** — `doc.save_as(...)` writes
wherever the caller says. The CLI and MCP surfaces are different: their inputs
can be prompt-injected, so every filesystem path they accept runs through a
[`PathPolicy`][wordlive._paths.PathPolicy] first.

Two boundaries:

- **Write targets** (save / save-as / export-pdf) — a *default-deny* directory
  whitelist. With no directories configured, saving is off; an operator opts in
  with `--save-dir` / `WORDLIVE_SAVE_DIRS`. Containment resolves the target
  **first** (so `..` and symlinks can't escape) and then requires it to sit
  inside an allowed directory.
- **Image-source reads** (`insert_image --path`) — a non-local *rejection* that
  runs before any filesystem probe (a UNC path's own `is_file()` check triggers
  SMB/NTLM auth, the sharpest threat), plus an optional image-directory
  allowlist (`--image-dir` / `WORDLIVE_IMAGE_DIRS`).

All refusals raise [`PathNotAllowedError`][wordlive.exceptions.PathNotAllowedError].
"""

from __future__ import annotations

import os
import re
from collections.abc import Iterable, Sequence
from pathlib import Path
from typing import Any

from .exceptions import PathNotAllowedError

# A URL/scheme prefix: a 2+ character scheme followed by a colon (``http:``,
# ``https:``, ``file:``, ``ftp:``, …). A Windows drive letter (``C:``) is a
# single character, so it never matches — drive paths stay local.
_SCHEME_RE = re.compile(r"^[A-Za-z][A-Za-z0-9+.\-]+:")


def _split_dirs(raw: str | None) -> list[str]:
    """Split an ``os.pathsep``-separated directory list (env-var form)."""
    if not raw:
        return []
    return [seg for seg in raw.split(os.pathsep) if seg.strip()]


def reject_nonlocal_image_path(raw: str | os.PathLike[str]) -> None:
    """Raise if `raw` is a non-local image path (UNC or URL/`file://`).

    Pure string inspection — performs **no** filesystem access — so it is safe to
    call before the `is_file()` probe in `_images.image_on_disk`, which is itself
    the attack surface: probing a UNC path (`\\\\host\\share\\x.png`) makes
    Windows authenticate to the remote SMB server, leaking NTLM credentials. URLs
    (`http://…`, `file://…`) are rejected too — `AddPicture` would otherwise
    fetch them, an SSRF / local-file-disclosure vector. Local drive paths
    (`C:\\img.png`) and relative paths pass through untouched.
    """
    s = str(raw)
    if s.startswith("\\\\") or s.startswith("//"):
        raise PathNotAllowedError(
            f"refusing a UNC image path ({s!r}): a network path is rejected on the "
            "CLI/MCP surface (it can leak credentials); pass a local file, or bytes/base64"
        )
    if _SCHEME_RE.match(s):
        raise PathNotAllowedError(
            f"refusing a non-local image source ({s!r}): URLs and file:// are rejected "
            "on the CLI/MCP surface; pass a local file path, or bytes/base64"
        )


class PathPolicy:
    """Default-deny filesystem policy for the gated CLI / MCP surfaces.

    `save_dirs` is the whitelist that gates write targets; an **empty** whitelist
    means saving is disabled. `image_dirs` is an optional allowlist that further
    restricts image-source paths (which are *always* screened for non-local UNC /
    URL forms, allowlist or not).
    """

    def __init__(
        self,
        save_dirs: Iterable[str | os.PathLike[str]] = (),
        image_dirs: Iterable[str | os.PathLike[str]] = (),
    ) -> None:
        self._save_dirs = tuple(self._normalise(d) for d in save_dirs)
        self._image_dirs = tuple(self._normalise(d) for d in image_dirs)

    @staticmethod
    def _normalise(d: str | os.PathLike[str]) -> Path:
        return Path(d).expanduser().resolve()

    @classmethod
    def from_env(
        cls,
        *,
        extra_save: Iterable[str | os.PathLike[str]] = (),
        extra_image: Iterable[str | os.PathLike[str]] = (),
        save_env: str = "WORDLIVE_SAVE_DIRS",
        image_env: str = "WORDLIVE_IMAGE_DIRS",
    ) -> PathPolicy:
        """Build a policy from env vars, with explicit (CLI flag) dirs merged in.

        Both the env var (an `os.pathsep`-separated list) and the explicit
        `extra_*` directories (the repeatable `--save-dir` / `--image-dir` CLI
        flags) contribute to the whitelist.
        """
        save = [*_split_dirs(os.environ.get(save_env)), *extra_save]
        image = [*_split_dirs(os.environ.get(image_env)), *extra_image]
        return cls(save_dirs=save, image_dirs=image)

    @property
    def save_dirs(self) -> tuple[Path, ...]:
        return self._save_dirs

    @property
    def image_dirs(self) -> tuple[Path, ...]:
        return self._image_dirs

    @property
    def saving_enabled(self) -> bool:
        return bool(self._save_dirs)

    def resolve_save_target(self, target: str | os.PathLike[str]) -> Path:
        """Resolve `target` to an absolute path inside the save whitelist, or raise.

        Resolves `target` **first** (collapsing `..` and following symlinks) so a
        crafted path can't escape the whitelist, then requires the result to sit
        inside one of the configured save directories. An empty whitelist raises
        unconditionally — saving is off until an operator opts in.
        """
        resolved = Path(target).expanduser().resolve()
        if not self._save_dirs:
            raise PathNotAllowedError(
                f"saving is disabled: no save directories are configured "
                f"(set WORDLIVE_SAVE_DIRS or pass --save-dir) — refused {str(target)!r}"
            )
        for d in self._save_dirs:
            if resolved == d or resolved.is_relative_to(d):
                return resolved
        allowed = ", ".join(str(d) for d in self._save_dirs)
        raise PathNotAllowedError(
            f"{str(resolved)!r} is outside the allowed save directories ({allowed})"
        )

    def screen_image_path(self, raw: str | os.PathLike[str]) -> None:
        """Vet an image-source path: reject non-local forms, then enforce the allowlist.

        Always rejects UNC / URL sources (the credential-leak / SSRF vectors)
        *before* any filesystem access. If an image-directory allowlist is
        configured, additionally requires the path to resolve inside it; with no
        allowlist, any local path is accepted (matching today's behaviour, just
        with the non-local hole closed).
        """
        reject_nonlocal_image_path(raw)
        if not self._image_dirs:
            return
        resolved = Path(raw).expanduser().resolve()
        for d in self._image_dirs:
            if resolved == d or resolved.is_relative_to(d):
                return
        allowed = ", ".join(str(d) for d in self._image_dirs)
        raise PathNotAllowedError(
            f"{str(resolved)!r} is outside the allowed image directories ({allowed})"
        )

    def screen_op_image_paths(self, ops: Sequence[Any]) -> None:
        """Screen every `insert_image` op's `path` field in a batch.

        The `exec` / `word_exec` boundary calls this before running the batch, so
        an `insert_image` op carrying a non-local or out-of-allowlist `path` is
        rejected up front (a base64/bytes image carries no `path` and is left
        alone). Ops are screened in order; the first bad path raises.
        """
        for op in ops:
            if (
                isinstance(op, dict)
                and op.get("op") == "insert_image"
                and op.get("path") is not None
            ):
                self.screen_image_path(op["path"])
