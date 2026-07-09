"""`word_exec` batch implementation."""

from __future__ import annotations

from typing import Any

from .. import attach
from .._ops import (
    pick_doc,
    run_batch,
)
from .._paths import PathPolicy
from ..exceptions import WordliveError
from ._worker import Worker


def _exec_impl(
    worker: Worker,
    ops: list[dict[str, Any]],
    *,
    doc: str | None,
    label: str | None,
    tracked: bool,
    policy: PathPolicy | None = None,
) -> tuple[dict[str, Any], WordliveError | None]:
    pol = policy if policy is not None else PathPolicy()

    def job() -> tuple[dict[str, Any], WordliveError | None]:
        with attach() as word:
            d = pick_doc(word, doc)
            # Vet image-source paths before any COM/filesystem access.
            pol.screen_op_image_paths(ops)
            return run_batch(d, ops, label=label or "MCP: exec", tracked=tracked)

    return worker.run_on_word(job)
