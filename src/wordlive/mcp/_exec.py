"""`word_exec` batch implementation."""

from __future__ import annotations

import json
from typing import Any

from .. import attach
from .._ops import (
    pick_doc,
    run_batch,
)
from .._paths import PathPolicy
from ..exceptions import OpError, WordliveError
from ._worker import Worker

_OPS_EXAMPLE = '[{"op": "append_paragraph", "text": "Hello"}]'


def coerce_ops(ops: Any) -> list[dict[str, Any]]:
    """Normalise the `ops` payload to a list of op dicts.

    `word_exec` accepts a JSON-encoded *string* as well as a real array, because
    double-encoding `ops` is an easy client-side mistake and the raw pydantic
    `list_type` error it used to produce (plus an `errors.pydantic.dev` URL) told
    an agent nothing it could act on. A string that decodes to an array of
    objects is accepted; anything else raises an `OpError` that names the actual
    problem and shows the shape wanted.
    """
    if isinstance(ops, str):
        try:
            decoded = json.loads(ops)
        except json.JSONDecodeError as exc:
            raise OpError(
                f"'ops' must be an array of op objects — received a string that is not "
                f"valid JSON ({exc.msg} at position {exc.pos}). Pass the array itself, "
                f"not a JSON-encoded string: ops={_OPS_EXAMPLE}"
            ) from exc
        if not isinstance(decoded, list):
            raise OpError(
                f"'ops' must be an array of op objects — received a JSON-encoded "
                f"{type(decoded).__name__}. Pass the array itself: ops={_OPS_EXAMPLE}"
            )
        ops = decoded
    if not isinstance(ops, list):
        raise OpError(
            f"'ops' must be an array of op objects — received {type(ops).__name__}. "
            f"Example: ops={_OPS_EXAMPLE}"
        )
    bad = next(((i, o) for i, o in enumerate(ops) if not isinstance(o, dict)), None)
    if bad is not None:
        i, o = bad
        raise OpError(
            f"'ops' must be an array of op objects — ops[{i}] is a "
            f"{type(o).__name__}, not an object. Example: ops={_OPS_EXAMPLE}"
        )
    return ops


def _exec_impl(
    worker: Worker,
    ops: Any,
    *,
    doc: str | None,
    label: str | None,
    tracked: bool,
    policy: PathPolicy | None = None,
) -> tuple[dict[str, Any], WordliveError | None]:
    pol = policy if policy is not None else PathPolicy()
    # Raises OpError (not a pydantic ValidationError) before Word is touched.
    batch = coerce_ops(ops)

    def job() -> tuple[dict[str, Any], WordliveError | None]:
        with attach() as word:
            d = pick_doc(word, doc)
            # Vet image-source paths before any COM/filesystem access.
            pol.screen_op_image_paths(batch)
            return run_batch(d, batch, label=label or "MCP: exec", tracked=tracked)

    return worker.run_on_word(job)
