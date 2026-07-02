"""The linter house-style / policy **profile** (`spec-linter.md` §6).

A profile is a small declarative config that lets the caller drive the **policy**
rules — which are off in the default set (§2: policy needs configuration). It does
three things:

- **opt a policy rule in** and supply its target/threshold
  (`"body-line-spacing": {"enabled": true, "target": "1.5"}`),
- **override a rule's severity** (`{"severity": "warning"}`),
- **disable a default rule** (`{"double-space": {"enabled": false}}`).

The `house_style` half of §6 (pinning consistency-rule targets to named style values
and fixing via `set_style`) is deferred — this ships only the policy-enabling + targets
half. `extends` is accepted and recorded but only `"default"` is meaningful today.

`doc.lint(profile=…)` / `doc.regularize(profile=…)` accept a path to a JSON file, an
inline `dict`, an existing `Profile`, or `None`; `run_lint` resolves it once via
[`Profile.load`][wordlive._lint_profile.Profile.load] and threads it to every rule's
`check` and to `_select_rules`.
"""

from __future__ import annotations

import json
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

from .exceptions import OpError


@dataclass(frozen=True)
class Profile:
    """A resolved lint profile. `rules` maps a rule id to its per-rule config dict
    (`enabled` / `severity` / `target` / `threshold` / …); `extends` is the base
    profile name, if any. An **empty** profile (`Profile()`) is the no-config
    default — every policy rule stays off and no severity is overridden."""

    rules: dict[str, dict[str, Any]] = field(default_factory=dict)
    extends: str | None = None

    @classmethod
    def load(cls, source: Any) -> Profile:
        """Resolve `source` into a `Profile`.

        `source` may be `None` (→ empty profile), an existing `Profile` (returned
        as-is), an inline mapping, or a path (`str`/`Path`) to a JSON file. Mirrors
        the tolerant config-read idiom used elsewhere in the CLI: a missing/empty
        file is an empty profile; malformed JSON or a non-object payload raises
        `OpError` (bad input, exit 1)."""
        if source is None:
            return cls()
        if isinstance(source, Profile):
            return source
        if isinstance(source, dict):
            return cls._from_mapping(source, where="profile")
        path = Path(source)
        try:
            raw = path.read_text(encoding="utf-8").strip()
        except OSError as e:
            raise OpError(f"could not read lint profile {str(source)!r}: {e}") from e
        try:
            data = json.loads(raw) if raw else {}
        except json.JSONDecodeError as e:
            raise OpError(f"lint profile {str(source)!r} is not valid JSON: {e}") from e
        if not isinstance(data, dict):
            raise OpError(f"lint profile {str(source)!r} must be a JSON object")
        return cls._from_mapping(data, where=str(source))

    @classmethod
    def _from_mapping(cls, data: dict[str, Any], *, where: str) -> Profile:
        raw_rules = data.get("rules", {})
        if not isinstance(raw_rules, dict):
            raise OpError(f"lint profile {where!r}: 'rules' must be an object")
        rules: dict[str, dict[str, Any]] = {}
        for rid, cfg in raw_rules.items():
            if cfg is None:
                rules[rid] = {}
            elif isinstance(cfg, dict):
                rules[rid] = dict(cfg)
            else:
                raise OpError(f"lint profile {where!r}: rule {rid!r} config must be an object")
        extends = data.get("extends")
        return cls(rules=rules, extends=extends if isinstance(extends, str) else None)

    def is_enabled(self, rule_id: str) -> bool | None:
        """`True`/`False` if the profile mentions `rule_id` (a bare mention enables
        it — `enabled` defaults to `True`), else `None` (the profile is silent, so
        the rule's own default stands)."""
        cfg = self.rules.get(rule_id)
        if cfg is None:
            return None
        return bool(cfg.get("enabled", True))

    def severity_for(self, rule_id: str) -> str | None:
        """The profile's severity override for `rule_id`, or `None` if unset."""
        cfg = self.rules.get(rule_id)
        if cfg is None:
            return None
        sev = cfg.get("severity")
        return sev if isinstance(sev, str) else None

    def config_for(self, rule_id: str) -> dict[str, Any]:
        """The raw per-rule config dict (so a policy rule reads its own `target` /
        `threshold`), or an empty dict if the profile doesn't mention the rule."""
        return self.rules.get(rule_id) or {}
