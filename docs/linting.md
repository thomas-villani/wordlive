# Linting & regularizing

Normalizing a Word document's formatting is one of the most tedious jobs in
office work: hunting down the paragraph someone left double-spaced, the heading
that got hand-set in Arial 16 instead of the style, the numeric table column
that's left-aligned, the stray space before a period. Done by hand it's an
afternoon of scrolling and clicking. wordlive turns it into two calls:

- **`lint`** reads the document and reports what's off — a severity-ranked list
  of findings, each pointing at an [anchor id](concepts.md#anchor-ids) you can
  act on.
- **`regularize`** applies the *fixable* findings in a single, atomic-undo edit.

It's built to be a power tool for both audiences: an LLM agent can call `lint`,
reason over the structured findings, and apply exactly the fixes it wants; a
human can run `wordlive regularize` and reclaim the afternoon. This page is the
guided tour — the [Python API](python-api.md#linting-regularizing),
[CLI](cli.md#linting-regularizing), and [MCP](mcp.md) pages are the exhaustive
reference.

!!! info "Politeness holds"
    `lint` is a pure read — layout rules repaginate content-neutrally, and your
    selection, scroll position, and the document's `Saved` state are left
    untouched. `regularize` wraps every fix in one `doc.edit("Regularize
    formatting")`, so **one Ctrl-Z reverts the whole pass** and the cursor never
    moves. See [Core invariants](concepts.md).

## The two verbs

`lint` answers *"what's off about this document before I hand it over?"* and
`regularize` is the write side that acts on the answer. Everything `regularize`
does, `lint` already told you it would — the fix is right there in the finding.

=== "Python"

    ```python
    import wordlive as wl

    with wl.attach() as word:
        doc = word.documents.active

        findings = doc.lint()                 # pure read — what's off?
        for f in findings:
            print(f["severity"], f["rule"], f["anchor_id"], f["fixable"])

        report = doc.regularize()             # apply the fixable ones, one undo
        print(report["ops_run"], "fixes applied")
    ```

=== "CLI"

    ```bash
    wordlive lint            # JSON array of findings on stdout
    wordlive regularize      # apply the fixable ones; prints the report
    ```

## Anatomy of a finding

Every finding is a plain dict (exported as the frozen
[`Finding`](python-api.md#wordlive.Finding) dataclass, with `.to_dict()`):

```json
{
  "rule": "space-before-punctuation",
  "kind": "consistency",
  "severity": "info",
  "anchor_id": "para:4",
  "message": "Whitespace before punctuation.",
  "fixable": true,
  "fix": {"op": "find_replace", "find": "[ \\t]+([,.;:\\)])", "text": "\\1",
          "in": "para:4", "all": true, "mode": "regex", "required": false},
  "observed": "space before , . ; : )",
  "expected": "no space before punctuation"
}
```

| field | meaning |
|---|---|
| `rule` | the rule id (stable — use it to select or suppress) |
| `kind` | `consistency` (drifted from the applied style), `structural` (an objective defect), or `policy` (deviates from a configured house-style target) |
| `severity` | `error` · `warning` · `info` — findings come back ranked worst-first |
| `anchor_id` | where it is, as an [anchor id](concepts.md#anchor-ids) you can pass to any op |
| `fixable` | whether `regularize` can fix it automatically |
| `fix` | present **iff** `fixable` — an op-shaped dict (or list of them), literally the `exec` op `regularize` runs |
| `observed` / `expected` | the drifted value and the target it's measured against |

A **report-only** finding (`fixable: false`) has no `fix` — it's flagging
something only you can resolve (an unresolved comment, a manual "heading" that
was never styled). `regularize` lists those under `skipped`.

## A guided walkthrough

We'll audit and clean a deliberately-messy sample, then layer a house style on
top. Everything below is real output from that document.

### Get the sample

The guide drives `messy-brief.docx` — a short project brief that looks like
someone typed it in a hurry.

=== "Download it"

    Grab
    [`messy-brief.docx`](https://github.com/thomas-villani/wordlive/blob/main/examples/sample/messy-brief.docx)
    from `examples/sample/` and open it in Word.

=== "Build it"

    Regenerate it from the committed script (so the binary stays reviewable —
    read the script to see exactly what's wrong with it):

    ```bash
    uv run --with python-docx python examples/sample/build_messy_brief.py
    ```

Its blemishes, one per line: a **Status** heading hand-set in Arial 16 (its
*Heading 1* style is Calibri 14); a body line with a **space before its period**
and **trailing spaces**; a line with a **double space**; a line with **leading
whitespace**; and a budget table whose **Budget** and **Used** columns are
numbers but left-aligned. Three sections: *Status* `[heading:3]`, *Budget*
`[heading:7]`, *Next Steps* `[heading:25]`.

### Step 1 — Audit

```python
findings = doc.lint()
for f in findings:
    print(f"[{f['severity']}] {f['rule']} ({f['anchor_id']}) {'✎' if f['fixable'] else '·'}")
```

```text
[warning] trailing-whitespace (para:4) ✎
[warning] leading-whitespace (para:6) ✎
[info] heading-font-consistent (para:3) ✎     ← name 'Arial' ≠ style 'Calibri'
[info] heading-font-consistent (para:3) ✎     ← size 16.0 ≠ style 14.0
[info] double-space (para:5) ✎
[info] space-before-punctuation (para:4) ✎
```

Six findings, every one fixable. The default set is deliberately conservative —
it runs the **on-by-default consistency and structural rules** and leaves the
opinionated and policy rules off (more on selection [below](#selecting-which-rules-run)).

### Step 2 — Preview

Want to see the plan before touching the document? `dry_run` runs the exact same
selection but writes nothing:

=== "Python"

    ```python
    plan = doc.regularize(dry_run=True)
    planned = [f["rule"] for f in plan["findings"] if f["fixable"]]
    print(planned)
    # ['trailing-whitespace', 'leading-whitespace', 'heading-font-consistent',
    #  'heading-font-consistent', 'double-space', 'space-before-punctuation']
    ```

=== "CLI"

    ```bash
    wordlive regularize --dry-run
    ```

!!! note "Where the plan lives on a dry run"
    On `dry_run=True`, `applied` stays `[]` (nothing was written) and the plan is
    in `findings` — the fixable ones each carry their `fix` op. `lint()` itself
    is also a fine preview: it already shows `fixable` and the `fix`.

### Step 3 — Fix

```python
report = doc.regularize()
print(sorted(f["rule"] for f in report["applied"]))
# ['double-space', 'heading-font-consistent', 'heading-font-consistent',
#  'leading-whitespace', 'space-before-punctuation', 'trailing-whitespace']
print(report["ops_run"])   # 6
```

All six fixes landed in **one undo record** — a single Ctrl-Z in Word reverts the
whole pass, and your cursor and scroll are exactly where you left them. The fixes
are **targeted and idempotent**: each writes the style's own value back as a
direct property (or rewrites just the offending text span), so running it again
is a clean no-op:

```python
doc.regularize()["applied"]   # []  — nothing left to fix
```

That idempotency is a tested invariant, and it's what makes `regularize` safe to
run in a loop or a pre-commit-style check.

!!! warning "What `regularize` will not do"
    It's a **formatting/structure** regularizer. Content-changing fixes — deleting
    a comment, inserting a real caption field, accepting revisions — are out of
    scope and always come back report-only. And it's Track-Changes-aware: with
    Track Changes on, the fixes are recorded as tracked revisions.

### Step 4 — Apply a house style

The defaults catch objective slips. A **house style** goes further: it pins
*policy* — "our body text is justified and 1.5-spaced; numeric table columns are
right-aligned." Those are the **policy rules**, and they're off until a profile
opts them in and supplies their targets. Point `lint`/`regularize` at a profile:

=== "Python"

    ```python
    profile = {
        "rules": {
            "body-justified":            {"enabled": True},
            "body-line-spacing":         {"enabled": True, "target": "1.5"},
            "table-numeric-right-align": {"enabled": True, "threshold": 0.8},
        }
    }
    doc.lint(profile=profile)          # or profile="wordlive.lint.json"
    doc.regularize(profile=profile)    # applies the policy fixes too
    ```

=== "CLI"

    ```bash
    wordlive lint --profile wordlive.lint.json
    wordlive regularize --profile wordlive.lint.json
    ```

On the sample, the profile adds **20 policy findings** to the six defaults — 7
left-aligned body paragraphs (`body-justified`), the same 7 at 1.15 line spacing
instead of 1.5 (`body-line-spacing`), and 6 left-aligned numeric cells across the
*Budget* and *Used* columns (`table-numeric-right-align`). `regularize(profile=…)`
applies all 26 in one undo record, and a second pass is still empty. See
[house-style profiles](#house-style-profiles) for the full file format.

## The rule catalog

Thirty rules ship today. `kind` sets whether a rule runs by default; `on`
marks the ones in the default set; `fix` marks the ones `regularize` can apply
automatically (the rest are report-only).

### Consistency — drift from the applied style

| rule | on | fix | tags |
|---|:-:|:-:|---|
| `body-font-consistent` | ✅ | ✎ | fonts |
| `heading-font-consistent` | ✅ | ✎ | headings, fonts |
| `heading-spacing-consistent` | ✅ | ✎ | headings, spacing |
| `mixed-run-format` | ✅ | · | headings, fonts |
| `double-space` | ✅ | ✎ | typography |
| `space-before-punctuation` | ✅ | ✎ | typography |
| `table-style-consistent` | ✅ | ✎ | typography, tables |
| `hyphen-as-range` | — | ✎ | typography, academia |
| `tabs-for-layout` | — | · | typography |
| `leftover-highlight` | — | ✎ | finalization |

### Structural — an objective defect

| rule | on | fix | tags |
|---|:-:|:-:|---|
| `heading-keep-with-next` | ✅ | ✎ | headings, pagination |
| `table-repeat-header` | ✅ | ✎ | tables, pagination |
| `list-numbering-continuity` | ✅ | ✎ | lists |
| `trailing-whitespace` | ✅ | ✎ | typography |
| `leading-whitespace` | ✅ | ✎ | typography |
| `manual-heading-formatting` | ✅ | · | typography, headings |
| `broken-cross-reference` | ✅ | · | crossref, academia |
| `caption-manual-numbering` | ✅ | · | captions, academia |
| `manual-line-break` | — | · | typography |
| `xref-as-literal-text` | — | · | crossref, academia |
| `comments-present` | — | · | finalization |
| `unaccepted-revisions` | — | · | finalization |
| `track-changes-on` | — | · | finalization |
| `hidden-text-present` | — | · | finalization |
| `stale-fields` | — | · | finalization |

### Policy — deviates from a configured target

Off in the default set; enabled by naming them, by tag, or via a profile (which
also supplies their targets).

| rule | fix | tags | config |
|---|:-:|---|---|
| `body-justified` | ✎ | alignment, policy | — |
| `body-line-spacing` | ✎ | spacing, policy | `target` (`"single"`/`"1.5"`/`"double"`) — required |
| `table-numeric-right-align` | ✎ | tables, policy | `threshold` (default `0.8`) |
| `em-dash-usage` | · | typography | — |
| `page-numbers-present` | · | layout | — |

## Selecting which rules run

With no selection, `lint`/`regularize` run the **default set**: every
on-by-default consistency and structural rule. Opinionated rules, the
finalization cluster, and policy rules stay off until you ask for them.

=== "Python"

    ```python
    doc.lint()                                  # the default set
    doc.lint(rules=["typography"])              # one whole tag cluster…
    doc.lint(rules=["em-dash-usage"])           # …or one off-by-default rule by id
    doc.lint(rules={"exclude": ["double-space"]})   # default set minus one
    doc.lint(within="heading:7")                # scope to one section
    ```

=== "CLI"

    ```bash
    wordlive lint --rule typography             # repeatable
    wordlive lint --rule em-dash-usage
    wordlive lint --exclude double-space
    wordlive lint --within heading:7
    ```

Naming a rule (by id or tag) **overrides its default-off status** — `rules=["typography"]`
lights up the entire typography cluster, including its off-by-default members
(`hyphen-as-range`, `em-dash-usage`, `tabs-for-layout`, `manual-line-break`).
`within` scopes the audit to a single anchor (`heading:N`, `range:S-E`,
`table:N:R:C`, or an [`Anchor`](python-api.md#wordlive.Anchor)).

!!! tip "The finalization pass"
    The `finalization` cluster is a ready-to-send checklist — unresolved comments,
    unaccepted revisions, Track Changes still on, hidden text, stale fields. It's
    all off by default; run it right before you send with `rules=["finalization"]`
    (or `--rule finalization`).

## House-style profiles

A **profile** is a small declarative config — an inline dict, or a
`wordlive.lint.json` file you commit next to the document. It does four things:

1. **opts policy rules in** (a bare `{"enabled": true}`),
2. **supplies their targets** (`target`, `threshold`),
3. **overrides a rule's severity**, and
4. **disables a rule** that's otherwise on by default.

```json
{
  "extends": "default",
  "rules": {
    "body-justified":            { "enabled": true, "severity": "warning" },
    "body-line-spacing":         { "enabled": true, "target": "1.5" },
    "table-numeric-right-align": { "enabled": true, "threshold": 0.8 },
    "double-space":              { "enabled": false }
  }
}
```

Pass it as a path or a dict; both `lint` and `regularize` accept `profile=`, and
resolve it once:

=== "Python"

    ```python
    doc.lint(profile="wordlive.lint.json")
    doc.regularize(profile={"rules": {"body-justified": {"enabled": True}}})
    ```

=== "CLI"

    ```bash
    wordlive lint --profile wordlive.lint.json
    wordlive regularize --profile wordlive.lint.json
    ```

| key | effect |
|---|---|
| `rules.<id>.enabled` | `true` opts a policy rule in (or a bare `<id>: {}` mention); `false` disables a rule that's otherwise on |
| `rules.<id>.target` | the target value a policy rule measures against (`body-line-spacing` needs it, or it no-ops) |
| `rules.<id>.threshold` | numeric cutoff (`table-numeric-right-align`: fraction of a column's cells that must parse numeric) |
| `rules.<id>.severity` | override the finding severity (`error`/`warning`/`info`) — retune what a lint failure means to you |
| `extends` | accepted and recorded; only `"default"` is meaningful today |

The three policy rules all fix idempotently through the same `format_paragraph`
vocabulary the rest of wordlive uses, so a profile-driven `regularize` obeys the
same one-undo, run-twice-is-a-no-op contract as the defaults.

## Where next

- [Python API → Linting & regularizing](python-api.md#linting-regularizing) —
  full `lint` / `regularize` / `Finding` signatures.
- [CLI → Linting & regularizing](cli.md#linting-regularizing) — every flag and
  exit code.
- [MCP server](mcp.md) — `word_read command=lint` and
  `word_write command=regularize` for agent tool use.
- [Tutorial](tutorial.md) — the broader guided editing session this sample's
  sibling (`quarterly-report.docx`) drives.
