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
  "adds_content": false,
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
| `adds_content` | `true` when the `fix` inserts/deletes content (not just re-formats) — `regularize` withholds these unless `allow_content` |
| `observed` / `expected` | the drifted value and the target it's measured against |

A **report-only** finding (`fixable: false`) has no `fix` — it's flagging
something only you can resolve (an unresolved comment, a manual "heading" that
was never styled). `regularize` lists those under `skipped`.

A fixable finding whose fix **adds or destroys content** (inserting a caption or
notice, deleting a stray paragraph, stripping a watermark) carries
`adds_content: true`. `regularize` **withholds** those by default — a formatting
pass shouldn't silently change what the document *says* — and lists them under
`deferred`. Pass `allow_content=True` (CLI `--allow-content`) to apply them too.

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

!!! warning "Content-changing fixes are opt-in"
    By default `regularize` is a **formatting/structure** pass: fixes that change
    what the document *says* — deleting a stray paragraph, inserting a caption or
    notice, stripping a watermark — are flagged `adds_content` and held back in
    `deferred` rather than applied. Pass `allow_content=True` (CLI
    `--allow-content`) to apply them in the same atomic-undo pass. Some things stay
    report-only regardless (an unresolved comment, accepting revisions — those are
    yours to judge). And it's Track-Changes-aware: with Track Changes on, the fixes
    are recorded as tracked revisions.

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

Forty-five rules ship today. In the tables below, **on** (✅) marks the rules in the
default set, and **fix** marks whether `regularize` can repair it automatically:
✎ fixable, · report-only (yours to resolve by hand). A **✎⊕** fix adds or deletes
content, so it's applied only under `allow_content` (see the warning above). The
**tags** are what you pass to `rules=[…]` / `--rule` to select a whole cluster at once.

### Consistency — drift from the applied style

A direct override that contradicts the paragraph's own style — the formatting
someone hand-applied that the style would otherwise have supplied.

| rule | what it catches | on | fix | tags |
|---|---|:-:|:-:|---|
| `body-font-consistent` | A body paragraph whose font name is hand-set, overriding its style's font. Table cells are skipped — see `table-style-consistent`. | ✅ | ✎ | fonts |
| `heading-font-consistent` | A heading whose font name, size, or bold is hand-set, overriding its heading style. | ✅ | ✎ | headings, fonts |
| `heading-spacing-consistent` | A heading whose space-before / space-after is overridden away from its style. | ✅ | ✎ | headings, spacing |
| `mixed-run-format` | A heading whose font varies run-to-run — part of it was separately restyled. | ✅ | · | headings, fonts |
| `heading-numbering-manual` | A heading numbered by hand (`3.1 Methods`) instead of automatic numbering, so it won't renumber when sections move. | — | · | headings, structure |
| `heading-trailing-period` | A heading whose text ends in a period — most house styles drop it. | — | ✎ | headings, structure |
| `double-space` | Two or more spaces between words. | ✅ | ✎ | typography |
| `space-before-punctuation` | Whitespace sitting before a `,` `.` `;` `:` or `)`. | ✅ | ✎ | typography |
| `table-style-consistent` | A table that isn't on the document's dominant table style. | ✅ | ✎ | typography, tables |
| `hyphen-as-range` | A numeric range written with a hyphen (`1990-1995`, `pp. 10-15`) rather than an en-dash. | — | ✎ | typography, academia |
| `tabs-for-layout` | Tabs used mid-paragraph to lay out text — the job of a table or real indents. | — | · | typography |
| `hyperlink-display-is-raw-url` | A hyperlink whose whole visible text is a bare URL, where a readable label was wanted. | — | · | hyperlinks, print |
| `header-footer-consistent` | A primary header/footer whose text disagrees across the document's own (non-linked) sections. | — | · | layout |
| `leftover-highlight` | Highlighter colour left on body text. | — | ✎ | finalization |

### Structural — an objective defect

Wrong regardless of any style or house rule — a mechanical slip that will bite
in layout, numbering, or hand-off.

| rule | what it catches | on | fix | tags |
|---|---|:-:|:-:|---|
| `heading-keep-with-next` | A heading with keep-with-next off, so it can be stranded alone at the foot of a page. | ✅ | ✎ | headings, pagination |
| `table-repeat-header` | A table that breaks across a page without repeating its header row. | ✅ | ✎ | tables, pagination |
| `list-numbering-continuity` | A numbered list Word split into independent runs, so the numbering restarts at 1. | ✅ | ✎ | lists |
| `trailing-whitespace` | A paragraph that ends in spaces or tabs. | ✅ | ✎ | typography |
| `leading-whitespace` | A paragraph that starts with literal spaces or tabs (use a paragraph indent). | ✅ | ✎ | typography |
| `stray-empty-paragraph` | An empty `Normal` paragraph between content blocks — a leftover blank line. | — | ✎⊕ | typography, whitespace |
| `manual-heading-formatting` | A short, all-bold or enlarged body paragraph that reads like a heading but was never styled as one. Table cells are skipped — a bold header row is not a heading. | ✅ | · | typography, headings |
| `heading-level-skip` | An outline that jumps a level — an H1 followed by an H3 with no H2 between them. | ✅ | · | headings, structure |
| `empty-heading` | A heading paragraph with no text — a stray styled blank line that pollutes the outline. | ✅ | · | headings, structure |
| `adjacent-headings` | Two headings in a row with no body text between them (often a heading whose body was deleted). | — | · | headings, structure |
| `toc-present-and-current` | The document has top-level headings but no table-of-contents field. | — | · | headings, structure, layout |
| `broken-cross-reference` | A `REF` / `PAGEREF` field rendering Word's "Error! Reference source not found." | ✅ | · | crossref, academia |
| `caption-manual-numbering` | A `Caption` paragraph numbered with literal text instead of a `SEQ` field, so it won't renumber. | ✅ | · | captions, academia |
| `manual-line-break` | A Shift+Enter line break inside a paragraph, where a real paragraph break likely belongs. | — | · | typography |
| `xref-as-literal-text` | A body paragraph naming a figure/table by literal number ("see Figure 3") with no `REF` field to keep it in sync. | — | · | crossref, academia |
| `hyperlink-broken-internal` | An internal jump (`HYPERLINK \l`) pointing at a bookmark that no longer exists — a dead link. | ✅ | · | hyperlinks |
| `draft-watermark-present` | A text watermark (a leftover DRAFT / CONFIDENTIAL stamp) still on the document. | — | ✎⊕ | layout, finalization |
| `comments-present` | Review comments still left in the document. | — | · | finalization |
| `unaccepted-revisions` | Tracked changes that were never accepted or rejected. | — | · | finalization |
| `track-changes-on` | Track Changes is still switched on (a document-global flag). | — | · | finalization |
| `hidden-text-present` | Runs formatted as hidden text — they print and export invisibly. | — | · | finalization |
| `stale-fields` | Updatable fields (`TOC` / `SEQ` / `REF` / `PAGE`) whose rendered result may have drifted — a refresh nudge. | — | · | finalization |

### Policy — deviates from a configured target

Off in the default set; enabled by naming them, by tag, or via a profile (which
also supplies their targets). These encode a *house style* — legitimate choices
that only become "wrong" once you've declared the target.

| rule | what it catches | fix | tags | config |
|---|---|:-:|---|---|
| `body-justified` | Body paragraphs that aren't justified. | ✎ | alignment, policy | — |
| `body-line-spacing` | Body paragraphs whose line spacing isn't the profile's target. | ✎ | spacing, policy | `target` (`"single"`/`"1.5"`/`"double"`) — required |
| `table-numeric-right-align` | A table column that's mostly numbers but not right-aligned. | ✎ | tables, policy | `threshold` (default `0.8`) |
| `em-dash-usage` | An em-dash is present — flags only; the `--` swap is too opinion-laden to auto-apply. | · | typography | — |
| `page-numbers-present` | No `PAGE` field in any header or footer. | ✎⊕ | layout | — |
| `hyperlink-bare-for-print` | An external link whose visible text doesn't contain its URL, so the destination is invisible on paper. | ✎⊕ | hyperlinks, print | — |
| `document-properties-filled` | A required built-in property (Title / Author) left empty. | · | layout | `required` (list of property names; default `["Title", "Author"]`) |
| `confidentiality-notice` | A required confidentiality notice missing from every header/footer and the body. | ✎⊕ | layout, notices | `text` — required (the notice string to look for) |
| `copyright-notice` | A copyright notice missing from every header/footer and the body. | ✎⊕ | layout, notices | `text` (default `"©"`) |

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
(`hyphen-as-range`, `em-dash-usage`, `tabs-for-layout`, `manual-line-break`,
`stray-empty-paragraph`).
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

The three policy rules in this profile all fix idempotently through the same
`format_paragraph` vocabulary the rest of wordlive uses, so a profile-driven
`regularize` obeys the same one-undo, run-twice-is-a-no-op contract as the
defaults. (Other policy rules carry content-adding fixes — a page-number field,
a footer notice — flagged `✎⊕`; those apply only under `allow_content`.)

## Where next

- [Python API → Linting & regularizing](python-api.md#linting-regularizing) —
  full `lint` / `regularize` / `Finding` signatures.
- [CLI → Linting & regularizing](cli.md#linting-regularizing) — every flag and
  exit code.
- [MCP server](mcp.md) — `word_read command=lint` and
  `word_write command=regularize` for agent tool use.
- [Tutorial](tutorial.md) — the broader guided editing session this sample's
  sibling (`quarterly-report.docx`) drives.
