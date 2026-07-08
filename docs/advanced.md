# Advanced: a power-features session

The [Tutorial](tutorial.md) walked the four invariants on one document. This
session picks up where it left off and drives the **power features** — the ones
that turn wordlive from "polite scripted edits" into a real document workbench:
reading a large document into a fixed token budget, pinning content so positional
ids survive your edits, embedding and formatting a chart, *seeing* the rendered
page, and enforcing a house style.

Like the tutorial it's a single continuous session with dual **Python** / **CLI**
tabs, and every step is one Ctrl-Z. It assumes you've read [Concepts](concepts.md)
and are comfortable with anchors and `doc.edit()`.

## Before you start

Use the same sample the tutorial drives —
[`quarterly-report.docx`](https://github.com/thomas-villani/wordlive/blob/main/examples/sample/quarterly-report.docx)
from `examples/sample/` (or rebuild it: `uv run --with python-docx python
examples/sample/build_quarterly_report.py`). **Open it in Word now.** Two steps
need optional extras:

```bash
pip install "wordlive[snapshot]"    # Step 4 — render pages to PNG (PyMuPDF)
# Step 3 (charts) needs Microsoft Excel installed — it's the one feature that
# reaches a second Office app.
```

## Step 1 — Read a big document without burning context

You rarely want to pour a whole document into a model. `read` (Python
`doc.read`) is a **token-budgeted digest**: headings verbatim — each tagged with
its `heading:N` anchor — tables collapsed to one-line stubs, body text sampled to
fit the budget, everything else elided to markers that *still name the `para:`
range*. So the loop is **skim cheap, then drill precisely**.

=== "Python"

    ```python
    import wordlive as wl

    with wl.attach() as word:
        doc = word.documents.active

        digest = doc.read(budget=1500)        # whole doc, elided to ~1500 tokens
        print(digest)
        # ## Introduction               <!-- heading:3 -->
        # Scope: integration milestone and Q2 budget. …
        # ## Budget                     <!-- heading:10 -->
        # > table:1 — 5 rows × 3 cols: Item | Q1 | Q2 …
        # … (12 paragraphs elided: para:14–para:26) …

        # The model picks a region from the digest; pull just that, in full.
        section = doc.to_markdown(within="heading:10")   # the Budget section
    ```

    `read` keeps every anchor addressable, so any elided region drills open with
    `to_markdown(within=…)` (or `to_html` when you need underline preserved).

=== "CLI"

    ```bash
    wordlive --text read digest --budget 1500        # whole-doc digest
    wordlive --text read markdown --within heading:10 # one section, in full
    ```

The digest is the map; `to_markdown` is the street view. An agent never has to
choose between "too little context" and "the whole document" — see
[Cookbook §20](cookbook.md#20-load-a-big-document-into-context-cheaply).

## Step 2 — Pin what you're about to move

`heading:N` and `para:N` are **positional** — they're paragraph indices, so an
insert earlier in the document renumbers everything after it. That's fine for a
single edit, but a multi-step session (or an agent making several passes) keeps
re-reading `outline` to stay in sync. A **pin** fixes that: it plants a durable
handle — a wordlive-managed hidden bookmark — that Word keeps attached to the same
content across inserts, deletes, and edits.

=== "Python"

    ```python
    with wl.attach() as word:
        doc = word.documents.active

        with doc.edit("Pin the budget block"):
            handle = doc.pin("heading:10", name="budget")   # → pin:budget
        pid = handle["anchor_id"]                            # "pin:budget"

        # Now insert *above* it — the positional id would shift, the pin doesn't.
        with doc.edit("Prepend an exec summary"):
            doc.prepend_paragraph("Executive summary.", style="Body Text")

        # pin:budget still points at the Budget heading, though it's now heading:11.
        print(doc.anchor_by_id(pid).text)       # "Budget"
    ```

    Omit `name=` for a random `pin:<code>`; reuse a slug to move the handle. Pin
    every heading at once with `doc.pin_outline()` (or `doc.outline(pin=True)`,
    which returns the outline *and* pins as a side effect). A pin vanishes if its
    content is deleted — resolving it then raises `AnchorNotFoundError`.

=== "CLI"

    ```bash
    wordlive pin heading:10 --name budget      # → {"anchor_id": "pin:budget", …}
    wordlive prepend --text "Executive summary." --style "Body Text"
    wordlive --text read text --anchor-id pin:budget   # still "Budget"

    wordlive pin-outline                       # mint a pin for every heading
    ```

Rule of thumb: **positional ids for a single read-decide-write; pins for a
multi-step session or an agent that edits then edits again.** In an `exec` batch,
add `bind: "slug"` to a creating op to pin its new content without a second call —
see [MCP → Durable handles in a batch](mcp.md#batches).

## Step 3 — Insert a chart and dress it up

`insert_chart` embeds an Excel-backed chart at any anchor, then breaks the data
link so the chart ships **static** — no live workbook travels with the document.
The formatting verbs (`add_trendline`, `set_axis`, `add_error_bars`, `format`)
operate on that static chart and need no Excel. Charts are the one feature that
reaches a second Office app, so Excel must be installed: `insert_chart` raises
[`ExcelNotAvailableError`](errors.md) (CLI exit `6`) up front, document untouched.

=== "Python"

    ```python
    with wl.attach() as word:
        doc = word.documents.active

        with doc.edit("Add revenue chart"):
            chart = doc.heading("Budget").insert_chart(
                "line",
                {"Q1": 42, "Q2": 55, "Q3": 61, "Q4": 78},   # {label: value}
                title="Quarterly revenue",
            )
            chart.add_trendline(kind="linear", display_equation=True)
            chart.set_axis("y", title="$000s", minimum=0)
    ```

    `insert_chart` returns a `ChartAnchor` (`chart:N`) and the formatting verbs
    chain. `kind` is `"bar"`, `"pie"`, `"line"`, or `"scatter"`; scatter takes
    `[x, y]` pairs instead of a mapping. Discover existing charts with
    `doc.charts`; reads report metadata only (the series data isn't read back).

=== "CLI"

    ```bash
    wordlive insert-chart --anchor-id heading:10 --kind line \
        --data '{"Q1":42,"Q2":55,"Q3":61,"Q4":78}' --title "Quarterly revenue"
    wordlive add-trendline --anchor-id chart:1 --kind linear --display-equation
    wordlive format-axis --anchor-id chart:1 --which y --title "\$000s" --minimum 0
    ```

Full depth (error bars, per-point colours, log axes, re-typing in place) is in
[Cookbook §16](cookbook.md#16-insert-a-chart-and-dress-it-up).

## Step 4 — See what you just did

Everything so far was structured I/O. But layout — did the chart land where you
meant, is that heading stranding at a page foot? — is *visual*. `snapshot` exports
a pixel-faithful PDF of the live document and rasterises the pages you ask for, so
a vision model (or you) can **look at the real page**, real fonts and geometry
included.

=== "Python"

    ```python
    with wl.attach() as word:
        doc = word.documents.active

        # The page(s) the Budget section — and its new chart — occupy.
        pages = doc.snapshot(pages=None, max_dim=1000)   # every page, long edge ≤1000px
        pages[0].png                                     # bytes → hand to a vision model

        # Or just the section you're iterating on, with revision marks visible:
        shot = doc.snapshot(out="budget.png", markup="all")
    ```

    `max_dim` caps each page's long edge — the lever for a cheap *whole-document*
    layout check (a vision model bills on pixel area, so the cap fixes a per-page
    token budget; ~1000 stays legible for "did my styling land?"). `markup="all"`
    renders tracked changes and comments as visible marks without touching the
    user's on-screen view. Read-only — the cursor never moves. Needs the
    `snapshot` extra.

=== "CLI"

    ```bash
    wordlive snapshot --anchor-id heading:10 --out budget.png   # the section's page(s)
    wordlive snapshot --max-dim 1000 --out doc.png              # whole doc, cheap
    wordlive snapshot --page 2 --markup all --out p2-marked.png
    ```

Without `--out`, the CLI returns base64 PNG data inline; over [MCP](mcp.md) the
`word_snapshot` tool returns native image content the model sees directly. This
is the feedback loop for formatting work: **edit → snapshot → look → adjust.**

## Step 5 — Enforce a house style

The [linting guide](linting.md) covers the defaults; here's the *advanced* move —
a **house-style profile** that turns on policy rules and pins their targets, then
regularizes the whole document to them in one undo. A profile is an inline dict or
a committed `wordlive.lint.json`.

=== "Python"

    ```python
    with wl.attach() as word:
        doc = word.documents.active

        profile = {
            "extends": "default",
            "rules": {
                "body-justified":            {"enabled": True},
                "body-line-spacing":         {"enabled": True, "target": "1.5"},
                "table-numeric-right-align": {"enabled": True, "threshold": 0.8},
            },
        }
        plan = doc.regularize(profile=profile, dry_run=True)   # preview
        report = doc.regularize(profile=profile)               # apply, one undo
        print(len(report["applied"]), "fixes")
        doc.regularize(profile=profile)["applied"]             # []  — idempotent
    ```

=== "CLI"

    ```bash
    wordlive lint --profile wordlive.lint.json          # what the policy flags
    wordlive regularize --profile wordlive.lint.json --dry-run
    wordlive regularize --profile wordlive.lint.json    # apply
    ```

The whole pass is one Ctrl-Z and idempotent — safe to wire into a pre-send check.
Content-changing fixes (a page-number field, a footer notice) stay withheld unless
you pass `allow_content=True` / `--allow-content`. Full file format and the 45-rule
catalog live in the [Linting guide](linting.md#house-style-profiles).

## What you just learned

| Step | Power feature | Reach for it when |
| --- | --- | --- |
| 1 | `read` digest → `to_markdown` drill | A document is too big to feed whole |
| 2 | `pin` / `pin_outline` | A multi-step session churns positional ids |
| 3 | `insert_chart` + formatting verbs | You need a static, self-contained chart |
| 4 | `snapshot` (`max_dim`, `markup`) | Layout is visual — let a model *see* it |
| 5 | `regularize(profile=…)` | A house style must be enforced, idempotently |

## Where to next

- [Agent patterns](agent-patterns.md) — how to compose these into a robust LLM
  loop (skim/drill, pin, verify, self-correct, look).
- [Cookbook](cookbook.md) — the same features as random-access recipes, plus
  floating shapes, watermarks, custom list templates, and cross-references.
- [Linting & regularizing](linting.md) — the full rule catalog and profile format.
- [Python API](python-api.md) · [CLI](cli.md) — every signature and flag.
