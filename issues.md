# wordlive — code review findings

Working backlog from a comprehensive review of v0.3 (commit `c03b7d1`). Each
item is self-contained so it can be picked up across sessions without
re-deriving context.

Severity:

- 🔴 **High** — observable bug, user-visible inconsistency, or breaks a
  documented contract.
- 🟡 **Medium** — edge-case bug, design wart, performance concern.
- 🟢 **Low** — polish, untested edge case, doc rot.

Status: **all open** as of 2026-05-19. Strike through and tag `[fixed]` when
closing.

---

## Bugs / correctness

### ~~🔴 B-1. `kind` field inconsistency: same content control reads as `"content_control"` from `write cc` and `"content control"` from every other CLI command~~ [fixed in-tree]

An LLM-driven loop that branches on the `kind` field will see two different
strings for the *same* content control depending on which verb produced the
response:

- `wordlive write cc Signatory --text X` → `{"anchor": {"kind": "content_control", ...}}` (underscore, hardcoded at `src/wordlive/cli/commands.py:226`)
- `wordlive replace --anchor-id cc:Signatory --text X` → `{"anchor": {"kind": "content control", ...}}` (space, from `Anchor.kind` at `src/wordlive/_anchors.py:223`)
- Same for `go-to`, `style apply`, `format-paragraph` (all use `anchor.kind`).

Bookmark and heading commands happen to be self-consistent only because
`Bookmark.kind == "bookmark"` and `Heading.kind == "heading"` already match
their anchor-id prefixes.

**Fix shape:** pick one canonical form and use it everywhere. Recommend
`"content_control"` (snake_case, matches JSON conventions) — change
`ContentControl.kind` in `_anchors.py` and let every command pull from
`anchor.kind`. Drop the hardcoded literal in `write_cc`. Add a CLI test that
asserts the same anchor produces the same `kind` regardless of verb (T-4
below).

### ~~🟡 B-2. `Bookmark.set_text` and `Heading.insert_paragraph_after` use Python `len(text)` to compute Word offsets — wrong for surrogate-pair characters~~ [fixed in-tree]

`src/wordlive/_anchors.py:177-179` and `:422` both do
`doc_com.Range(start, start + len(text))` after writing `text` into a
range. Word's COM API counts characters in UTF-16 code units; Python's
`len()` counts code points. For BMP-only text these match. For text
containing emoji or other astral-plane chars (U+10000+), each Python char
maps to two UTF-16 code units, so the computed `Range` overshoots.

Symptoms:

- `Bookmark.set_text("🎉 done")`: new bookmark is re-added covering the wrong
  span (too short) — subsequent reads of the bookmark may include unrelated
  trailing characters.
- `Heading.insert_paragraph_after("🎉 ...", style="Body Text")`: the
  style-application range is shorter than intended; some of the new
  paragraph keeps the previous style.

**Fix shape:** after the assignment, read back the actual range bounds from
Word (`bm.Range.End` after re-querying, or `insert_rng.End` after the text
assignment, since Word updates the range to span the inserted text). Or
compute the new end via `start + len(text.encode("utf-16-le")) // 2`.
Untested in CI because no fixture exercises non-BMP input — worth adding.

### ~~🟡 B-3. `_cc_by_name("")` matches the first content control whose Title and Tag are both empty~~ [fixed in-tree]

`src/wordlive/_anchors.py:217`:
`if str(cc.Title or "") == name or str(cc.Tag or "") == name`.

When `name == ""`, this matches any CC where both `Title` and `Tag` are
falsy. Content controls without titles or tags are rare but legal in Word
templates. An LLM that emits an empty `cc:` anchor id (e.g. from a stripped
payload) would silently hit the wrong CC.

**Fix shape:** reject empty `name` early in
`ContentControlCollection.__getitem__` and `_cc_by_name` — raise
`AnchorNotFoundError("content control", "")`.

### ~~🟡 B-4. Malformed `exec` script ops raise raw `KeyError`, escape the typed-exception path, and produce a stack trace + exit 1~~ [fixed in-tree]

`src/wordlive/cli/commands.py:519-553` (`_apply_op`) indexes `op["name"]`,
`op["text"]`, etc. without validation. An `exec` script with a typo'd field
(e.g. `{"op": "write_bookmark", "nme": "Address", "text": "..."}`) raises
`KeyError("name")`. `_run` only catches `WordliveError`, so the Python
traceback prints to stderr and the exit code becomes 1 (Click's default for
uncaught exceptions inside `runner.invoke(catch_exceptions=True)`).

Two issues:

1. Exit code is 1 ("other") instead of an obvious validation failure code.
2. Stderr emits a Python traceback — hostile to LLM tool-use loops that
   parse the failure mode.

**Fix shape:** validate the op payload in `_apply_op` before dispatch.
Either raise `click.ClickException("write_bookmark requires 'name' and 'text'")`
(non-zero exit, clean stderr line, no traceback) or define a new typed
error like `BadOpScriptError(WordliveError)` and add it to `_exit_for`.
Test in `tests/test_cli.py` alongside the existing `test_exec_unknown_op_is_click_error`.

### ~~🟢 B-5. `BookmarkCollection.list()` returns Word's hidden internal bookmarks (`_Toc...`, `_Ref...`) alongside user-created ones~~ [fixed in-tree]

`src/wordlive/_anchors.py:200-202` iterates every bookmark in the doc,
including the ones Word auto-creates for table-of-contents entries and
cross-references (names beginning with `_`). For documents with a TOC or
many cross-references this drowns the user's actual bookmarks in noise.

**Fix shape:** filter out names starting with `_` in `list()` and
`__iter__`. Keep `__getitem__` and `Exists` working on hidden names
(escape hatch for advanced users). Add a `include_hidden=False` kwarg if
you want to keep them accessible.

### ~~🟢 B-6. `Style` property accesses don't cache — every `style.type`/`style.builtin`/`style.in_use` walks the entire `Styles` collection~~ [fixed in-tree]

`src/wordlive/_styles.py:53-71`. Each property goes through `self.com`,
which iterates `self._doc.com.Styles` and matches by `NameLocal`. For a
document with 300 styles that's 300 COM dispatches per attribute read.
`StyleCollection.list()` is fine (one pass), and `style.to_dict()` only
hits COM once. Code that reads multiple properties off a single Style is
where this bites:

```python
s = doc.styles["Heading 1"]
print(s.type, s.builtin, s.in_use)   # three full Styles scans
```

The module docstring justifies the no-cache choice by saying renames /
deletions mid-session would return stale data. That tradeoff is worth
keeping — but we can still avoid re-scanning *within* a single property
read by trying `doc.com.Styles(name)` first (direct lookup) and falling
back to iteration only if that raises `pywintypes.com_error`. Direct
lookup is what Word's COM does internally; we only iterate for the
membership-check path because Word doesn't reserve an HRESULT for
"missing style." Once we've already established the style exists, direct
lookup is safe.

**Fix shape:** in `Style.com`, try `self._doc.com.Styles(self._name)` first
inside `translate_com_errors`; fall back to iteration if it raises. Keep
`StyleCollection.__contains__` and `__getitem__` doing the membership
walk first (so missing styles still get the typed error).

---

## Design inconsistencies

### ~~🟡 D-1. `Document.heading()` is a method but `bookmarks` / `content_controls` are properties returning collections — asymmetric API~~ [fixed in-tree]

`Document.heading(name)` returns a `Heading` directly. `Document.bookmarks`
returns a `BookmarkCollection` and you index into it. There's no
`Document.headings` collection equivalent to `bookmarks` / `styles`.

Consequence: a user iterating over all anchors of a kind can do
`for bm in doc.bookmarks:` but must call `for item in doc.outline():`
then resolve via `anchor_by_id` to enumerate headings.

**Fix shape:** add `Document.headings` returning a `HeadingCollection` with
`__iter__`, `__contains__`, `__getitem__(name | int)`, `list()`. Keep
`Document.heading(name)` as a sugar for `doc.headings[name]`. Document
both. Non-breaking: existing code paths still work.

### ~~🟡 D-2. `Anchor` base class isn't an ABC; `_range()` raises `NotImplementedError` at call time, not at construction~~ [fixed in-tree]

`src/wordlive/_anchors.py:50-67`. `Anchor` is documented as abstract but
isn't `abc.ABC`. `Anchor(doc, "name")` succeeds; only calling `.text` /
`.set_text` etc. surfaces the missing implementation. A future contributor
adding a fourth anchor type and forgetting to override `_range()` gets a
silent-until-runtime failure.

**Fix shape:** make `Anchor(abc.ABC)`, mark `_range` and `set_text` with
`@abstractmethod`. Existing subclasses already provide both.

### ~~🟢 D-3. `Anchor.anchor_id` base default `f"{kind}:{name}"` doesn't match the actual scheme used by content controls~~ [fixed in-tree]

The base returns `"{kind}:{name}"`, which for a `ContentControl` (where
`kind == "content control"`) would yield `"content control:Signatory"` —
but the subclass overrides to return `"cc:Signatory"`. The base default is
correct only for the kinds whose `kind` happens to match their anchor-id
scheme. A future contributor relying on the inherited default would emit
the wrong scheme.

**Fix shape:** drop the base default (raise `NotImplementedError`), or
introduce a class attribute `anchor_id_scheme` (e.g. `"bookmark"`,
`"cc"`, `"heading"`) and have the base use that. Either way, force
subclasses to opt in.

### ~~🟢 D-4. `Document.go_to` is not undoable and doesn't open an edit scope — but neither is documented as such~~ [fixed in-tree]

`src/wordlive/_document.py:238-248`. Moving the cursor is fast and not
something users expect on the undo stack, so this is reasonable. But the
docstring just says "Move the user's Selection" without making the
contract explicit. A reader comparing this to `set_text`-style methods
(which all wrap in `translate_com_errors` and *are* polite-by-default)
might wonder where the politeness boundary is.

**Fix shape:** add one line to the docstring: "No `UndoRecord` is opened —
cursor moves don't belong on the user's undo stack."

---

## Performance

(Most of the perf concerns are folded into B-6 above. No standalone
items at the moment — the surface is small and live-COM dispatch
dominates anyway.)

---

## Testing gaps

### ~~🟡 T-1. `apply_style` / `format_paragraph` are only exercised on `Bookmark` in unit tests; never on `ContentControl` or `Heading`~~ [fixed in-tree]

`tests/test_styles.py` applies `apply_style("Heading 2")` and
`format_paragraph(alignment=...)` exclusively to
`doc.bookmarks["Address"]`. The E2E script covers Bookmark
(`t_apply_style_to_bookmark`) and Heading (`t_format_paragraph`) but
neither exercises ContentControl. Since the methods live on the base
`Anchor`, behavior *should* be identical — but a regression in `_range()`
on ContentControl would go unnoticed.

**Fix shape:** add three unit tests:
- `test_apply_style_on_content_control_writes_through(fake_word)`
- `test_apply_style_on_heading_writes_through(fake_word)`
- `test_format_paragraph_on_heading_sets_alignment(fake_word)`

Each can mirror the existing bookmark cases.

### ~~🟡 T-2. No test verifies the `kind` field is consistent across CLI commands for the same anchor (would catch B-1)~~ [fixed in-tree]

A single parametrized test that drives `write_cc`, `replace --anchor-id`,
`go-to`, `style apply`, and `format-paragraph` against the same content
control and asserts `data["anchor"]["kind"]` matches across all five
responses would have caught the underscore-vs-space discrepancy before
v0.3.

**Fix shape:** add to `tests/test_cli.py`:

```python
def test_kind_field_is_consistent_across_cc_commands(fake_word):
    cmds = [
        ["write", "cc", "Signatory", "--text", "X"],
        ["replace", "--anchor-id", "cc:Signatory", "--text", "X"],
        ["go-to", "--anchor-id", "cc:Signatory"],
        ["style", "apply", "--anchor-id", "cc:Signatory", "--name", "Heading 2"],
        ["format-paragraph", "--anchor-id", "cc:Signatory", "--alignment", "left"],
    ]
    kinds = {tuple(c): json.loads(_invoke(c)[1])["anchor"]["kind"] for c in cmds}
    assert len(set(kinds.values())) == 1, kinds
```

### ~~🟢 T-3. `_section_range` is untested for the case where the target heading is the very last paragraph in the document~~ [fixed in-tree]

`tests/test_anchors.py:209-245` covers "section runs to end of doc" and
"section stops at next same-level heading" but not "target paragraph is
the last paragraph". The code in `src/wordlive/_anchors.py:320-322`
returns `Range(end, end)` (zero-width) in this case — probably correct
but unverified.

**Fix shape:** add a fixture with a heading as the last paragraph and
assert `section_range().Start == section_range().End`.

### ~~🟢 T-4. `from_com_error` is tested for only two HRESULTs; `_BUSY_HRESULTS` has five entries~~ [fixed in-tree]

`tests/test_anchors.py:66-99` covers `0x80010001` (busy) and `-2147352567`
(generic). The four other busy HRESULTs in `exceptions.py:82-90`
(`0x8001010A`, `0x80010005`, and the negative-signed forms) are
unverified. A parametrized test would close this in two lines.

**Fix shape:** parametrize `test_from_com_error_classifies_busy` over
every entry in `_BUSY_HRESULTS`.

### ~~🟢 T-5. Malformed `exec` op payloads (missing required field, wrong type) — no test (relates to B-4)~~ [fixed in-tree]

Once B-4 is fixed, add coverage for:

- `{"op": "write_bookmark", "text": "..."}` (missing `name`)
- `{"op": "find_replace", "find": "..."}` (missing `text`)
- `{"op": "apply_style", "anchor_id": "bookmark:Address"}` (missing `name`)

Each should produce a deterministic non-zero exit code and a structured
failure on stdout (or stderr) — not a Python traceback.

---

## Documentation issues

### ~~🟡 Doc-1. `feature-plan.md` v0.3 section is labeled "collaboration features" but v0.3 actually shipped styles~~ [fixed in-tree]

`feature-plan.md:90` reads `## v0.3 — collaboration features` (comments,
track changes, find/replace, RangeAnchor). But v0.3 shipped as
**styles + paragraph formatting** (per commit `c03b7d1`). The styles
bullets under the v0.2 header were correctly marked "shipped in v0.3"
but the v0.3 section header itself wasn't renumbered.

Also: find/replace was already shipped in v0.2 (commits `f90e0a9`,
`cbe89cc`) but is listed as a v0.3 bullet.

**Fix shape:** renumber `## v0.3 — collaboration features` → `## v0.4 —
collaboration features`, drop the "Find / replace" bullet (shipped), and
rebase the rest. Move "RangeAnchor" out of collaboration into its own
section if appropriate.

### ~~🟡 Doc-2. `pyproject.toml` still pinned at `version = "0.1.0"` despite v0.2 and v0.3 releases~~ [fixed in-tree]

`pyproject.toml:3`. If you publish to PyPI, the wheel will be `0.1.0`
regardless of what the v0.3 features added. Even if PyPI is not in scope
today, version drift between code and metadata is a future headache.

**Fix shape:** bump to `0.3.0` now and adopt a "bump on each merged
release" workflow. Tag the commit if you want a Git milestone.

### ~~🟡 Doc-3. `README.md` doesn't mention `style list` / `style apply` / `format-paragraph` CLI commands or the new exec ops~~ [fixed in-tree]

`README.md:38-50` lists the CLI verbs as of v0.1: `status`, `outline`,
`read`, `write`, `insert`, `replace`, `go-to`, `exec`. The styles +
paragraph-formatting surface is missing. A first-time reader hitting the
README won't know those commands exist.

**Fix shape:** add a short stanza after the existing `wordlive go-to`
line, e.g.:

```
# Restyle and format paragraphs (atomic-undo):
wordlive style list
wordlive style apply --anchor-id heading:3 --name "Heading 2"
wordlive format-paragraph --anchor-id heading:3 --alignment center --space-before 6
```

And add `apply_style` / `format_paragraph` to the example `ops.json`.

### ~~🟢 Doc-4. `getting-started.md` exit-code table omits code 5 (ambiguous match)~~ [fixed in-tree]

`docs/getting-started.md:99-105` lists exit codes 0, 2, 3, 4, 1 — but not
5 (`AmbiguousMatchError`, used by `replace --find` with multiple hits).
The full table in `cli.md` and `errors.md` is correct; only the
quickstart abridgement is missing it.

**Fix shape:** add a row for `5` to the table; one line.

### ~~🟢 Doc-5. `design.md:103-105` reads "At time of writing v0.1 covers …" — stale~~ [fixed in-tree]

By the time a reader gets to design.md, v0.3 has shipped. The doc still
positions v0.1 as the high-water mark.

**Fix shape:** either bump the version reference per release or rephrase
to "the current release covers …" without naming a version, so it stops
rotting.

### ~~🟢 Doc-6. `concepts.md` references `src/wordlive/_document.py:62` for `anchor_by_id` — actual location is line 71~~ [fixed in-tree]

`docs/concepts.md:122`. Line numbers drift; this one drifted after the
`styles` property was added at lines 60-61 of `_document.py`. The mkdocs
build doesn't catch stale source-line links.

**Fix shape:** drop the line numbers from inline links (`_document.py`
without `:N`) — the function names in the URL fragment do the work. Same
treatment for the other two implementation sidebars on the page if you
want to be consistent.

### ~~🟢 Doc-7. `cookbook.md §3a` shows `bookmark.insert_after(...)` without noting the bookmark range is *not* extended~~ [fixed in-tree]

`docs/cookbook.md:158-159`. `Bookmark.insert_after` inserts text *outside*
the bookmark's range. The bookmark itself still covers the original
text. Contrast with `Bookmark.set_text` (which re-adds the bookmark
covering the new content). A reader could reasonably expect "insert
after" to extend the bookmark.

**Fix shape:** one sentence after the example: "Note: `insert_before` /
`insert_after` leave the bookmark's range unchanged — only `set_text`
re-adds the bookmark to cover its new content."

### ~~🟢 Doc-8. CLI `--alignment` accepts `centre` (UK spelling) but the docs only list `center`~~ [fixed in-tree]

`src/wordlive/cli/commands.py:457` declares
`click.Choice(["left", "center", "centre", "right", "justify"], case_sensitive=False)`.
`docs/cli.md:343-344` only documents the US spelling. Harmless
side-by-side; either drop `centre` for parsimony or document it.

### ~~🟢 Doc-9. `commands.py:567` exec-command docstring lists only the v0.1 ops~~ [fixed in-tree]

```python
"""Apply a batch of ops in a single atomic-undo scope.

…
Supported ops: write_bookmark, write_cc, insert_after_heading, replace,
find_replace.
"""
```

`apply_style` and `format_paragraph` are missing. `wordlive exec --help`
shows the stale list.

**Fix shape:** append the two ops, or rewrite to "see docs/cli.md for the
full ops list" to avoid future rot.

---

## Polish

### ~~🟢 M-1. `_paragraph_text` is private (leading underscore) but imported across modules~~ [fixed in-tree]

`src/wordlive/_anchors.py:284` defines `_paragraph_text`;
`src/wordlive/_document.py:14` imports it. By convention, names with
leading underscores aren't intended for use outside their defining
module.

**Fix shape:** rename to `paragraph_text` (drop the prefix), or move into
a new internal helpers module (e.g. `_text.py`).

### ~~🟢 M-2. `WdStyleType` constants are UPPERCASE but `_style_type_name` returns lowercase strings — mapping isn't documented~~ [fixed in-tree]

`constants.py:44-48`: `WdStyleType.PARAGRAPH = 1`. The CLI / JSON output
emits `"type": "paragraph"`. A user who imports `WdStyleType` to
construct filters expects to compare against the upper-cased enum name.

**Fix shape:** add one line to the `WdStyleType` docstring noting that
the public JSON form is lowercase (`"paragraph"`, etc.). Or expose
`WdStyleType.PARAGRAPH.json_name == "paragraph"` as a property if you
want machine-readable mapping.

### ~~🟢 M-3. `_FakeStyles.__call__` in the test fake is dead code — wordlive never calls `doc.Styles(name)` directly~~ [resolved by B-6: direct-lookup path now makes this load-bearing]

`tests/conftest.py:132-136` implements `__call__` raising `KeyError` for
direct-lookup. The library iterates `Styles` everywhere (see B-6 above).
If B-6's fix lands (try direct lookup first), this fake call path
suddenly becomes load-bearing — good. If B-6 doesn't land, the fake's
`__call__` is dead and could be deleted.

**Fix shape:** keep until B-6 is decided. If B-6 lands, add an
assertion in tests that this path is exercised at least once.

---

## How to use this file

- Each item is independently fixable; pick one per session if you want
  to ratchet quality without scope creep.
- 🔴 items first if you're prioritising; 🟢 items are good "fill-in"
  tasks between feature work.
- Before closing an item, write a one-line test that would have caught
  it (where applicable), so the regression can't sneak back.
- When closing, strike the heading and append ` [fixed in <commit>]`
  rather than deleting — the file is a record, not a TODO.
