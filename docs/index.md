---
title: wordlive
hide:
  - navigation
---

# wordlive

**Drive a running Microsoft Word instance from Python — `xlwings`, but for Word.**

Built for both human scripting and LLM agents. Windows-only.

<div class="grid cards" markdown>

-   :material-rocket-launch:{ .lg .middle } **New here?**

    ---

    Install, attach to Word, and run your first polite edit in five minutes.

    [:octicons-arrow-right-24: Getting started](getting-started.md)

-   :material-lightbulb-on:{ .lg .middle } **How it thinks**

    ---

    Anchors, atomic undo, politeness — the four ideas that drive the API.

    [:octicons-arrow-right-24: Concepts](concepts.md)

-   :material-console:{ .lg .middle } **CLI**

    ---

    JSON-in / JSON-out commands designed to drop into an LLM tool-use loop.

    [:octicons-arrow-right-24: CLI reference](cli.md)

-   :material-code-braces:{ .lg .middle } **Python API**

    ---

    Every public class and function, generated from source docstrings.

    [:octicons-arrow-right-24: Python API](python-api.md)

-   :material-broom:{ .lg .middle } **Clean up formatting**

    ---

    Audit a document and normalize its formatting in one atomic-undo pass.

    [:octicons-arrow-right-24: Linting & regularizing](linting.md)

</div>

---

{%
   include-markdown "../README.md"
   start="## Install"
   end="## Design"
%}

## Design principles

- **Politeness first** — operations preserve the user's `Selection`, view, and
  scroll. The user keeps editing alongside you.
- **Semantic anchors over `Selection`** — operations target bookmarks, content
  controls, or headings — never the live cursor unless you ask.
- **Atomic undo** — every `doc.edit()` opens a Word `UndoRecord`, so a single
  Ctrl-Z reverts the whole block.
- **Escape hatch** — every wrapper exposes `.com` for the raw COM object;
  you're never blocked by missing coverage.

See the [Design](design.md) page for the full rationale.
