"""MCP server for wordlive — drive Word from Claude Desktop and other MCP clients.

The `mcp` SDK is an optional extra (`pip install "wordlive[mcp]"`). It is imported
lazily inside `build_server()` / `main()`, so `import wordlive.mcp` succeeds even
without the extra installed; only actually starting the server needs it.

Run it with the `wordlive-mcp` console script or `python -m wordlive.mcp`.
"""

from __future__ import annotations

from typing import TYPE_CHECKING

if TYPE_CHECKING:
    from ._worker import Worker

__all__ = ["build_server", "main"]


def build_server(worker: Worker | None = None):  # noqa: ANN201 — return type needs the extra
    """Construct the FastMCP server. See `wordlive.mcp.server.build_server`."""
    from .server import build_server as _build

    return _build(worker)


def main() -> None:
    """Entry point for `wordlive-mcp` / `python -m wordlive.mcp` (stdio transport)."""
    from .server import main as _main

    _main()
