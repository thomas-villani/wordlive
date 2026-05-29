"""MCPB launcher for the wordlive MCP server (uv runtime).

Thin wrapper over wordlive's ``wordlive-mcp`` entry point
(``wordlive.mcp.__main__:main``). It exists only to give the MCPB ``uv`` runtime
a stable entry file; all the real logic lives in the installed ``wordlive``
package. The server needs no configuration — it drives whatever Word document
the user has open — so there are no ``user_config`` env vars to read.
"""

from wordlive.mcp.__main__ import main

if __name__ == "__main__":
    raise SystemExit(main())
