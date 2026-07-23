# Demo GIF choreography

Scripts that drive a live Word window through a repeatable, screen-recordable
story. They exist so the README's demo GIFs can be re-recorded any time the
look of the product changes — run the script, capture the Word window, done.

Both scripts create a **throwaway document** (`Documents.Add`) and never touch
the document you're working in. They leave it open and unsaved; close it
without saving afterwards. Word must already be running.

| script | story | target GIF |
|---|---|---|
| `demo_regularize.py` | messy draft → `lint` findings → `regularize` → **one Ctrl-Z reverts it all** | `docs/assets/regularize.gif` |
| `demo_agent_showcase.py` | an agent builds a styled brief (headings, table) and leaves a review comment | `docs/assets/agent-drive.gif` |

## Recording

1. Put the Word window on a clean desktop, sized to roughly a 4:3 crop, zoom
   ~100–120% so body text is legible at README width (~800 px).
2. Start the recorder on the Word window only ([ScreenToGif](https://www.screentogif.com/)
   or ShareX record straight to GIF; Xbox Game Bar (Win+G) records MP4 you can
   convert). For `demo_regularize.py`, make sure the *end* of the take includes
   you pressing **Ctrl-Z once** — the whole cleanup reverting is the money shot.
3. Run the script from a terminal that is *not* in frame:

   ```
   uv run python examples/demos/demo_regularize.py
   ```

   Each script pauses a few seconds between beats (`PAUSE` at the top) so the
   recording has room to breathe — tune it to taste.
4. Aim for ≤ 10 seconds of idle-free footage and well under 10 MB (GitHub's
   README asset ceiling); 10–15 fps is plenty for UI motion.
5. Drop the result in `docs/assets/` and swap the `<!-- TODO(demo): … -->`
   placeholders at the top of the root `README.md` for real image embeds.
