# Self-hosted Material Symbols (icons)

The Axis portal self-hosts a **subset** of Material Symbols Outlined instead of
loading the full icon font from Google's CDN. This keeps icons same-origin
(~14 KB) so nothing render-blocks or "distorts" on slow connections, and no
literal ligature text (e.g. "dashboard") can ever flash in.

## Files

- `src/fonts/material-symbols-subset.woff2` — the subset font (only the icons we
  use, FILL axis kept so `Icon filled` works).
- `src/components/ui/material-symbols-codepoints.ts` — generated name→codepoint
  map. `Icon` renders glyphs by codepoint, not ligature name.
- `scripts/subset-icons.py` — regenerates both of the above.

## Adding / changing an icon

1. Add the Material Symbols name to the `USED` (or `BUFFER`) list in
   `scripts/subset-icons.py`.
2. Run it (needs `pip install fonttools brotli`):

   ```bash
   # Download the variable TTF + codepoints once (10 MB, not committed):
   #   https://github.com/google/material-design-icons/tree/master/variablefont
   # Place them next to the script as MaterialSymbols.ttf and codepoints.txt, then:
   python scripts/subset-icons.py
   ```

3. Rebuild. The generated woff2 + TS map land in `src/`.
