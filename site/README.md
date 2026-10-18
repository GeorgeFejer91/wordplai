# wordplai website

The site has two ordinary, permanent page routes:

- `/wordplai/` — Description
- `/wordplai/download/` — Download

Run `npm ci`, `npm run site:build`, then `npm run test:site` from the repository root.
The build writes `site/dist/`; tests write screenshots, measured label geometry and a
JSON result to `site/test-results/`. Both directories are ignored. Tests start a
local HTTP server and a new headless Chrome instance, with no existing browser
profile. Set `CHROME_PATH` if Chrome is installed outside the usual locations.

`release.json` is the release metadata source. It contains exactly `version`,
`asset`, `downloadUrl`, `sha256`, `sizeBytes`, and `publishedAt`. The URL must name
this repository and an explicit release asset. The SHA-256, exact positive byte
count, and publication time must all be supplied before a download link is shown.
Missing metadata renders a readable unavailable state; malformed metadata fails
the build. The preview URL never uses `releases/latest`.

The small browser script bundles locked Pretext 0.0.8 to measure navigation and
action labels against their CSS boxes. The same locally served regular/bold font
faces are used in the DOM and canvas measurements. Long labels reflow; enlarged
text is not shrunk. The page remains readable without scripts or fonts, and
unsupported text spacing uses the browser's ordinary layout. Body prose stays in
normal document flow. Asset provenance and licenses are in `assets/`.

The rendered tests cover both routes at desktop, 320px, landscape, doubled text,
long translated labels and expanded text spacing, plus keyboard navigation,
missing-font and no-script fallbacks. They validate site behavior, not the Word
add-in or the real ChatGPT connection.
