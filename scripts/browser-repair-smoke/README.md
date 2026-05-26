# browser-repair-smoke

Proof that **`repairDocxInMemory` runs in a real browser** — same in-memory path
as Node, no native `libxmljs2`, no Node `Buffer`.

## What it proves

The repair logic parses/serializes with `@xmldom/xmldom` (pure JS) and operates
on a `MemoryPartFS` (Uint8Array + TextEncoder/TextDecoder, no `Buffer`). The
native `libxmljs2` is only loaded lazily by the **XSD validation** path (never
the repair path), so a browser bundle that only repairs never touches it.

## Files

- `entry.mjs` — imports `repairDocxInMemory` from the built `dist/`, unpacks a
  DOCX with JSZip, repairs the parts in memory, exposes `globalThis.runRepair`.
- `stub.mjs` — browser stubs for the Node builtins the repair path never calls
  (`node:fs`/`module`/`url`/`os`/…, plus the `tmp` package). `node:path` is
  polyfilled with `./scripts/browser-repair-smoke/mini-path.mjs`.
- `index.html` — loads the bundle as an ES module.
- `runner.mjs` — Playwright: serves the page, calls `runRepair` on a real DOCX,
  asserts the result.

## Run

```bash
npm i -D esbuild jszip playwright && npx playwright install chromium  # node:path → ./mini-path.mjs (no polyfill pkg)
# build dist first: npm run build
npx esbuild scripts/browser-repair-smoke/entry.mjs --bundle --format=esm --platform=browser \
  --banner:js="globalThis.process=globalThis.process||{env:{},argv:[],platform:'browser',cwd:function(){return '/'}};" \
  --alias:tmp=./scripts/browser-repair-smoke/stub.mjs \
  --alias:node:path=./scripts/browser-repair-smoke/mini-path.mjs --alias:path=./scripts/browser-repair-smoke/mini-path.mjs \
  --alias:node:fs=./scripts/browser-repair-smoke/stub.mjs --alias:fs=./scripts/browser-repair-smoke/stub.mjs \
  --alias:node:module=./scripts/browser-repair-smoke/stub.mjs --alias:node:url=./scripts/browser-repair-smoke/stub.mjs \
  --alias:node:child_process=./scripts/browser-repair-smoke/stub.mjs --alias:child_process=./scripts/browser-repair-smoke/stub.mjs \
  --alias:node:process=./scripts/browser-repair-smoke/stub.mjs --alias:process=./scripts/browser-repair-smoke/stub.mjs \
  --alias:node:os=./scripts/browser-repair-smoke/stub.mjs --alias:os=./scripts/browser-repair-smoke/stub.mjs \
  --alias:node:crypto=./scripts/browser-repair-smoke/stub.mjs --alias:crypto=./scripts/browser-repair-smoke/stub.mjs \
  --alias:node:util=./scripts/browser-repair-smoke/stub.mjs --alias:util=./scripts/browser-repair-smoke/stub.mjs \
  --alias:node:stream=./scripts/browser-repair-smoke/stub.mjs --alias:stream=./scripts/browser-repair-smoke/stub.mjs \
  --alias:node:events=./scripts/browser-repair-smoke/stub.mjs --alias:events=./scripts/browser-repair-smoke/stub.mjs \
  --alias:node:assert=./scripts/browser-repair-smoke/stub.mjs --alias:assert=./scripts/browser-repair-smoke/stub.mjs \
  --alias:node:buffer=./scripts/browser-repair-smoke/stub.mjs --alias:buffer=./scripts/browser-repair-smoke/stub.mjs \
  --alias:constants=./scripts/browser-repair-smoke/stub.mjs \
  --outfile=/tmp/browser-repair/bundle.js
# then: serve /tmp/browser-repair, run runner.mjs
```

## Proven result (headless Chromium)

Repairing `nested-comments-marker.docx` (source has whitespace `<w:t>` missing
`xml:space`) in the browser:

```
{ ok:true, repairs:22, partCount:14, xmlSpaceCount:3, bufferDefined:false, processDefined:true }
```

- `repairs:22` and `xmlSpaceCount` 0→3 — the repair ran and worked.
- **`bufferDefined:false`** — Node's `Buffer` is absent ⇒ genuinely the browser.
- Identical repair count to the Node path ⇒ "identical in Node and browser".

Note: `node:path` is polyfilled (`./scripts/browser-repair-smoke/mini-path.mjs`); the bundler (vite/esbuild)
supplies it. The native `libxmljs2` (XSD validation) is **not** browser-capable —
swapping it for a WASM libxml2 (`libxml2-wasm`) would make validation run in the
browser too; that is a separate enhancement.
