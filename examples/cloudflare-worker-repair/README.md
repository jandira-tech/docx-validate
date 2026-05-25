# Cloudflare Worker DOCX repair demo

Proves docx-validate's in-memory repair runs in a **no-Node, no-native**
runtime (Cloudflare Workers / workerd — same constraints as a browser).

`worker-entry.mjs` imports `repairDocxInMemory`, accepts a POSTed .docx, unpacks
with JSZip, repairs in memory, repacks. Pre-bundle with esbuild (stub the Node
builtins the repair path never calls, `node:path` → path-browserify, exclude the
native `libxmljs2`; see `scripts/browser-repair-smoke/README.md` for the flags),
then `wrangler dev --no-bundle`.

## Verified result (workerd + a real browser uploading 10 broken .docx)

10/10 processed; `buffer=false` in every response (Node Buffer absent ⇒ workerd).
Re-validated in Node: **7/10 broken → valid**. The 3 not fixed
(unmatched comment marker, missing rels sidecar, orphan part) are not
auto-repairable.

Repair is browser/worker-capable today. Full **XSD validation** in this runtime
needs the native `libxmljs2` swapped for a WASM libxml2 (`libxml2-wasm`,
benchmarked best-fit) — tracked separately.
