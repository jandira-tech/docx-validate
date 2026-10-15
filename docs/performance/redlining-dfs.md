# Redlining traversal measurements

The DFS preserves deletion diagnostics across mixed Word namespaces and excludes drawings/pictures inside deletions. Insertions now emit one issue per offending text node, even when insertion elements nest; the old descendant-query loops emitted duplicates. Text diagnostics follow document order (deletion text remains grouped before instruction text).

Reproduce against two checkouts with Node 24 and installed dependencies:

```sh
pnpm exec tsx scripts/bench-redlining.ts /path/to/parent/src/scripts/office/validators/docx.ts src/scripts/office/validators/docx.ts
```

The synthetic input has one insertion, nested content controls, a deletion, and equally many deleted/inserted text runs. Depth and run count both scale with `size`. Each row is the mean of five runs after one warmup, including file reading, XML parsing, both validators, and diagnostic construction. Both implementations return identical counts on this input.

| Size | Parent | DFS |
| --- | ---: | ---: |
| 200 | 22.97 ms | 17.46 ms |
| 800 | 60.90 ms | 67.71 ms |

These local results are noisy and do **not** establish a consistent end-to-end speedup. Other runs on the same machine favored DFS at both sizes; XML parsing, allocation, and concurrent work affect the total. The implementation removes repeated ancestor walks and descendant scans, but this is an algorithmic change, not evidence for a fixed product-level speedup. Benchmark representative real documents before making stronger performance claims.
