## 2026-05-22 - Avoid xpath ancestor:: resolution in @xmldom
**Learning:** Using `xpath` (from the `xpath` NPM package) with `@xmldom/xmldom` is extremely slow when querying with the `ancestor::` axis (e.g., `.//w:p[not(ancestor::w:txbxContent)]`). This causes significant performance bottlenecks for large documents because it traverses the tree for every matched element dynamically instead of just caching parent lookups.
**Action:** When complex ancestor exclusions are needed on large node lists, rely on native DOM APIs (`getElementsByTagNameNS`) combined with a fast `parentNode` while loop in JavaScript. This simple rewrite improved paragraph counting performance by nearly 100x.

## 2026-05-22 - Avoid xpath descendant queries (.//) in @xmldom
**Learning:** Using `xpath` with descendant queries like `.//w:del//w:t` on large XML trees using `@xmldom/xmldom` is heavily unoptimized. Processing 1000 `<w:del>` blocks took ~870ms with `xpath` compared to ~18ms with native `getElementsByTagNameNS` and a deduplicating `Set`. The `xpath` implementation traverses the tree dynamically for every query.
**Action:** Always prefer native DOM APIs like `getElementsByTagNameNS` coupled with a `Set<Element>` (to handle nested parent matching anomalies) over `xpath` descendant queries when performing broad node discoveries on large OOXML documents.
