## 2026-05-22 - Avoid xpath ancestor:: resolution in @xmldom
**Learning:** Using `xpath` (from the `xpath` NPM package) with `@xmldom/xmldom` is extremely slow when querying with the `ancestor::` axis (e.g., `.//w:p[not(ancestor::w:txbxContent)]`). This causes significant performance bottlenecks for large documents because it traverses the tree for every matched element dynamically instead of just caching parent lookups.
**Action:** When complex ancestor exclusions are needed on large node lists, rely on native DOM APIs (`getElementsByTagNameNS`) combined with a fast `parentNode` while loop in JavaScript. This simple rewrite improved paragraph counting performance by nearly 100x.

## 2026-05-31 - Replace xpath descendants queries with native DOM
**Learning:** XPath descendant queries (e.g. `.//w:del//w:t`) using `xpath` against `@xmldom/xmldom` documents are very slow on large node trees.
**Action:** Replace XPath descendants queries (like `.//`) with native DOM methods like `getElementsByTagNameNS`. When targeting elements within specific parent nodes, fetch the parent nodes first, and reuse the namespace URI to query descendants for optimal performance.
