## 2026-05-22 - Avoid xpath ancestor:: resolution in @xmldom

**Learning:** Using `xpath` (from the `xpath` NPM package) with `@xmldom/xmldom` is extremely slow when querying with the `ancestor::` axis (e.g., `.//w:p[not(ancestor::w:txbxContent)]`). This causes significant performance bottlenecks for large documents because it traverses the tree for every matched element dynamically instead of just caching parent lookups.
**Action:** When complex ancestor exclusions are needed on large node lists, rely on native DOM APIs (`getElementsByTagNameNS`) combined with a fast `parentNode` while loop in JavaScript. This simple rewrite improved paragraph counting performance by nearly 100x.

## 2026-05-23 - Avoid xpath descendant queries (.//) in @xmldom

**Learning:** Using `xpath` with `@xmldom/xmldom` for descendant queries (like `.//w:t`) is extremely slow, as it dynamically searches the entire tree.
**Action:** When searching for descendants within specific elements, use native DOM traversal APIs like `getElementsByTagNameNS` instead of `xpath` queries to significantly improve performance. Specifically, when iterating over namespaces to find XML elements, reuse the parent's matched namespace URI to query its child elements.
