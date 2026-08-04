## 2026-05-22 - Avoid xpath ancestor:: resolution in @xmldom
**Learning:** Using `xpath` (from the `xpath` NPM package) with `@xmldom/xmldom` is extremely slow when querying with the `ancestor::` axis (e.g., `.//w:p[not(ancestor::w:txbxContent)]`). This causes significant performance bottlenecks for large documents because it traverses the tree for every matched element dynamically instead of just caching parent lookups.
**Action:** When complex ancestor exclusions are needed on large node lists, rely on native DOM APIs (`getElementsByTagNameNS`) combined with a fast `parentNode` while loop in JavaScript. This simple rewrite improved paragraph counting performance by nearly 100x.

## 2026-08-04 - Avoid xpath descendant (//) queries in @xmldom
**Learning:** Using `xpath` (from the `xpath` NPM package) with `@xmldom/xmldom` is extremely slow when querying with descendant axes (e.g., `.//w:del//w:t`) on large documents. This causes significant performance bottlenecks traversing the tree dynamically.
**Action:** When descendant node lookups are needed, rely on native DOM APIs (`getElementsByTagNameNS`) combined with iterating over known namespaces. Deduplicate using a `Set<Node>` to mimic xpath behavior. This simple rewrite improved validation performance significantly.
