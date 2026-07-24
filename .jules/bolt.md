## 2026-05-22 - Avoid xpath ancestor:: resolution in @xmldom
**Learning:** Using `xpath` (from the `xpath` NPM package) with `@xmldom/xmldom` is extremely slow when querying with the `ancestor::` axis (e.g., `.//w:p[not(ancestor::w:txbxContent)]`). This causes significant performance bottlenecks for large documents because it traverses the tree for every matched element dynamically instead of just caching parent lookups.
**Action:** When complex ancestor exclusions are needed on large node lists, rely on native DOM APIs (`getElementsByTagNameNS`) combined with a fast `parentNode` while loop in JavaScript. This simple rewrite improved paragraph counting performance by nearly 100x.

## 2026-05-23 - Avoid xpath descendant (//) resolution in @xmldom
**Learning:** Using `xpath` with `@xmldom/xmldom` to resolve descendant axes (like `.//w:del//w:t`) is extremely slow for large documents, as it dynamically evaluates and filters the entire subtree.
**Action:** When searching for descendants, especially nested ones, replace `xpath` strings with native `getElementsByTagNameNS` loops and use a `Set<Node>` to deduplicate any identically-matched nodes in nested structures.
