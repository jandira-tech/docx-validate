## 2026-05-22 - Avoid xpath ancestor:: resolution in @xmldom
**Learning:** Using `xpath` (from the `xpath` NPM package) with `@xmldom/xmldom` is extremely slow when querying with the `ancestor::` axis (e.g., `.//w:p[not(ancestor::w:txbxContent)]`). This causes significant performance bottlenecks for large documents because it traverses the tree for every matched element dynamically instead of just caching parent lookups.
**Action:** When complex ancestor exclusions are needed on large node lists, rely on native DOM APIs (`getElementsByTagNameNS`) combined with a fast `parentNode` while loop in JavaScript. This simple rewrite improved paragraph counting performance by nearly 100x.

## 2026-05-23 - Avoid xpath descendant queries (.//) in @xmldom
**Learning:** Using `xpath` (from the `xpath` NPM package) with `@xmldom/xmldom` to query descendants (e.g., `$$(".//w:del//w:t")`) is extremely slow, especially when resolving nested queries in large DOM trees, as it causes severe performance bottlenecks dynamically searching the tree.
**Action:** Replace `xpath` descendant queries with native DOM traversal methods like `getElementsByTagNameNS()`. When querying specific namespaces, iterate over them (like `WORD_PARAGRAPH_NAMESPACES`) and reuse the namespace for child node lookups. Use a `Set<Node>` to deduplicate matched child elements to ensure the same node-set behavior as `xpath`.
