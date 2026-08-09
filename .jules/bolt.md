## 2026-05-22 - Avoid xpath ancestor:: resolution in @xmldom
**Learning:** Using `xpath` (from the `xpath` NPM package) with `@xmldom/xmldom` is extremely slow when querying with the `ancestor::` axis (e.g., `.//w:p[not(ancestor::w:txbxContent)]`). This causes significant performance bottlenecks for large documents because it traverses the tree for every matched element dynamically instead of just caching parent lookups.
**Action:** When complex ancestor exclusions are needed on large node lists, rely on native DOM APIs (`getElementsByTagNameNS`) combined with a fast `parentNode` while loop in JavaScript. This simple rewrite improved paragraph counting performance by nearly 100x.

## 2026-10-15 - Avoid xpath descendant query in @xmldom
**Learning:** Using `xpath` with descendant queries like `.//` (e.g. `.//w:del//w:t`) causes severe performance bottlenecks for large documents because it dynamically searches the tree.
**Action:** Instead of xpath descendant queries, use native DOM API traversal APIs like `getElementsByTagNameNS` combined with iterating over known namespaces for much better performance.
