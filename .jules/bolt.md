## 2026-05-22 - Avoid xpath ancestor:: resolution in @xmldom
**Learning:** Using `xpath` (from the `xpath` NPM package) with `@xmldom/xmldom` is extremely slow when querying with the `ancestor::` axis (e.g., `.//w:p[not(ancestor::w:txbxContent)]`). This causes significant performance bottlenecks for large documents because it traverses the tree for every matched element dynamically instead of just caching parent lookups.
**Action:** When complex ancestor exclusions are needed on large node lists, rely on native DOM APIs (`getElementsByTagNameNS`) combined with a fast `parentNode` while loop in JavaScript. This simple rewrite improved paragraph counting performance by nearly 100x.

## 2026-05-22 - Avoid full-tree traversal for conditional descendant queries
**Learning:** When searching for descendants (like `<w:t>`) that must exist within a specific parent (like `<w:del>`), using `getElementsByTagNameNSAll` on the entire document and checking parents via `parentNode` is inefficient for large documents with many targets but few parents.
**Action:** Query for the specific parents first, then query for the descendants within those parent nodes, using a `Set` to deduplicate overlapping matches due to nested elements.
