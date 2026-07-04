## 2026-05-22 - Avoid xpath ancestor:: resolution in @xmldom
**Learning:** Using `xpath` (from the `xpath` NPM package) with `@xmldom/xmldom` is extremely slow when querying with the `ancestor::` axis (e.g., `.//w:p[not(ancestor::w:txbxContent)]`). This causes significant performance bottlenecks for large documents because it traverses the tree for every matched element dynamically instead of just caching parent lookups.
**Action:** When complex ancestor exclusions are needed on large node lists, rely on native DOM APIs (`getElementsByTagNameNS`) combined with a fast `parentNode` while loop in JavaScript. This simple rewrite improved paragraph counting performance by nearly 100x.

## 2025-02-12 - Replaced xpath query in docx.ts
**Learning:** `xpath`'s descendant queries (`.//`) dynamically search the tree and are very slow compared to the native DOM traversal API `getElementsByTagNameNS`.
**Action:** Always replace `xpath` descendant queries with `getElementsByTagNameNS` iterating over known namespaces when modifying elements or validating XML structure. Use a `Set<Element>` to deduplicate results when nested tags might lead to `getElementsByTagNameNS` matching the inner element multiple times.
