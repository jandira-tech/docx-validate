## 2026-05-22 - Avoid xpath ancestor:: resolution in @xmldom
**Learning:** Using `xpath` (from the `xpath` NPM package) with `@xmldom/xmldom` is extremely slow when querying with the `ancestor::` axis (e.g., `.//w:p[not(ancestor::w:txbxContent)]`). This causes significant performance bottlenecks for large documents because it traverses the tree for every matched element dynamically instead of just caching parent lookups.
**Action:** When complex ancestor exclusions are needed on large node lists, rely on native DOM APIs (`getElementsByTagNameNS`) combined with a fast `parentNode` while loop in JavaScript. This simple rewrite improved paragraph counting performance by nearly 100x.

## 2026-05-23 - Optimize XPath descendant traversal for deleted elements
**Learning:** Searching the entire document for specific descendant nodes (like w:t) before checking if they have a specific ancestor (w:del) is extremely slow when the target ancestor is rare. Iterating through all nodes to check parentNode scales poorly.
**Action:** Optimize DOM traversal by querying for the rare ancestor first (e.g. w:del), then searching for the target descendants only within those matched ancestors. Use a Set to deduplicate matches due to potential nested structures.
