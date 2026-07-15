## 2026-05-22 - Avoid xpath ancestor:: resolution in @xmldom
**Learning:** Using `xpath` (from the `xpath` NPM package) with `@xmldom/xmldom` is extremely slow when querying with the `ancestor::` axis (e.g., `.//w:p[not(ancestor::w:txbxContent)]`). This causes significant performance bottlenecks for large documents because it traverses the tree for every matched element dynamically instead of just caching parent lookups.
**Action:** When complex ancestor exclusions are needed on large node lists, rely on native DOM APIs (`getElementsByTagNameNS`) combined with a fast `parentNode` while loop in JavaScript. This simple rewrite improved paragraph counting performance by nearly 100x.

## 2026-07-15 - Avoid nesting namespace iteration loops when retrieving child elements
**Learning:** When retrieving namespaced XML elements using `getElementsByTagNameNS`, iterating over potential namespaces (like `WORD_PARAGRAPH_NAMESPACES`) for the parent element and then repeating the loop for its child elements is highly inefficient. Elements in standard OOXML typically reside within the same namespace as their immediate parent.
**Action:** Iterate through possible namespaces to locate the parent node. Once found, use that *same* successfully matched namespace URI to execute queries for its child elements to avoid nested loop overhead.
