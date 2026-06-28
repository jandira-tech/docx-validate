## 2026-05-22 - Avoid xpath ancestor:: resolution in @xmldom
**Learning:** Using `xpath` (from the `xpath` NPM package) with `@xmldom/xmldom` is extremely slow when querying with the `ancestor::` axis (e.g., `.//w:p[not(ancestor::w:txbxContent)]`). This causes significant performance bottlenecks for large documents because it traverses the tree for every matched element dynamically instead of just caching parent lookups.
**Action:** When complex ancestor exclusions are needed on large node lists, rely on native DOM APIs (`getElementsByTagNameNS`) combined with a fast `parentNode` while loop in JavaScript. This simple rewrite improved paragraph counting performance by nearly 100x.

## 2026-05-23 - Avoid xpath descendant (//) resolution in @xmldom
**Learning:** Using `xpath` with descendant queries like `.//w:del//w:t` and `.//w:del//w:instrText` on `@xmldom/xmldom` is extremely slow because it forces the query engine to dynamically traverse the entire subtree for every element. In tests, a single document verification using `xpath` for this specific case took ~22ms, whereas native DOM `getElementsByTagNameNS` took ~7ms—a 3x improvement.
**Action:** When searching for descendants, especially on large documents, use native `getElementsByTagNameNS` loops and deductively add the results to a `Set` (to maintain XPath's unique node-set properties) instead of using `$$(".//...")`.
