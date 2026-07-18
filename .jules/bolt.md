## 2026-05-22 - Avoid xpath ancestor:: resolution in @xmldom
**Learning:** Using `xpath` (from the `xpath` NPM package) with `@xmldom/xmldom` is extremely slow when querying with the `ancestor::` axis (e.g., `.//w:p[not(ancestor::w:txbxContent)]`). This causes significant performance bottlenecks for large documents because it traverses the tree for every matched element dynamically instead of just caching parent lookups.
**Action:** When complex ancestor exclusions are needed on large node lists, rely on native DOM APIs (`getElementsByTagNameNS`) combined with a fast `parentNode` while loop in JavaScript. This simple rewrite improved paragraph counting performance by nearly 100x.
## 2026-07-18 - Avoid xpath descendant:: resolution in @xmldom
**Learning:** Using `xpath` with `@xmldom/xmldom` is extremely slow when querying with descendant axes like `.//w:del//w:t`. This causes O(n) traversals on each match, scaling extremely poorly on large files.
**Action:** When querying for known elements, iterate over the known namespaces and use native DOM APIs (`getElementsByTagNameNS`) instead of xpath descendant queries.
