## 2026-05-22 - Avoid xpath ancestor:: resolution in @xmldom
**Learning:** Using `xpath` (from the `xpath` NPM package) with `@xmldom/xmldom` is extremely slow when querying with the `ancestor::` axis (e.g., `.//w:p[not(ancestor::w:txbxContent)]`). This causes significant performance bottlenecks for large documents because it traverses the tree for every matched element dynamically instead of just caching parent lookups.
**Action:** When complex ancestor exclusions are needed on large node lists, rely on native DOM APIs (`getElementsByTagNameNS`) combined with a fast `parentNode` while loop in JavaScript. This simple rewrite improved paragraph counting performance by nearly 100x.

## 2026-05-23 - Replacing XPath queries with Native DOM APIs
**Learning:** Using `xpath` (from the `xpath` NPM package) with `@xmldom/xmldom` to query descendants (e.g., `.//w:del//w:t`) causes a performance bottleneck by dynamically searching the tree structure on every query, rather than efficiently crawling sub-structures natively.
**Action:** Replace slow XPath descendant queries with native DOM API methods like `getElementsByTagNameNS` across expected namespaces (e.g., `WORD_PARAGRAPH_NAMESPACES`), taking care to use a `Set<Element>` to deduplicate any matches when querying elements inside nested structures.
