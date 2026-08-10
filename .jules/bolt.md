## 2026-05-22 - Avoid xpath ancestor:: resolution in @xmldom
**Learning:** Using `xpath` (from the `xpath` NPM package) with `@xmldom/xmldom` is extremely slow when querying with the `ancestor::` axis (e.g., `.//w:p[not(ancestor::w:txbxContent)]`). This causes significant performance bottlenecks for large documents because it traverses the tree for every matched element dynamically instead of just caching parent lookups.
**Action:** When complex ancestor exclusions are needed on large node lists, rely on native DOM APIs (`getElementsByTagNameNS`) combined with a fast `parentNode` while loop in JavaScript. This simple rewrite improved paragraph counting performance by nearly 100x.

## 2026-06-03 - Avoid xpath descendant queries in @xmldom
**Learning:** Using `xpath` (from the `xpath` NPM package) with `@xmldom/xmldom` is extremely slow when querying with descendant axes like `.//w:del//w:t`.
**Action:** When finding descendant elements in large XML files, rely on native DOM APIs (`getElementsByTagNameNS`) combined with iterating over standard namespaces (`WORD_PARAGRAPH_NAMESPACES`). This avoids traversing the entire document multiple times via xpath and significantly speeds up validation.
