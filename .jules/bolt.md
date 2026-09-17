## 2026-05-22 - Avoid xpath ancestor:: resolution in @xmldom
**Learning:** Using `xpath` (from the `xpath` NPM package) with `@xmldom/xmldom` is extremely slow when querying with the `ancestor::` axis (e.g., `.//w:p[not(ancestor::w:txbxContent)]`). This causes significant performance bottlenecks for large documents because it traverses the tree for every matched element dynamically instead of just caching parent lookups.
**Action:** When complex ancestor exclusions are needed on large node lists, rely on native DOM APIs (`getElementsByTagNameNS`) combined with a fast `parentNode` while loop in JavaScript. This simple rewrite improved paragraph counting performance by nearly 100x.

## 2026-05-22 - Optimize DOM collection in docx insertions/deletions validation
**Learning:** `getElementsByTagNameNSAll` iterates through the whole document to find matching tags. When looking for nested matching tags in `collectDeletedRunText` and `validateInsertions`, it is more performant to retrieve the parent first (e.g., `<w:del>` or `<w:ins>`) and then query the descendants directly inside that node, rather than query the whole document and then check if it's inside the parent node. Using a `Set` handles potential deduplication.
**Action:** Always optimize nested tag lookups by utilizing parent queries directly instead of checking parents of all document-wide elements.
