## 2026-05-22 - Avoid xpath ancestor:: resolution in @xmldom

**Learning:** Using `xpath` (from the `xpath` NPM package) with `@xmldom/xmldom` is extremely slow when querying with the `ancestor::` axis (e.g., `.//w:p[not(ancestor::w:txbxContent)]`). This causes significant performance bottlenecks for large documents because it traverses the tree for every matched element dynamically instead of just caching parent lookups.
**Action:** When complex ancestor exclusions are needed on large node lists, rely on native DOM APIs (`getElementsByTagNameNS`) combined with a fast `parentNode` while loop in JavaScript. This simple rewrite improved paragraph counting performance by nearly 100x.

## 2026-05-22 - Avoid xpath descendant (//) resolution in @xmldom

**Learning:** Using `xpath` (from the `xpath` NPM package) with `@xmldom/xmldom` is extremely slow when querying with the descendant (`//`) axis (e.g., `$$(".//w:del//w:t", dom)`). This dynamic node-set resolution causes severe performance bottlenecks for large documents, taking hundreds of milliseconds to process nested structures.
**Action:** Replace `$$(".//...//...")` with native DOM APIs (`getElementsByTagNameNS`) using a loop over the namespaces. To replicate exact XPath node-set deduplication for nested tags (like `<w:del>` inside `<w:del>`), use a `Set<Node>`. This reduces execution time from nearly a second to less than ~30ms for 500+ nodes.
