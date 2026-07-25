## 2026-05-22 - Avoid xpath ancestor:: resolution in @xmldom
**Learning:** Using `xpath` (from the `xpath` NPM package) with `@xmldom/xmldom` is extremely slow when querying with the `ancestor::` axis (e.g., `.//w:p[not(ancestor::w:txbxContent)]`). This causes significant performance bottlenecks for large documents because it traverses the tree for every matched element dynamically instead of just caching parent lookups.
**Action:** When complex ancestor exclusions are needed on large node lists, rely on native DOM APIs (`getElementsByTagNameNS`) combined with a fast `parentNode` while loop in JavaScript. This simple rewrite improved paragraph counting performance by nearly 100x.

## 2026-05-23 - Avoid xpath query in validateDeletions
**Learning:** Using `xpath.select` (e.g. `$$(".//w:del//w:t", dom)`) is unnecessarily slow for simple descendant queries in large documents. The overhead of `@xmldom/xmldom` combined with `xpath` causes measurable delays when processing elements like `<w:del>` which frequently appear in tracked changes.
**Action:** Replace `xpath.select` with native DOM APIs (`getElementsByTagNameNS`) and use a `Set` to deduplicate nodes when iterating over multiple namespaces. This significantly speeds up validation tasks on docx files without adding much code complexity.
