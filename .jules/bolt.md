## 2026-05-22 - Avoid xpath ancestor:: resolution in @xmldom
**Learning:** Using `xpath` (from the `xpath` NPM package) with `@xmldom/xmldom` is extremely slow when querying with the `ancestor::` axis (e.g., `.//w:p[not(ancestor::w:txbxContent)]`). This causes significant performance bottlenecks for large documents because it traverses the tree for every matched element dynamically instead of just caching parent lookups.
**Action:** When complex ancestor exclusions are needed on large node lists, rely on native DOM APIs (`getElementsByTagNameNS`) combined with a fast `parentNode` while loop in JavaScript. This simple rewrite improved paragraph counting performance by nearly 100x.

## 2026-06-17 - Avoid xpath descendant (//) searches in @xmldom
**Learning:** Using `xpath` with `@xmldom/xmldom` is extremely slow when querying with the descendant (`//`) axis (e.g., `.//w:del//w:t`). It causes severe performance bottlenecks for large documents (O(N^2) dynamic tree traversal).
**Action:** Replace `xpath` descendant searches with native DOM APIs (`getElementsByTagNameNS`) using known namespaces. This improved tracked changes validation performance by almost 500x on large simulated documents.
