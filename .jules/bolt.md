## 2026-05-22 - Avoid xpath ancestor:: resolution in @xmldom
**Learning:** Using `xpath` (from the `xpath` NPM package) with `@xmldom/xmldom` is extremely slow when querying with the `ancestor::` axis (e.g., `.//w:p[not(ancestor::w:txbxContent)]`). This causes significant performance bottlenecks for large documents because it traverses the tree for every matched element dynamically instead of just caching parent lookups.
**Action:** When complex ancestor exclusions are needed on large node lists, rely on native DOM APIs (`getElementsByTagNameNS`) combined with a fast `parentNode` while loop in JavaScript. This simple rewrite improved paragraph counting performance by nearly 100x.

## 2026-05-22 - Optimize DOM element counting with exclusion rules
**Learning:** When needing to count elements (like `<w:p>`) while excluding those nested inside specific ancestors (like `<w:txbxContent>` or `<v:textbox>`), iterating over all elements and walking up the `parentNode` chain for each is an O(N*D) operation that performs poorly on deeply nested or large documents.
**Action:** Use a stack-based tree traversal starting from the root. This inherently skips descending into excluded branches (by simply not pushing their children to the stack), completely eliminating the need for `parentNode` checks and dropping execution time drastically (from ~318ms to ~2.5ms in extreme cases).
