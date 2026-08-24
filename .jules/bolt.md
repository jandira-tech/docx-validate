## 2026-05-22 - Avoid xpath ancestor:: resolution in @xmldom

**Learning:** Using `xpath` (from the `xpath` NPM package) with `@xmldom/xmldom` is extremely slow when querying with the `ancestor::` axis (e.g., `.//w:p[not(ancestor::w:txbxContent)]`). This causes significant performance bottlenecks for large documents because it traverses the tree for every matched element dynamically instead of just caching parent lookups.
**Action:** When complex ancestor exclusions are needed on large node lists, rely on native DOM APIs (`getElementsByTagNameNS`) combined with a fast `parentNode` while loop in JavaScript. This simple rewrite improved paragraph counting performance by nearly 100x.

## 2026-05-22 - Avoid repeated ancestor traversals with getElementsByTagNameNS

**Learning:** When trying to find elements but excluding those inside specific parent trees (like <w:p> inside <w:txbxContent>), combining getElementsByTagNameNS with a parentNode while-loop for every matched element causes O(N _ depth) traversals and severe performance bottlenecks. For large documents, traversing up from every element is extremely inefficient.
**Action:** Replace getElementsByTagNameNS + upward traversal with a downward iterative Depth-First Search (DFS) that skips entire subtrees when it encounters an excluded parent tag. This turns an O(N _ depth) operation into an O(N) single-pass operation.
