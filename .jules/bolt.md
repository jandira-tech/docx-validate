## 2026-05-22 - Avoid xpath ancestor:: resolution in @xmldom
**Learning:** Using `xpath` (from the `xpath` NPM package) with `@xmldom/xmldom` is extremely slow when querying with the `ancestor::` axis (e.g., `.//w:p[not(ancestor::w:txbxContent)]`). This causes significant performance bottlenecks for large documents because it traverses the tree for every matched element dynamically instead of just caching parent lookups.
**Action:** When complex ancestor exclusions are needed on large node lists, rely on native DOM APIs (`getElementsByTagNameNS`) combined with a fast `parentNode` while loop in JavaScript. This simple rewrite improved paragraph counting performance by nearly 100x.

## 2026-10-03 - Avoid `getElementsByTagName('*')` in `xmldom`
**Learning:** Calling `getElementsByTagName('*')` on large parsed XML documents triggers an expensive pre-order traversal in `@xmldom/xmldom` to create a live `NodeList`. Accessing its elements via `list.item(i)` dynamically re-evaluates the list, resulting in $O(N^2)$ traversal times when heavily nested elements are queried.
**Action:** Always replace `dom.getElementsByTagName('*')` with a custom iterative tree-walking algorithm (DFS) that pushes elements to a static array, reducing the traversal time to $O(N)$.
