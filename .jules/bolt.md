## 2026-05-22 - Avoid xpath ancestor:: resolution in @xmldom

**Learning:** Using `xpath` (from the `xpath` NPM package) with `@xmldom/xmldom` is extremely slow when querying with the `ancestor::` axis (e.g., `.//w:p[not(ancestor::w:txbxContent)]`). This causes significant performance bottlenecks for large documents because it traverses the tree for every matched element dynamically instead of just caching parent lookups.
**Action:** When complex ancestor exclusions are needed on large node lists, rely on native DOM APIs (`getElementsByTagNameNS`) combined with a fast `parentNode` while loop in JavaScript. This simple rewrite improved paragraph counting performance by nearly 100x.

## 2026-09-22 - Optimize collectDeletedRunText DOM traversal

**Learning:** Searching for specific child elements (like `t`) across the entire document and then checking their parent hierarchy for a specific element (like `del`) is incredibly slow for large documents (O(N) search + O(Depth) traversal per item).
**Action:** When searching for descendants within a specific parent element, optimize DOM traversal by first querying for the parent elements (e.g., `del`) and then querying for the descendants only within those matched parents. Use a `Set` to deduplicate matches due to potential nested structures.
