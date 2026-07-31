## 2026-05-22 - Avoid xpath ancestor:: resolution in @xmldom

**Learning:** Using `xpath` (from the `xpath` NPM package) with `@xmldom/xmldom` is extremely slow when querying with the `ancestor::` axis (e.g., `.//w:p[not(ancestor::w:txbxContent)]`). This causes significant performance bottlenecks for large documents because it traverses the tree for every matched element dynamically instead of just caching parent lookups.
**Action:** When complex ancestor exclusions are needed on large node lists, rely on native DOM APIs (`getElementsByTagNameNS`) combined with a fast `parentNode` while loop in JavaScript. This simple rewrite improved paragraph counting performance by nearly 100x.

## 2026-05-23 - Avoid xpath descendant (//) queries in @xmldom

**Learning:** Using `xpath` with `@xmldom/xmldom` is extremely slow when using descendant queries (e.g., `.//w:del//w:t`) because it dynamically searches the tree structure repeatedly.
**Action:** When searching for descendants inside specific nodes (like `<w:t>` inside `<w:del>`), replace the slow xpath query with native `getElementsByTagNameNS()`. Because `getElementsByTagNameNS()` matches against the entire subtree, deduplicate results with a `Set<Element>` when iterating over parent loops to avoid processing inner tags multiple times. This optimization reduces querying time from seconds to milliseconds for large documents.
