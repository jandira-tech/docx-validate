## 2026-05-22 - Avoid xpath ancestor:: resolution in @xmldom
**Learning:** Using `xpath` (from the `xpath` NPM package) with `@xmldom/xmldom` is extremely slow when querying with the `ancestor::` axis (e.g., `.//w:p[not(ancestor::w:txbxContent)]`). This causes significant performance bottlenecks for large documents because it traverses the tree for every matched element dynamically instead of just caching parent lookups.
**Action:** When complex ancestor exclusions are needed on large node lists, rely on native DOM APIs (`getElementsByTagNameNS`) combined with a fast `parentNode` while loop in JavaScript. This simple rewrite improved paragraph counting performance by nearly 100x.

## 2026-05-22 - Optimize DOM traversal for descendant elements
**Learning:** Searching an entire large XML document for a specific element (like `<w:t>`) and then repeatedly checking if each node belongs to a certain parent (`<w:del>`) causes excessive tree iterations. For documents with many paragraphs but few deletions, this is a major bottleneck.
**Action:** When querying for descendants that must exist within specific parents, first query for the parents (`getElementsByTagNameNSAll(root, ns, 'del')`) and then query for the descendants only within those matched parents. Use a `Set` to deduplicate matches correctly in case of nested structures.
