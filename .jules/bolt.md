## 2026-05-22 - Avoid xpath ancestor:: resolution in @xmldom
**Learning:** Using `xpath` (from the `xpath` NPM package) with `@xmldom/xmldom` is extremely slow when querying with the `ancestor::` axis (e.g., `.//w:p[not(ancestor::w:txbxContent)]`). This causes significant performance bottlenecks for large documents because it traverses the tree for every matched element dynamically instead of just caching parent lookups.
**Action:** When complex ancestor exclusions are needed on large node lists, rely on native DOM APIs (`getElementsByTagNameNS`) combined with a fast `parentNode` while loop in JavaScript. This simple rewrite improved paragraph counting performance by nearly 100x.

## 2026-05-23 - Avoid xpath descendant (//) resolution in @xmldom
**Learning:** Using `xpath` with `@xmldom/xmldom` is extremely slow when querying with descendant paths like `.//w:t` or `.//w:instrText` (e.g. `.//w:del//w:t`), as it evaluates dynamically across large subtrees repeatedly.
**Action:** Replace `xpath` descendant `//` queries with nested native DOM `getElementsByTagNameNS()` calls combined with `Set<Node>` deduplication (since `getElementsByTagNameNS` traverses subtrees and could duplicate matches for nested elements like nested `<w:del>`).
