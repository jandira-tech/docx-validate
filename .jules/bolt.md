## 2026-05-22 - Avoid xpath descendant and ancestor resolution in @xmldom

**Learning:** Using `xpath` (from the `xpath` NPM package) with `@xmldom/xmldom` is extremely slow when querying with the `ancestor::` axis (e.g., `.//w:p[not(ancestor::w:txbxContent)]`) and descendant axes like `.//` (e.g., `.//w:del//w:t`). This causes significant performance bottlenecks for large documents because it traverses the tree for every matched element dynamically instead of caching or direct traversal.
**Action:** When complex descendant inclusions or ancestor exclusions are needed on large node lists, rely on native DOM APIs (`getElementsByTagNameNS`) combined with a fast `parentNode` while loop in JavaScript. This simple rewrite improved paragraph counting and validation performance by 50x-100x.
