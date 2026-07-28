## 2026-05-22 - Avoid xpath ancestor:: resolution in @xmldom
**Learning:** Using `xpath` (from the `xpath` NPM package) with `@xmldom/xmldom` is extremely slow when querying with the `ancestor::` axis (e.g., `.//w:p[not(ancestor::w:txbxContent)]`). This causes significant performance bottlenecks for large documents because it traverses the tree for every matched element dynamically instead of just caching parent lookups.
**Action:** When complex ancestor exclusions are needed on large node lists, rely on native DOM APIs (`getElementsByTagNameNS`) combined with a fast `parentNode` while loop in JavaScript. This simple rewrite improved paragraph counting performance by nearly 100x.

## 2024-05-31 - Extreme XPath Descendant Bottleneck in xmldom
**Learning:** Using `xpath.select` for descendant queries (like `.//w:del//w:t`) on large XML documents parsed with `@xmldom/xmldom` is pathologically slow in this codebase. A synthetic document with 5,000 paragraphs took ~38 seconds using xpath, but only ~110 milliseconds using native `dom.getElementsByTagNameNS()` deduplicated with a Set.
**Action:** When validating or searching XML nodes globally across large documents, always prioritize native DOM traversal methods like `getElementsByTagNameNS` combined with a `Set` (to resolve duplicate yields in nested structures) over `xpath` descendant axes.
