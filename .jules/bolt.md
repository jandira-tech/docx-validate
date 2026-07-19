## 2026-05-22 - Avoid xpath ancestor:: resolution in @xmldom
**Learning:** Using `xpath` (from the `xpath` NPM package) with `@xmldom/xmldom` is extremely slow when querying with the `ancestor::` axis (e.g., `.//w:p[not(ancestor::w:txbxContent)]`). This causes significant performance bottlenecks for large documents because it traverses the tree for every matched element dynamically instead of just caching parent lookups.
**Action:** When complex ancestor exclusions are needed on large node lists, rely on native DOM APIs (`getElementsByTagNameNS`) combined with a fast `parentNode` while loop in JavaScript. This simple rewrite improved paragraph counting performance by nearly 100x.

## 2026-05-26 - Avoid xpath descendant (//) queries in @xmldom
**Learning:** Using `xpath` with `@xmldom/xmldom` to resolve deep descendants using the `.//` axis (like `$$(".//w:del//w:t")`) is computationally expensive compared to native DOM traversal due to continuous dynamic DOM evaluation and token parsing.
**Action:** Replace `xpath` queries containing descendant axes with native DOM methods like `getElementsByTagNameNS`. Note that since native `getElementsByTagNameNS` acts recursively and can return overlapping nodes if parents are nested (e.g. `<w:del>` inside another `<w:del>`), the results must be placed into a JavaScript `Set` to deduplicate and mimic XPath's node-set semantics safely.
