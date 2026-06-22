## 2026-05-22 - Avoid xpath ancestor:: resolution in @xmldom
**Learning:** Using `xpath` (from the `xpath` NPM package) with `@xmldom/xmldom` is extremely slow when querying with the `ancestor::` axis (e.g., `.//w:p[not(ancestor::w:txbxContent)]`). This causes significant performance bottlenecks for large documents because it traverses the tree for every matched element dynamically instead of just caching parent lookups.
**Action:** When complex ancestor exclusions are needed on large node lists, rely on native DOM APIs (`getElementsByTagNameNS`) combined with a fast `parentNode` while loop in JavaScript. This simple rewrite improved paragraph counting performance by nearly 100x.

## 2026-06-22 - Avoid xpath descendant axis (.//) in @xmldom
**Learning:** Using the xpath descendant axis (`.//`) in the `xpath` NPM package with `@xmldom/xmldom` evaluates extremely slowly compared to native DOM traversals, similarly to the `ancestor::` axis. For example, querying `.//w:del//w:t` on a document with 1000 items takes ~800ms compared to ~15ms with native `getElementsByTagNameNS`.
**Action:** When querying descendants, avoid xpath descendant axes where possible. Instead, use native DOM traversal methods like `getElementsByTagNameNS` directly.
