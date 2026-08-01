## 2026-05-22 - Avoid xpath ancestor:: resolution in @xmldom

**Learning:** Using `xpath` (from the `xpath` NPM package) with `@xmldom/xmldom` is extremely slow when querying with the `ancestor::` axis (e.g., `.//w:p[not(ancestor::w:txbxContent)]`). This causes significant performance bottlenecks for large documents because it traverses the tree for every matched element dynamically instead of just caching parent lookups.
**Action:** When complex ancestor exclusions are needed on large node lists, rely on native DOM APIs (`getElementsByTagNameNS`) combined with a fast `parentNode` while loop in JavaScript. This simple rewrite improved paragraph counting performance by nearly 100x.

## 2024-05-22 - Avoid xpath descendant (//) searches on large generic node sets

**Learning:** Using `xpath` (from the `xpath` NPM package) with `@xmldom/xmldom` is extremely slow when querying with the descendant `//` operator globally or on dynamic subtrees where elements can nest arbitrarily (like `<w:del>` tracking changes in `.docx`). In benchmarks matching `<w:t>` inside `<w:del>` loops across 5000 elements, it takes over 23 seconds.
**Action:** When searching deep node trees or large documents with simple path exclusions, rely on native DOM traversal APIs (`getElementsByTagNameNS`) and standard iterations combined with a `Set` to deduplicate matched elements. This eliminates tree traversal overhead in pure Javascript land for an almost 1000x improvement (~50ms).
