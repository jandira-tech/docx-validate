## 2026-05-22 - Avoid xpath ancestor:: resolution in @xmldom
**Learning:** Using `xpath` (from the `xpath` NPM package) with `@xmldom/xmldom` is extremely slow when querying with the `ancestor::` axis (e.g., `.//w:p[not(ancestor::w:txbxContent)]`). This causes significant performance bottlenecks for large documents because it traverses the tree for every matched element dynamically instead of just caching parent lookups.
**Action:** When complex ancestor exclusions are needed on large node lists, rely on native DOM APIs (`getElementsByTagNameNS`) combined with a fast `parentNode` while loop in JavaScript. This simple rewrite improved paragraph counting performance by nearly 100x.
## 2026-05-30 - Fix jules pr reviewer action precondition
**Learning:** The Jules PR reviewer Github action will fail with a 400 'FAILED_PRECONDITION' error if the rules_file expected by it is missing from the repository.
**Action:** Create the missing `.github/jules-review-rules.md` file.
