## 2026-05-22 - Avoid xpath ancestor:: resolution in @xmldom
**Learning:** Using `xpath` (from the `xpath` NPM package) with `@xmldom/xmldom` is extremely slow when querying with the `ancestor::` axis (e.g., `.//w:p[not(ancestor::w:txbxContent)]`). This causes significant performance bottlenecks for large documents because it traverses the tree for every matched element dynamically instead of just caching parent lookups.
**Action:** When complex ancestor exclusions are needed on large node lists, rely on native DOM APIs (`getElementsByTagNameNS`) combined with a fast `parentNode` while loop in JavaScript. This simple rewrite improved paragraph counting performance by nearly 100x.

## 2026-07-23 - XPath Code Review Hallucinations
**Learning:** Automated code reviews may flag false positives (e.g., claiming 'makeSelect' was improperly removed, or 'WORD_PARAGRAPH_NAMESPACES' is undeclared) when in fact the implementation matches the surrounding file's patterns and imports perfectly. The reviewer's static analysis might hallucinate or assume context that is not accurate.
**Action:** Always independently verify code review claims (using `grep`) against the actual codebase rather than blindly accepting that a variable is undeclared or an import is unsafe.
