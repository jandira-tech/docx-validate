## 2026-05-23 - Avoid `ancestor::` XPath queries with `@xmldom/xmldom`
**Learning:** Using the `ancestor::` axis in XPath queries (e.g., `not(ancestor::w:del)`) with `@xmldom/xmldom` and the `xpath` package causes severe performance bottlenecks for large documents.
**Action:** Replace `ancestor::` XPath queries with native DOM traversal APIs (like `getElementsByTagNameNS` combined with a `parentNode` while loop) to significantly improve performance.
