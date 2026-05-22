## 2024-05-22 - XPath ancestor:: axis causes severe performance bottlenecks for large documents

**Learning:** Using `ancestor::` in XPath queries (like `.//w:p[not(ancestor::w:txbxContent)]`) with `@xmldom/xmldom` and the `xpath` package is extremely slow for large documents (e.g. 5000 paragraphs), taking over 140s vs ~20ms with native DOM traversal.
**Action:** Avoid `ancestor::` in XPath queries within this codebase. Instead, use native DOM APIs like `getElementsByTagNameNS` and walk up the `parentNode` chain manually for performance-critical path exclusions.
