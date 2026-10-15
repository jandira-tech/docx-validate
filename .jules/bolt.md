## 2026-05-22 - Avoid xpath ancestor:: resolution in @xmldom

**Learning:** Using `xpath` (from the `xpath` NPM package) with `@xmldom/xmldom` is extremely slow when querying with the `ancestor::` axis (e.g., `.//w:p[not(ancestor::w:txbxContent)]`). This causes significant performance bottlenecks for large documents because it traverses the tree for every matched element dynamically instead of just caching parent lookups.
**Action:** When complex ancestor exclusions are needed on large node lists, rely on native DOM APIs (`getElementsByTagNameNS`) combined with a fast `parentNode` while loop in JavaScript. This simple rewrite improved paragraph counting performance by nearly 100x.

## 2026-05-22 - Optimize DOM element counting with exclusion rules

**Learning:** When needing to count elements (like `<w:p>`) while excluding those nested inside specific ancestors (like `<w:txbxContent>` or `<v:textbox>`), iterating over all elements and walking up the `parentNode` chain for each is an O(N*D) operation that performs poorly on deeply nested or large documents.
**Action:** Use a stack-based tree traversal starting from the root. This inherently skips descending into excluded branches (by simply not pushing their children to the stack), completely eliminating the need for `parentNode` checks and dropping execution time drastically (from ~318ms to ~2.5ms in extreme cases).

## 2026-06-03 - O(N) targeted querying for redlining stripping

**Learning:** In large documents, stripping tracked changes using `collectAllElements` (which does a deep `getElementsByTagName('*')`) and iterating child-by-child is O(N²) and extremely slow.
**Action:** Replace the custom full-tree collection and child filtering with direct, targeted native `getElementsByTagNameNSAll(root, NS.W, 'ins')` / `del` queries. Processing them in a single pass is O(N) and resulted in a 2x-5x speed improvement.

## 2026-10-06 - Avoid .item(i) on getElementsByTagNameNS in @xmldom

**Learning:** `@xmldom/xmldom`'s live `NodeList` caches its contents until the document changes (`_updateLiveList` compares the document's `_inc`). A read-only `.item(i)` loop is therefore O(N): measured on 10k-80k `w:r` elements it scales linearly, and the flat array is only 1.7-3x faster. A loop that edits the DOM while iterating (even `setAttribute`) refreshes the whole list after every edit: 2000 `setAttribute` calls over an 80k-run document took 62 s. That is the quadratic case.
**Action:** Implement `getElementsByTagNameNSAll` with a stack-based DFS traversal. This processes the DOM in a single O(N) pass, completely eliminating `.item(i)` and providing nearly a 50% speedup on large document iterations.

## 2026-10-06 - Optimize DOM collection in docx insertions/deletions validation

**Learning:** `getElementsByTagNameNSAll` iterates through the whole document to find matching tags. When looking for nested matching tags in `collectDeletedRunText` and `validateInsertions`, it is more performant to retrieve the parent first (e.g., `<w:del>` or `<w:ins>`) and then query the descendants directly inside that node, rather than query the whole document and then check if it's inside the parent node. Using a `Set` handles potential deduplication.
**Action:** Always optimize nested tag lookups by utilizing parent queries directly instead of checking parents of all document-wide elements. (Ported from the rejected bolt-optimize-xml-lookups PR — the code landed via the del-first `collectDeletedRunText`, this entry preserves the learning.)

## 2026-10-11 - Do not re-propose .item(i) rewrites as O(N²) fixes

**Learning:** PRs #188 and #190 proposed the same `getElementsByTagNameNSAll` rewrite; #190 landed. Read-only loops gain a constant factor only; repair loops that mutate while iterating are where the snapshot matters, and only when the mutation does not add or remove matching nodes (collect first, then remove, as `removeMcAlternateContent` does).
**Action:** Before claiming a complexity win, measure at two sizes and say which loops mutate. Do not reformat files in a performance PR (#188 rewrote 5,900 lines to two-space indentation).

## 2026-10-15 - Single-pass DFS traversal over post-filtered DOM querying
**Learning:** Parent traversal (`parentNode` loops) for every matched node in a deep document is O(N*D). Combining this with querying all elements leads to huge performance degradation.
**Action:** Implement an integrated top-down DFS traversal instead of querying all elements and then checking their parent chains. This removes redundant parent tree traversals and prunes excluded subtrees. End-to-end timings can still be dominated by XML parsing and allocation; see `docs/performance/redlining-dfs.md` for mixed local measurements.
