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

## 2026-10-25 - What the salvage of PRs #198-#202, #205, #265, #421 and #633 kept

**Learning:** Of these nine PRs, one change survived measurement: a single walk for the deleted-text checks (#199), 2.3-3.6x faster on the text-box check and 1.9-2.0x on the move-source check it now also covers. Even that is about 0.4% of a 12-second word-valid run on a 2.4 MB document. The `childNodes` loop in `getElementsByTagNameNSAll` (#633) measured 0.89-1.76x, with no gain on synthetic documents. The duplicate-ID DFS (#265) measured 0.93-1.18x, because the ancestor walk it removes runs only for elements that carry an ID. The `collectText` DFS (#201, #198) ran 1.2-1.5x in isolation and no faster when building the inventory, against an A/A noise floor of 5-6%. Querying the containers first (#421, #200, #202) was slower than one walk, and the first versions missed deleted text in the other Word namespace. #205 swapped qualified-name matching for namespace matching that lost Strict documents.
**Action:** Benchmark both the function itself and the end-to-end path that calls it, with interleaved samples and an A/A run for the noise floor. Propose a change only when it beats that floor. When replacing an ancestor walk, check both Word namespaces. Append to this journal; never rewrite it (#633 deleted every earlier entry).

## 2026-10-28 - #677, #686 and #936 re-proposed rejected changes

**Learning:** Three PRs opened after the #674 salvage repeated changes the 2026-10-25 entry had already measured and rejected. #677 queried the `txbxContent`/`moveFrom` containers first, like #421, #200 and #202, though one walk for both checks had landed in 58e53b73. #686 and #936 replaced the `lastChild`/`previousSibling` walk in `getElementsByTagNameNSAll` with `childNodes` iteration, which is #633's change. None reported an end-to-end measurement. All three also edited workflow files (#1027).
**Action:** Read this journal before proposing a change. A traversal rewrite that an entry here rejects needs interleaved end-to-end measurements that beat the A/A noise floor. A performance PR touches no workflow or rules file.

## 2024-05-15 - [Optimize getElementsByTagNameNSAll]

**Learning:** In heavily used XML DOM traversal utilities, repeatedly pushing/popping every descendant node just to filter out the vast majority based on namespace/localName is a major bottleneck (nearly 2s for 500 calls on 10k nodes). By flattening the inner loop into specific branches based on `namespaceURI === "*"` and `localName === "*"`, we can hoist the condition outside the loop and avoid checking every element's namespace when not needed, reducing time to ~880ms for a specific name lookup and saving ~150-200ms per 500 calls even for wildcard lookups.
**Action:** Always prefer unrolling conditions outside of tight loops (like tree walks) to produce the most efficient branchless-like paths for the hottest execution branches.
