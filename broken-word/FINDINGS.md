# broken-word: fixtures Microsoft Word rejects

Ground truth = real MS Word (macOS probe). `reference/` = pristine copies of the
original false-negatives (do not edit). `playground/` = experiment copies + proven
fixes. `word-rejected/` = copies of ALL 65 fixtures Word does not open cleanly
(full-corpus probe), with `MANIFEST.csv` (Word outcome + validator verdict).

## What shipped
`ignorable-undeclared` is now an `error` under the `word-valid` profile
(`src/scripts/office/validators/docx.ts`, `isWordBlockingIssue`). Word refuses to
open a package whose `mc:Ignorable` lists an undeclared prefix (e.g.
`mc:Ignorable="w14 w15 wp14"` with only `w14` declared). Proven by experiment
(`playground/sample.ignorable-fixed.PROVEN-clean-in-word.docx` opens clean) and
safe: exactly 3 fixtures carry the code, all 3 are Word-rejected, zero Word-clean.

## What was tried and REVERTED (do not re-add without new evidence)
Escalating `xsd-error` for malformed `docProps/*` atomic values (e.g.
`<TotalTime>3\n  </TotalTime>`, invalid `xs:int`). It creates false positives:
`endnotes.paraid-overflow.docx` and `empty.missing-content-type.docx` carry the
same malformed-integer defect yet Word opens them **clean**.

Also NOT escalated: the comment-integrity codes (`comment-orphan-start`,
`comment-marker-missing`, `comment-thread-count-mismatch`). They fire on Word-CLEAN
files too (`broken/` and `word-regenerate-invalid/regenerated/` comment copies open
fine with the orphan markers present), so escalating them would create false
positives and misses the real cause.

## The residual mismatches are genuinely NOT safely fixable
Full-corpus Word probe: validator agrees with Word on **567/573 (98.95%)**. The 6
residual mismatches (≈3 distinct files, each duplicated across fixture trees) sit in
a zone of Word non-determinism:
- `comments.unmatched-comment-marker.docx`, `footnotes.duplicate-paraid.docx`,
  and the rejected `tiny-picture.broken-rels.docx` variant all share pretty-printed
  `docProps/app.xml` with malformed integers, and Word rejects them.
- BUT `endnotes.paraid-overflow.docx` (Word-clean) has the same malformed-`TotalTime`
  defect; and fixing `comments`' Pages+Words while leaving `TotalTime` malformed
  STILL makes Word reject it. Same single malformed field, opposite Word verdicts.
- The 2 false positives are clean `tiny-picture.broken-rels` variants: a broken
  `media/*.png` reference, which Word tolerates (missing-image placeholder). The
  `rels-broken` rule blocks it anyway. Narrowing `rels-broken` to `customXml` only
  would fix those 2 FP but turn the 2 genuinely-rejected variants into false
  negatives (their only error code is also `rels-broken`) — net zero, and FN is the
  worse error for a validator.

Conclusion: no escalation distinguishes the Word-broken from the Word-clean files
here without creating offsetting errors. Left as documented known edge cases rather
than degrading the validator. Verified independently by an adversarial review agent
(all findings survived; the `rels-broken` heuristic confirmed as the shakiest rule).

## Original root-cause proofs (still valid, single-variable, Word-confirmed)
- `Sample Document.docx` → `mc:Ignorable` undeclared prefixes (FIXED, shipped).
- `comments.unmatched-comment-marker.docx` → malformed `docProps/app.xml`; the
  orphan `commentRangeStart id=999999` is NOT the cause (Word opens the file clean
  with it intact once app.xml whitespace is collapsed; verified with a re-zip control).
