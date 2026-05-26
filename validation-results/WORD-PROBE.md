# Word-adjudicated truth: docx-validate vs OpenXML SDK on the disagreement set

When docx-validate (`--profile word-valid`) and Microsoft's OpenXML SDK disagree, **real
Microsoft Word is the tiebreaker.** Opened all 200 disagreement fixtures in Word (macOS,
`scripts/probe-word-fixtures.ts`). Raw log: `word-probe-disagreements.jsonl`.

## Result: Word sided with us on 197/200 (98.5%)

| | count | Word outcome | verdict |
|---|--:|---|---|
| **We reject, SDK accepts** (FP vs SDK) | 7 | **all 7 → Word "unreadable content"** | **we were right, SDK too lenient** |
| **We accept, SDK rejects** (FN vs SDK) | 193 | 190 clean-open / 3 warned | 190 → **we right, SDK too strict** |
| **Genuine misses** (we accept, Word rejects) | 3 | unreadable-content-warning | **real gaps to fix** |

Word outcomes overall: `clean-open` 190, `unreadable-content-warning` 10. No crashes, errors,
timeouts, or password prompts.

So the SDK's strictness produced **190 "errors" that Word does not care about**, and missed
nothing on these files that we also missed. Our 7 "false positives vs the SDK" were not false at
all — Word rejects them too.

## The 3 genuine misses (we said valid, Word shows "unreadable content")

1. `tests/fixtures/unknown/unknown/comments.unmatched-comment-marker.docx` (and its
   `word-regenerate-invalid/original/...` copy) — **we DO detect it**, but only at `warning`:
   `comment-orphan-start`, `comment-marker-missing`, `comment-thread-count-mismatch`
   (orphan `commentRangeStart id=999999`, marker → non-existent comment, range/Count mismatch).
   Word actually refuses to open it cleanly. **Fix: escalate these comment-integrity codes to
   `error` under the `word-valid` profile** — we already see the problem, just under-classify it.
2. `fixtures/eigen-extended/Sample Document.docx` — we emit only info/warning (e.g.
   `run-props-redundant`); Word flags unreadable content. Needs investigation into which
   construct Word rejects (SDK reported 106 schema errors here).

## Note on tooling

`scripts/probe-word-fixtures.ts` was extended this session with a distinct **`word-crashed`**
outcome (process launches then vanishes → recorded as a crash, not a timeout) plus a crash-reporter
dismisser, per request. No crashes occurred in this run, but the capability is now in place. This is
an uncommitted change on the current branch.
