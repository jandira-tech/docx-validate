# [docxodus 0888] failed: no document defect; crashes_word label is a harness false positive

## Summary

`tool_redlines/crashes_word/docxodus__01a11a49_0888.docx` was a docxodus-produced tracked-changes
redline (79 `w:ins` / 4973 `w:del`, author `Open-Xml-PowerTools`, `w:date 2000-01-01T00:00:00Z`)
over a document with three headers, footnotes, endnotes and two SmartArt diagram sets
(`word/diagrams/data|colors|layout|quickStyle|drawing 1–2`, per the preserved part listing at
`.scratch/libxml2-copy-unique-paths.txt:652-680`).

**No defect was found in it.** The `crashes_word` lane label rests on a manifest row quoting
`Microsoft Word got an error: Connection is invalid` (`d1_failed_in: soffice`) — an AppleScript
automation-transport failure, not a document crash — and is contradicted by the lane's own proven
assessment: Word opened the original clean (`docs/really-broken/assessments.jsonl:17`,
`word.original: clean-open`).

This session adds a provenance complication: **the specimen is gone from the repo.** The live Word
check could not open it, and neither can any re-validation:

- `really_broken/tool_redlines/` contains only `repair_prompt/` (no `crashes_word/` directory);
- repo-wide `find … -name "*0888*"` returns nothing; the unpacked copy
  `really_broken/unpacked/crashes_word__docxodus__01a11a49_0888` does not exist;
- `really_broken/manifest.csv` does not exist (no `*.csv` in `really_broken/` at all);
- the analyst's validate command now errors: `pnpm exec tsx src/scripts/office/validate.ts
  really_broken/tool_redlines/crashes_word/docxodus__01a11a49_0888.docx` →
  `Error: … does not exist`.

Only the proven-cause records (`docs/really-broken/assessments.jsonl`, `inspection-log.jsonl`) and
the scratch part listing survive.

## Word behaviour

- **Live result (this ask):** never opened — outcome `failed`, because the file does not exist.
  Absence re-verified first-hand with `ls really_broken/tool_redlines/` and the repo-wide `find`
  above. No Word observation was or can be made until the specimen is restored.
- **Manifest row (ask-given, quoted by the analyst before the file was deleted):**
  `really_broken/manifest.csv:32` — category `crashes_word`, local_word_error
  `Microsoft Word got an error: Connection is invalid`, `d1_failed_in: soffice`. That string is the
  AppleScript error text for a lost application connection (Word quit or the event timed out
  mid-run); it says nothing about the document.
- **Prior proven assessment (read this ask, the established Word cause):**
  `docs/really-broken/assessments.jsonl:17` — `status: proven`, `defects: []`, `defect: "None.
  Microsoft Word opens the original file with no warning, error, or offer to repair."`,
  `change: none`, `word.original: clean-open`; mirrored at `docs/really-broken/inspection-log.jsonl:17`.
  It credits no docx-validate check, because there is no defect to attribute. The crash label was
  never the document's.
- **Cohort context (counted this ask):** 16 `tool_redlines/crashes_word/` rows in
  `assessments.jsonl`; 8 are proven clean-open no-defect files (docxodus 0876/**0888**/1860/1862 and
  jubarte 0876/1360/1862/1990), 8 carry real defects (e.g. docxodus 2183: `w:delText` with no
  `w:del` in a header2 text box; docxodus 2243: `w:delInstrText`/`w:delText` in footer1). This file
  sits squarely in the mislabeled half, so the misclassification is systematic, not one-off.

## Violations

**None reported, none found — there is no defect to reproduce.** The analyst's `violations` list is
empty. What was verified first-hand this ask vs. not run:

| Claim | Source | This ask |
| --- | --- | --- |
| Proven record: `proven`, `defects: []`, `clean-open` | `assessments.jsonl:17`, `inspection-log.jsonl:17` | **Reproduced** — read both full rows |
| Label contradicted by the record | same rows | **Reproduced** |
| Specimen, unpacked copy and manifest all absent | filesystem | **Reproduced** — `ls`, repo-wide `find`, validator error above |
| Part list (headers 1–3, footnotes, endnotes, 2 SmartArt sets) | `.scratch/libxml2-copy-unique-paths.txt:652-680` | **Reproduced** (names only; bytes gone) |
| Full validate run → `All validations PASSED!` | analyst (ask-given) | **Not run by me** — file absent; my re-run errored `does not exist` |
| `xmllint --schema ISO-IEC29500-4_2016/wml.xsd` on `word/document.xml`, header1–3, footnotes, endnotes → only noise | analyst (ask-given) | **Not run by me** — parts absent |
| Redline structure: 79 `w:ins` / 4973 `w:del`; 0 of 5902 `w:delText` outside a `w:del` | analyst (ask-given) | **Not run by me** — parts absent |
| Live Word open | live-result agent (ask-given) | **Failed** — file missing, never opened |

The XSD findings the analyst did see are all known non-defects, not docxodus bugs: Microsoft
2010/2011 extension attributes (`w14:paraId`/`w14:textId` on `w:p`; `wp14:anchorId`/`wp14:editId`
on `wp:inline`) and MCE markup (`mc:Ignorable`; six `mc:AlternateContent` SmartArt Choice/Fallback
sites) that the ISO transitional XSDs under `src/scripts/office/schemas/ISO-IEC29500-4_2016/` do
not model. Those constructs live in the bundled Microsoft extension schemas
(`src/scripts/office/schemas/microsoft/wml-*.xsd`), and this repo's schema walk already applies
markup compatibility before validating a part — which is why its own run passed.

## Where the fixes belong and why

- **really-broken lane harness (d1 Word-driving stage + manifest writer) — the one real defect.**
  A stage whose AppleScript connection dropped recorded `crashes_word` for a document Word opens
  clean. Fix: classify transport/automation errors (`Connection is invalid`, AppleScript timeouts,
  app-quit mid-run) as `unknown` with a retry, never as a document crash; only a Word log showing a
  crash report or dialog for these exact bytes may record a crash. 8 of the 16 `crashes_word` rows
  in `assessments.jsonl` are proven clean-opens — the regression corpus for this fix already exists.
- **Not docx-validate.** No check is missing: the document is valid and the full run (analyst
  evidence) correctly passed. A check that fired here would itself be the false positive.
- **Not docxodus.** The generated redline is valid OOXML; its only schema deltas are Microsoft
  extension attributes and standard MCE wrapping that Word itself emits and accepts.
- **Not repair.** The document opens clean; there is nothing to repair.

## Verification

1. **Nothing left to prove about the document** — the clean-open cause is the lane's own proven
   record (`assessments.jsonl:17`), which per this repo's probe protocol already rests on two
   agreeing Word runs.
2. **To re-establish reproducibility** (blocked until the specimen is restored from the source
   corpus that produced the assessments): re-run
   `pnpm exec tsx src/scripts/office/validate.ts <restored-file>` — expect `All validations PASSED!`
   (the analyst's result) — and one two-run Word open per the Real-Word probes rule; expect
   `clean-open`, matching the record. If either disagrees, reopen this issue with the logs.
3. **For the harness fix:** a unit test that feeds the d1 stage the recorded
   `Microsoft Word got an error: Connection is invalid` result and asserts the recorded outcome is
   `unknown` (retryable), not a crash category; then re-run the cohort's 8 clean-open files through
   the stage and assert none lands in `crashes_word`.
