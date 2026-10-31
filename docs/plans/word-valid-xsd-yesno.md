# Word-valid yes/no: skip XSD when a cheaper check already fails

Arthur, the first suggestion is wrong for this corpus. "On these 76 files, return before XSD starts" does not hold. The word-valid probe of `word-invalid-fixtures/` (76 files, all rejected by Word and by `profile: "word-valid"`) splits like this:

- 50 files have no `xsd-error`.
- 2 files have `xsd-error` plus `docprops-boolean-invalid`.
- 24 files have `xsd-error` as their only blocking code. Example: `OPEN_ERROR/tests__fixtures__broken__header-with-rid-no-sidecar.docx`.

52 files can return before XSD. 24 still start it. A transitional file that passes word-valid still walks XSD. Timing the bad corpus and calling that "the speedup" would lie about a clean file. The two numbers are measured separately, each against the parent revision's full walk (`8aec85e`), and the report quotes the settled tail after a warmup, not the median of the ramp. The old ~10.9 s raw full-part XSD scan is not the baseline.

## What ships

1. On `word-valid`, run the blocking non-XSD checks first. If `applyWordValidProfile` is already invalid, do not call `validateAgainstXsd`. The other non-XSD checks still run; only XSD is skipped. Lenient and strict keep today's `Promise.all`, including XSD.
2. When XSD still runs on word-valid, stop at the first Word-blocking `xsd-error`. Later parts are not sent to the engine. A file whose XSD issues are only non-blocking, or only matched by the ignore list, is walked far enough to know that and stays valid. Lenient and strict still walk every eligible part and still return the messages.
3. A schema-valid part asks libxml2 for `xmlSchemaValidateDoc`'s status and returns no issues when the status is 0. That path does not call `XsdValidator.validate` and does not build an `XmlValidateError`. It does install a no-op structured callback, because libxml2's default handler otherwise prints schema errors to stderr. The callback does not build a detail object. Status above 0 still collects messages with a second call, because the integer cannot tell a Word-blocking misplaced element from a Dublin Core fault or a Word-tolerated `pgSz`. Invalid parts are validated twice. A schema-load failure returns before the status probe, is an error, and is never a silent pass. The ignore needles `hyphenationZone`, `purl.org/dc/terms`, `dublincore.org`, and `Invalid XSD schema` still suppress that known false positive. The list is empty since #1209.
4. One reused native libxml2 validation context with no error handler is timed in scratch, against the libxml2-wasm hot path, on the same parts. It is not a product dependency. The Rust `libxml` crate is not this measurement: that wrapper builds a diagnostics vector per part and was about three times slower. A slower native number is reported as slower.

## Edges the tests lock

Tests call `validate(path, { profile })` from `src/scripts/office/validate.ts`. The result carries `xsdWalk` attached after `mergeResults`: `started` (entered `validateAgainstXsd`), `parts` (relative paths actually sent to the engine), `eligible` (how many parts a full walk would send). Strict-namespace files and unmapped parts are not engine parts. A missing trace fails the assertion.

- Comment-thread orphans (`sample-document.empty-rels.docx`) are a no, and XSD does not start.
- `xsd-error`-only (`header-with-rid-no-sidecar.docx`) still starts XSD.
- A later part added after a Word-blocking misplaced element is not sent to the engine on word-valid. Lenient and strict still send it, and still return `xsd-error` messages.
- A single part whose first three schema strings are tolerated (`pgMar` missing attributes) and whose later string is Word-blocking (`tblCellSpacing`) still returns invalid. That blocking string is one of the emitted `xsd-error` issues, and the walk still stops before a later part. Stopping on an unsurfaced error used to leave only warnings.
- `word-tolerated-misplaced-pgsz.docx` stays valid, starts XSD, and finishes the walk (`parts.length === eligible`).
- `word-clean-strict01.docx` mixes a strict `word/document.xml` with transitional package parts. The walk records `xsd-strict-skipped`, does not send `word/document.xml`, may send the transitional parts, finishes that eligible set (`parts.length === eligible`), and stays valid. Strict skip is per file, not a whole-package skip.
- Ignore-list needles do not turn that passing file (it has `docProps/core.xml`) into a failure.
- A missing schemas directory on an otherwise passing file is an error (`Schema load failed`), not a pass.
- Body-level `m:sPre` (`math-body-level.docx`) stays `word-math-spre-body` and does not need XSD for the no. The same `m:sPre` inserted inside a `w:p` does not produce that code.
- A broken media relationship stays a pass. A broken `customXml` relationship (`shapes/vectors.docx`) stays a no, and XSD does not start.
- An OLE/CFB package is `package-open-failed` and does not start XSD.

`xsdValidateStats.statusOnlyPasses` counts engine calls that returned on status 0. `messageCollections` counts calls that still built messages. A word-valid pass must increase `statusOnlyPasses`. A Word-blocking schema error must increase `messageCollections`.

## What this does not do

- It does not replace libxml2-wasm, add a native addon, or add the Rust `libxml` crate.
- It does not patch bundled XSDs (`pitchFamily`, Dublin Core imports).
- It does not stop early or drop messages on lenient or strict.
- It does not change which document defects count as Word-blocking. Schema-load failure stays an error because a validator must not call a document valid against a schema it could not load (AGENTS.md note 4). That is an engine failure, not a new Word-defect code.
- It does not open Word again.

## Measurement

Filled from `{SCRATCH}/timings.log` and `{SCRATCH}/libxml2-reused-context.log` before the PR. Bad corpus and clean file are separate. Skip count and enter count are printed from the shipped `xsdWalk` on the 76 files. Expected skip 52, enter 24. If the measured counts differ, the measured counts win and the reason is stated here.

### Result

Same machine. Parent is `8aec85e`, full walk. New path is this branch's `validate()`. Each process warms the schema cache on the 76-file corpus before the clean repeats. Both tails are hot-cache. The bad-corpus time is a different number from the clean-file time. The old ~10.9 s raw full-part XSD scan is not this baseline.

Of the 76 files, 52 returned before XSD and 24 entered XSD, on every repeat (`xsd_skip_counts 52,52,52,52,52` and `xsd_enter_counts 24,24,24,24,24`). Suggestion 1 does not skip XSD on the 24 files whose only blocking code is `xsd-error`.

Bad corpus, 76 files, all invalid, 0 throws. Parent repeats: 21939.0, 23125.3, 22528.3, 21858.1, 23114.0 ms. Settled tail (last 3): 22528.3, 21858.1, 23114.0 ms (mean 22500.1). New repeats: 18638.2, 18867.9, 18633.6, 18597.6, 18479.1 ms. Settled tail: 18633.6, 18597.6, 18479.1 ms (mean 18570.1). The bad corpus got faster, by about 3.9 s on the tail means. The gain is modest because the 24 `xsd-error`-only files still pay for XSD and dominate the corpus.

Clean file `tests/fixtures/working/word-tolerated-misplaced-pgsz.docx`, valid on every repeat, XSD started, 8 parts sent, 8 eligible. Parent repeats: 20.2, 26.9, 18.8, 18.1, 17.3, 33.2, 17.8, 18.4 ms (minimum 17.3). Settled tail: 33.2, 17.8, 18.4 ms (mean 23.1; the 33.2 is a spike inside the tail). New repeats: 12.8, 12.7, 15.5, 13.3, 13.1, 13.6, 13.5, 11.7 ms (maximum 15.5). Settled tail: 13.6, 13.5, 11.7 ms (mean 12.9). The clean file got faster. Every new repeat is below every parent repeat, so the speedup is not the 33.2 spike. The file still pays for a full XSD walk. Its non-blocking `pgSz` errors are validated twice (status probe, then message collection) and the file was still faster.

Separate engine probe, outside the package, on the same raw parts from that clean file (7 parts, `docProps/core.xml` excluded) and from `header-with-rid-no-sidecar.docx` (12 parts, `core.xml` excluded). One reused native libxml2 2.9.13 validation context, no error handler, versus libxml2-wasm `XsdValidator.validate` (new context, structured errors). 8 repeats. Sum of per-schema settled-tail means: native 1.2 ms, wasm 1.4 ms. Native was faster. The visible gap is header `wml.xsd` (native tail 0.5, 0.5, 0.5 ms vs wasm 0.7, 0.7, 0.7 ms). Several other groups round to the same tenth and the script's unrounded comparison flips either way. `wml.xsd` compile on the clean file was 32.6 ms native and 49.9 ms wasm; that compile is outside the validate tail. These totals are about a millisecond. They do not explain the product's ~13 ms clean file or the ~18 s corpus. The native context has no error handler, so libxml2 wrote schema errors to stderr; that context is not the word-valid verdict, because a status with no message cannot apply `isWordBlockingXsdIssue` or the ignore list. The Rust `libxml` crate was not used.
