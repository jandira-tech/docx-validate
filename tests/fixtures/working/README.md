# Working DOCX fixtures

Counterpart to [`../broken/`](../broken/README.md). Each file here MUST
validate clean under both `strict` and `lenient` profiles, and is pinned
that way in `tests/fixtures-all.manifest.json`. Used as a regression
floor — if any of these starts failing, the validator has gotten stricter
in a way that's not aligned with real-world Word output.

| File | Source | Why it's golden |
| --- | --- | --- |
| `sample-document.afterword-repaired.docx` | Real `Sample Document.docx` (a Plate-pipeline export) opened in Word, then saved. | Word's writer pipeline is treated as the canonical "fully-repaired" reference for the matching `broken/sample-document.*.docx` chain. |
| `sample-document.really-repaired.docx`    | Same source, second Word save. | Pinned alongside the first to catch any non-determinism in our validator's response to two functionally-equivalent Word outputs. |
| `sample-document.our-repaired.docx`       | Output of our own `repair()` pipeline applied to `broken/sample-document.id-overflow.docx`. | Pins the end-state of our auto-repair on the most-defective member of the broken chain. If we regress a repair pass, this fixture will start failing. |
| `sample-document.word-repair-of-our-output.docx`       | Word's save of `broken/sample-document.our-pipeline-output-tr-no-textid.docx` (i.e. Word repairing OUR pre-textId-fix output). | Captures what Word does to fix the asymmetric `<w:tr>` paraId-without-textId shape — adds `w14:textId="77777777"` on every row, strips the malformed `commentsExtensible.xml`, and otherwise leaves the structure intact. Reference for what a "Word will open this cleanly" output looks like. |
| `sample-document.word-repair-of-our-output-iter2.docx` | Second Word save of the same input. | Pinned alongside iter1 to catch non-determinism between two Word saves of the same input. |
