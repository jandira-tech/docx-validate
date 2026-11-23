# Word-fatal shapes found in Apryse's redlines (2026-11-01)

Status: open, to add to the validator later. Found while benchmarking Apryse
`DOCXCompare` (apryse-sdk 12.2.1) in neurotic_docx_bench (#100 there). The same
list is in `../jubarte-redlines/docs/TODO.md` §15, for `jubarte validate`.

## Evidence

Word 16.115, driven by the bench's `scripts/word_pdf.py`, one batch alone, so no
other document could poison an open. Word refused 33 Apryse redlines: 18 of a
random 50, plus 15 from an earlier stage with Apryse's watermark removed. The
watermark is not the cause: those 15 fail with it and without it. Each file was
repaired two ways and reopened:

| Repair | Opens in Word |
|---|---|
| hyperlink shape only (shape 1) | 17 / 33 |
| shapes 1 to 4, plus comment ranges lifted out of runs and fractional twips rounded | 29 / 33 |

Comparing the 33 refused files with the 32 that open (Open XML SDK 3.5.1 with no
error cap, plus lxml), these four shapes appear only in refused files, and every
refused file has at least one:

| # | Shape | Refused / open | Status |
|---|---|---|---|
| 1 | `w:hyperlink` as a child of `w:r`, the run inside `w:del`: `<w:del><w:r><w:hyperlink r:id=…><w:r><w:t>…` | 24 / 0 | **Causal**: fixing it alone opened 17 of 33. Inside `w:ins` Word tolerates it (9 refused / 6 open), so flag the `w:del` case as Word-fatal and the `w:ins` case as a schema error only. Word's own shape: `<w:hyperlink><w:del><w:r><w:delText>…` |
| 2 | `w:fldSimple` as a child of `w:r` inside `w:del`/`w:ins` | 2 / 0 | needs a single-defect Word probe |
| 3 | `w:footnoteReference` / `w:endnoteReference` whose id has no note; `w:commentReference` in a package with no `comments.xml` | 12 / 0 | the dangling-footnote diagnostic exists (`docx-diagnostics.ts`); check it fires on these files and covers the comment case |
| 4 | two `w:bookmarkStart` with the same `w:id` in one part | 7 / 0 | needs a single-defect Word probe |

The XSD pass may already reject shapes 1 and 2 as unexpected children of `w:r`.
What's missing is the Word-fatal classification and an issue code naming the
shape, so a caller can tell it apart from the tolerated schema noise below.

## Not Word-fatal (measured, do not flag as fatal)

- `w:pPr` children out of schema order (e.g. `w:rPr` before `w:pStyle`):
  present in 27 of the 32 Apryse files that open. The SDK reports it as
  "unexpected child element 'w:pStyle'", hundreds of times per file.
- `w:hyperlink` inside a run inside `w:ins` (above).

## Still unexplained

Four files still fail after the full repair: three crash Word
(`87dc2ebfec…`, `ec3dd0df69_file_53…`, `fc29454fff…`) and one does not load
(`8441c4ae21_sd_2517_localized_heading_styles…`). The SDK shows no defect that
these four share and the 29 repaired-and-opening files lack. Bisecting them in
Word is the next step. An earlier half-split of `8441c4ae21` put the failure in
the first half, which holds a deleted TOC with 133 instances of shape 1.

## Specimens

These live in git-ignored scratch in the bench checkout; copy them into
`tests/fixtures/broken/` (named `<subject>.<error>.docx`) and
`word-invalid-fixtures/` before relying on them:

- originals: `~/temp/T/neurotic_docx_bench/out/apryse_fixab/orig/`
- repaired arms: `…/out/apryse_fixab/hyperlink/`, `…/out/apryse_fixab/all/`
- Word's verdict per file: `…/out/apryse_fixab/word/pdf/manifest.jsonl`
- repair script (diagnostic only): `~/temp/T/ndb-apryse/out/apryse/diag_fix.py`
- SDK validator output: `…/out/apryse_fixab/oxval_{orig,hyperlink,all}.jsonl`
