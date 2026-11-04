# [jubarte 0876] sectPrChange's embedded pgMar omits the schema-required w:gutter

## Summary

`tool_redlines/crashes_word/jubarte__01a11a49_0876.docx` is a jubarte-produced redline
(tracked changes authored `w:author="Redline"`, 15 XML parts under `word/`). The whole
package carries exactly one substantive schema violation: inside jubarte's own
`w:sectPrChange` revision record, the embedded pre-change `w:sectPr`'s `w:pgMar` omits
the required `w:gutter` attribute. The document's live `w:sectPr` `w:pgMar`
(document.xml:3004) does carry `w:gutter="0"` — jubarte dropped the attribute only when
synthesizing the revision record.

Despite the lane's `crashes_word` category, this defect is **not** Word-blocking: the
proven assessment records a clean open, and this repo's word-valid profile (the CLI
default for `.docx`) passes the file. The "crash" in the manifest is an automation-level
AppleEvent failure, not a document verdict.

Path note: `really_broken/tool_redlines/crashes_word/jubarte__01a11a49_0876.docx` and
`really_broken/unpacked/analyst-tool_redlines__crashes_word__jubarte__01a11a49_0876` no
longer exist in the working tree, so the live Word check could not open anything. All
re-verification below ran against the corpus copy at
`/Users/arthrod/temp/T/neurotic_docx_bench/maybe_future_corpus/really_broken/…`, whose
sha256 `9c7e3bebb1dd5ea1711d9ee1e79fac890a4057bc510fdb5ec236f00f8cb8acea` matches the
manifest row exactly (verified with `shasum -a 256` this session).

## Word behaviour

- **Live Word result (this session):** never ran — the named file is absent from
  `really_broken/` (only `repair_prompt/` exists; a repo-wide `find *_0876*` returns
  nothing). `really_broken/manifest.csv` likewise does not exist in the repo; the row
  lives in the corpus tree copy.
- **Manifest row 96** (corpus `really_broken/manifest.csv`): category `crashes_word`,
  tool `jubarte`, `d1_failed_in` `soffice`, `local_word_error` "Word crashed on it in
  2 passes: Microsoft Word got an error: Connection is invalid." — that string is the
  macOS AppleEvent error (-609, "Connection is invalid") from the automation harness,
  not a Word document-open verdict.
- **Proven assessment** (`docs/really-broken/assessments.jsonl:20`): status `proven`,
  defect "None. Microsoft Word opens the original file with no warning, error, or offer
  to repair.", change `none`, `word.original` `clean-open`. This overrides the crash
  label: real Word opens these bytes cleanly.

## Violation 1 — `xsd-error`: `w:pgMar` in `w:sectPrChange` missing `w:gutter`

**Offending XML** — `word/document.xml:3009-3015` (unpacked view; end of `w:body`):

```xml
    <w:sectPr>
      …
      <w:pgMar w:top="1440" w:right="1008" w:bottom="1440" w:left="1872"
               w:header="1440" w:footer="792" w:gutter="0"/>          <!-- line 3004: live, valid -->
      …
      <w:sectPrChange w:author="Redline" w:date="1970-01-01T00:00:00Z" w:id="193">
        <w:sectPr>
          <w:pgSz w:h="15840" w:w="12240" w:orient="portrait"/>
          <w:pgMar w:bottom="720" w:top="720" w:left="720" w:right="720"
                   w:header="720" w:footer="576"/>                     <!-- line 3012: no w:gutter -->
          <w:titlePg w:val="1"/>
        </w:sectPr>
      </w:sectPrChange>
    </w:sectPr>
```

**Schema rule** — `src/scripts/office/schemas/ISO-IEC29500-4_2016/wml.xsd:1451-1459`
(all seven attributes are `use="required"`):

```xsd
<xsd:complexType name="CT_PageMar">
  <xsd:attribute name="top" type="ST_SignedTwipsMeasure" use="required"/>
  <xsd:attribute name="right" type="s:ST_TwipsMeasure" use="required"/>
  <xsd:attribute name="bottom" type="ST_SignedTwipsMeasure" use="required"/>
  <xsd:attribute name="left" type="s:ST_TwipsMeasure" use="required"/>
  <xsd:attribute name="header" type="s:ST_TwipsMeasure" use="required"/>
  <xsd:attribute name="footer" type="s:ST_TwipsMeasure" use="required"/>
  <xsd:attribute name="gutter" type="s:ST_TwipsMeasure" use="required"/>
</xsd:complexType>
```

The embedded `sectPr` is typed `CT_SectPrBase` via `CT_SectPrChange`
(`wml.xsd:961-967`, `<xsd:element name="sectPr" type="CT_SectPrBase" minOccurs="0"/>`),
so the revision record's `w:pgMar` is judged by the same `CT_PageMar`.

**Reproduction (all run this session):**
- `xmllint --noout --schema src/scripts/office/schemas/ISO-IEC29500-4_2016/wml.xsd
  <corpus-unpacked>/word/document.xml` → exit 3; the only missing-attribute error is
  `document.xml:3012: element pgMar: … The attribute '…main}gutter' is required but
  missing.` The other 28 errors are `w14:paraId` "not allowed" noise (raw xmllint is
  MCE-blind); `header1.xml` raw xmllint is noise-only. Package-wide there are exactly
  two `w:pgMar` elements; exactly one lacks `w:gutter`.
- `pnpm exec tsx src/scripts/office/validate.ts --profile lenient <docx>` → exit 1,
  exactly one error: `Element '{…main}pgMar': The attribute '{…main}gutter' is required
  but missing.` (issue code `xsd-error`, assigned at
  `src/scripts/office/validators/base.ts:1900`; strict path at `base.ts:1890`).
- `--profile strict` → exit 1, the same `xsd-error` (line marked as the preprocessed
  validation view) plus strict-only `w14:paraId`/`textId` completeness findings.
- `--profile word-valid` (the CLI default for `.docx`,
  `src/scripts/office/validate.ts:167,1600`) → exit 0, "All validations PASSED!".

**Proposed fix:** in jubarte's revision-record synthesis, serialize the pre-change
`w:sectPr` with a complete `CT_PageMar` — carry the source document's `w:gutter` over,
or emit `w:gutter="0"` when reconstructing from a parsed margin model that lost it.

**Fix target:** `jubarte` (generator bug — the redline's own XML is schema-invalid).

Corrections to the analyst report, verified against the current tree: there is no
`isWordBlockingXsdIssue` at `validators/docx.ts:4228-4257` (docx.ts is 3747 lines; the
name survives only in a stale comment at `src/lib/xsd-validator.ts:325`). The actual
reason word-valid passes is that DOCX word-valid runs no XSD walk at all —
`profileWalksSchemas` (`src/lib/mode-policy.ts:142-143`) is true for DOCX only under
lenient/strict. Also, the analyst's plain (no-flag) run "exit 1" does not reproduce
today: the plain run defaults to word-valid and exits 0; the exit-1 single-pgMar-error
result reproduces exactly under `--profile lenient`.

## Where the fixes belong and why

- **jubarte** — the only genuine defect. The invalid XML is jubarte's synthesized
  `w:sectPrChange` content (its `w:author="Redline"` stamp), not source-document
  content: the live sectPr beside it is complete. The generator must keep all seven
  required `CT_PageMar` attributes when embedding the pre-change `w:sectPr`.
- **Not docx-validate** — the validator already reports the defect precisely: one
  `xsd-error` under lenient (`base.ts:1900`) and strict (`base.ts:1890`), correctly
  invisible under word-valid, which models Word openability and matches the proven
  clean open. No missing check, no false positive.
- **Not repair** — the source documents that fed this redline are not the broken input;
  the redline generator itself produced the invalid attribute set.

## Verification

1. **Minimal-change experiment (run this session, proves the single defect):** copying
   the docx and adding only `w:gutter="0"` to the line-3012 `w:pgMar` (a one-attribute
   patch, everything else byte-identical) makes
   `pnpm exec tsx src/scripts/office/validate.ts --profile lenient` exit 0 with
   "All validations PASSED!" (before: exit 1, the single gutter error).
2. **Regression fixture for docx-validate:** a `broken/` specimen named
   `sectprchange-pgmar-missing-gutter.docx` (this document) asserting the lenient run
   yields exactly one `xsd-error` naming `gutter` on `word/document.xml`, and the
   word-valid run passes — pinning both the detection and the non-blocking verdict.
3. **jubarte-side test:** synthesize a sectPr-change redline from a document whose
   `w:pgMar` carries `w:gutter` and assert the embedded `w:pgMar` retains it;
   re-run this repo's lenient profile on the regenerated redline (expected: pass).
