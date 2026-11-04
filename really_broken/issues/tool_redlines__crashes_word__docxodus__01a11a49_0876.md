# [docxodus 0876] failed: w:pgMar in w:sectPrChange omits required w:gutter (site unverifiable)

## Summary

`tool_redlines/crashes_word/docxodus__01a11a49_0876.docx` was a docxodus-produced tracked-changes
redline (change records authored `Open-Xml-PowerTools`, `w:date 2000-01-01T00:00:00Z`) whose
`word/document.xml` — per the analyst's inspection, no longer re-checkable — carried two `w:pgMar`
elements: the live one at `document.xml:3559` with `w:gutter="0"`, and the tracked-change snapshot
inside `w:sectPrChange` whose `w:pgMar` omitted `w:gutter`. That omission is the package's single
XSD defect.

**This issue is disputed for provenance, not for the rule.** The schema rule is real and was
re-verified first-hand against the bundled XSDs (quoted below). The offending *site* cannot be
re-verified: the specimen is gone from the repo, exactly as for its sibling 0888
(`really_broken/issues/crashes_word__docxodus__01a11a49_0888.md`, read this session):

- `really_broken/tool_redlines/` contains only `repair_prompt/` — no `crashes_word/` directory;
- repo-wide `find . -name "*0876*" -not -path "./node_modules/*" -not -path "./.git/*"` returns
  nothing; the unpacked copy `really_broken/unpacked/tool_redlines__crashes_word__docxodus__01a11a49_0876`
  does not exist;
- `really_broken/manifest.csv` does not exist (no `*.csv` in `really_broken/` at all);
- the analyst's own validate command now errors (run this session, exit 1):
  `pnpm exec tsx src/scripts/office/validate.ts really_broken/tool_redlines/crashes_word/docxodus__01a11a49_0876.docx`
  → `Error: … does not exist`.

## Word behaviour

- **Live result (this ask, ask-given):** never opened — outcome `failed`; the live agent found the
  same absences reported above. No Word observation was or can be made until the specimen is restored.
- **Manifest row (ask-given, analyst-quoted before deletion; not verifiable now):**
  `really_broken/manifest.csv` row 29 — `d1_failed_in: soffice`, `local_word_error` "Word crashed on
  it in 2 passes: Microsoft Word got an error: Connection is invalid." That string is the
  AppleScript error text for a lost application connection (Word process dropped mid-open); it says
  nothing about the document, and the direct inspection below superseded it.
- **Prior proven assessment (read this session, the established Word cause):**
  `docs/really-broken/assessments.jsonl:16`, mirrored at `docs/really-broken/inspection-log.jsonl:16`
  — `status: proven`, `defects: []`, `defect: "None. Microsoft Word opens the original file with no
  warning, error, or offer to repair."`, `change: none`, `word.original: clean-open`. So the missing
  `w:gutter` is **not Word-blocking**: Word opens these exact bytes clean. It credits no
  docx-validate check because the assessment records no *Word* defect; the XSD defect is separate
  and was found by the validator run, not by Word.
- **Cohort context (counted this session):** 16 `tool_redlines/crashes_word/` rows in
  `assessments.jsonl`; 8 are proven clean-open no-defect files (docxodus 0876/0888/1860/1862 and
  jubarte 0876/1360/1862/1990), 8 carry real defects (superdoc-redlines 0589/0813/1812/1993/2136/2213,
  docxodus 2183/2243). This file sits in the mislabeled half; the misclassification is systematic.

## Violation 1 — `xsd-error`: `w:pgMar` inside `w:sectPrChange` omits required `w:gutter`

**Offending XML** (analyst-quoted, ask-given; **not re-read this session — the part is absent**),
`word/document.xml`:

```xml
<w:sectPrChange w:author="Open-Xml-PowerTools" w:id="374" w:date="2000-01-01T00:00:00Z">
  <w:sectPr>
    <w:pgSz w:h="15840" w:w="12240" w:orient="portrait"/>
    <w:pgMar w:bottom="720" w:top="720" w:left="720" w:right="720" w:header="720" w:footer="576"/>
    <w:pgNumType w:start="1"/>
    <w:titlePg w:val="1"/>
  </w:sectPr>
</w:sectPrChange>
```

**Schema rule** (re-verified this session, `src/scripts/office/schemas/ISO-IEC29500-4_2016/wml.xsd`).
The embedded `w:sectPr` is inside the rule's scope — the chain is complete in the one file:

- `wml.xsd:965` (in `CT_SectPrChange`, lines 961–969): `<xsd:element name="sectPr" type="CT_SectPrBase" minOccurs="0"/>`
- `wml.xsd:1640` (in `CT_SectPrBase`, lines 1638–1643): `<xsd:group ref="EG_SectPrContents" minOccurs="0"/>`
- `wml.xsd:1615` (in `EG_SectPrContents`, lines 1609+): `<xsd:element name="pgMar" type="CT_PageMar" minOccurs="0"/>`
- `wml.xsd:1458` (in `CT_PageMar`, lines 1451–1459):

```xml
<xsd:complexType name="CT_PageMar">
  <xsd:attribute name="top" type="ST_SignedTwipsMeasure" use="required"/>
  <xsd:attribute name="right" type="s:ST_TwipsMeasure" use="required"/>
  <xsd:attribute name="bottom" type="ST_SignedTwipsMeasure" use="required"/>
  <xsd:attribute name="left" type="s:ST_TwipsMeasure" use="required"/>
  <xsd:attribute name="header" type="s:ST_TwipsMeasure" use="required"/>
  <xsd:attribute name="footer" type="s:ST_TwipsMeasure" use="required"/>
  <xsd:attribute name="gutter" type="s:ST_TwipsMeasure" use="required"/>   <!-- line 1458 -->
</xsd:complexType>
```

The ISO Strict twin (`src/scripts/office/schemas/iso-strict/wml.xsd:1373-1381`) requires `gutter`
identically. A `w:pgMar` with all six other attributes but no `w:gutter` violates the schema
wherever it sits — the `w:sectPr` inside `w:sectPrChange` is a previous-section snapshot typed
`CT_SectPrBase`, and its `w:pgMar` is the same `CT_PageMar`.

**Proposed fix:** docxodus must emit the full required attribute set on every `w:pgMar` it writes,
including the previous-section snapshot inside `w:sectPrChange` — copy the source section's
`w:gutter` verbatim (`w:gutter="0"` would match the sibling live `w:pgMar` at `document.xml:3559`).

**Fix target:** `docxodus` (generator bug), conditional on the site being confirmed once the
specimen is restored.

The analyst's other xmllint output — 29 `w14:paraId` "attribute not allowed" sites — is
Microsoft-extension noise the lane excludes by rule, and this repo's validator already handles it:
`_preprocessXmlForXsd` (`src/scripts/office/validators/base.ts:2287`) applies markup compatibility
with understood namespaces before the XSD walk, which is why the analyst's validator run reported
only the gutter defect while raw xmllint showed the w14 noise too.

## Where the fixes belong and why

- **docxodus — the generator.** If the site is confirmed, the redline engine (change records
  authored `Open-Xml-PowerTools`) snapshots the previous `w:sectPr` into `w:sectPrChange` but drops
  `w:gutter` from the embedded `w:pgMar`. A snapshot that is schema-invalid cannot round-trip
  through strict validators even though Word tolerates it, so the generator must copy all seven
  required attributes.
- **really-broken lane harness (d1 Word-driving stage + manifest writer) — the label defect.**
  A stage whose AppleScript connection dropped recorded `crashes_word` for a document the lane's own
  proven assessment says Word opens clean. Fix: classify transport/automation errors
  (`Connection is invalid`, AppleScript timeouts, app-quit mid-run) as `unknown` with a retry,
  never as a document crash; 8 of the 16 `crashes_word` rows are proven clean-opens — the
  regression corpus already exists.
- **Not docx-validate.** No check is missing and none misreports: the lenient XSD walk's
  `xsd-error` on the missing `w:gutter` is correct detection of a genuine `wml.xsd` violation, and
  word-valid correctly does not block on it (DOCX/DOCM word-valid runs no XSD, and the proven
  Word observation agrees).
- **Not repair.** The document opens clean in Word; there is no broken source document to
  prescribe a repair for.

## Verification

Checks run this session: read `assessments.jsonl:16` and `inspection-log.jsonl:16` (both re-read in
full via grep); counted the 16-row cohort (8 clean-open / 8 defective); repo-wide `find` for
`*0876*`; `ls really_broken/` (no `manifest.csv`, no `crashes_word/`); the validator command above
(exit 1, `does not exist`); grep of `wml.xsd` / `iso-strict/wml.xsd` for `CT_PageMar`,
`CT_SectPrChange`, `CT_SectPrBase`, `EG_SectPrContents`. Not run: any check touching the
document's XML (validate full run, xmmlint) — the specimen is absent; live Word open — same.

1. **Restore the specimen** from the corpus that produced the assessments, then re-run
   `pnpm exec tsx src/scripts/office/validate.ts really_broken/tool_redlines/crashes_word/docxodus__01a11a49_0876.docx`
   — expect exit 1 with exactly the one `xsd-error` on `word/document.xml` (the analyst's result).
   Minimal-change experiment: add `w:gutter="0"` to the `w:pgMar` inside `w:sectPrChange` and
   nothing else — the `xsd-error` must disappear. That confirms both the single-defect claim and
   this issue; if it does not, close this issue as not reproduced.
2. **For docxodus:** a generator unit test that redlines a document whose sections differ in
   margins and asserts every emitted `w:sectPrChange//w:pgMar` carries all seven required
   attributes (`top right bottom left header footer gutter`).
3. **For the harness:** a unit test feeding the d1 stage the recorded
   `Microsoft Word got an error: Connection is invalid` result and asserting the recorded outcome
   is `unknown` (retryable), not a crash category; then re-run the cohort's 8 clean-open files
   through the stage and assert none lands in `crashes_word`. A two-run Word open of the restored
   specimen should confirm `clean-open`, matching `assessments.jsonl:16`.
