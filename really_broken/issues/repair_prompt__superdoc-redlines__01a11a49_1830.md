# [superdoc-redlines 1830] repair-prompt: empty `w:id` value on the TOC `w:sdt`

## Summary

`really_broken/tool_redlines/repair_prompt/superdoc-redlines__01a11a49_1830.docx` (sha256
`6db54e9f4fb3653525c00db038dbefe8c08f4376babb6398ca57c6f6dddf7c5f`) is a redline produced by
**superdoc-redlines**: `word/document.xml` carries 48 `w:ins` and 92 `w:del` tracked changes around a
Table-of-Contents structured document tag. That single `w:sdt` — the only one in the package —
declares its id as `<w:id w:val=""/>`, and the empty string is not a lexical `xsd:integer`, so the
`ST_DecimalNumber`-typed attribute is schema-invalid and **Word cannot open the file at all**.

- **Live Word outcome:** refused with no recovery offer — the only dialog was *"Word experienced an
  error trying to open the file. Try these suggestions. … (superdoc-redlines__01a11a49_1830.docx)"*
  (OK only); no windows remained, no crash report. (The probe's `outcome` field spells this
  `repair-prompt`; the dialog text and the proven assessment below record it as `open-error`.)
- **Manifest row:** `really_broken/manifest.csv` does not exist in this repo (searched the whole tree
  excluding `node_modules`; the only `manifest.csv` is the jubarte lane's at
  `tests/lane/jubarte-0.13.0/manifest.csv`), so no row could be read; the live probe's `manifestSays`
  is likewise empty.

This document is also the repo's own calibration: it is byte-identical to the shipped fixture
`tests/fixtures/broken/package.word-sdt-id-empty.docx` (same sha256, verified this run), which is the
fixture the `word-sdt-id-empty` check cites in its header.

## Word behaviour

- **Live result (this run):** refused as above; Word quit cleanly; `agrees: true` with the recorded
  assessment.
- **Prior assessment — PROVEN** (`docs/really-broken/assessments.jsonl:125` and
  `docs/really-broken/inspection-log.jsonl:125`, source `word-ab`):
  - original bytes → `open-error`;
  - **set the empty `w:sdtPr/w:id` `w:val` to `1`** → **`clean-open`**, with the eight nil-GUID
    embedded font keys (`w:fontKey="{00000000-0000-0000-0000-000000000000}"`) and everything else
    untouched;
  - probe-time `validate()` was valid on the original only because the check did not exist yet; the
    shipped `word-sdt-id-empty` reports the original and not the repair.
- **Hash-bound manifest evidence** (`tests/fixtures-all.manifest.json`): the byte-identical broken
  fixture carries `word: { outcome: "open-error", sha256: 6db54e9f…, wordVersion: 16.115 }`, and
  `tests/fixtures/working/package.word-sdt-id-empty-repaired.docx` (sha256 `57b25983…`, whose only
  delta is `<w:id w:val="1"/>`, verified this run) carries `outcome: "clean-open"`.

## Violation — `word-sdt-id-empty` / `sdt-empty-id-val`

**Part:** `word/document.xml` — 1 site (line 629), re-verified in the unpacked package at
`really_broken/unpacked/tool_redlines__repair_prompt__superdoc-redlines__01a11a49_1830/` (the ask's
name without the `tool_redlines__` prefix does not exist on disk). `grep -c '<w:sdt>'` = 1 in
`document.xml` and 0 in `header1/2.xml`, `footer1/2.xml`, `footnotes.xml`, so this site is the whole
defect.

**Offending XML** (`word/document.xml:627-634`):

```xml
<w:sdt>
  <w:sdtPr>
    <w:id w:val=""/>                     <!-- line 629: empty string, not an integer -->
    <w:docPartObj>
      <w:docPartGallery w:val="Table of Contents"/>
      <w:docPartUnique/>
    </w:docPartObj>
  </w:sdtPr>
```

**Schema rule** (ISO/IEC 29500-4:2016, bundled `src/scripts/office/schemas/ISO-IEC29500-4_2016/wml.xsd`,
lines read directly):

```xml
<!-- wml.xsd:56-58 -->
<xsd:simpleType name="ST_DecimalNumber">
  <xsd:restriction base="xsd:integer"/>
</xsd:simpleType>
<!-- wml.xsd:59-61 -->
<xsd:complexType name="CT_DecimalNumber">
  <xsd:attribute name="val" type="ST_DecimalNumber" use="required"/>
</xsd:complexType>
<!-- wml.xsd:2053, inside CT_SdtPr's sequence (declared wml.xsd:2050) -->
<xsd:element name="id" type="CT_DecimalNumber" minOccurs="0"/>
```

`w:sdtPr/w:id` is XSD-optional, but when present its `w:val` must be an integer; `""` is not a lexical
`xsd:integer`. Raw `xmllint --noout --schema …/wml.xsd word/document.xml` (run this session) reports at
line 629: `element id … attribute 'val': '' is not a valid value of the atomic type
'ST_DecimalNumber'`. The other 911 xmllint errors are Microsoft-extension noise — undeclared
`w14:paraId`, `w:authorEmail`, `w:rsid*` attributes — plus 140 bare-GUID revision ids
(`w:id="96d9ca50-…"` on 48 `w:ins` + 92 `w:del`, counted this run) that lexically violate
`ST_DecimalNumber` too but are Word-tolerated: the proven repair left every one of them in place and
still opened clean, and the repo has no check for them, so they are context here, not violations.

**Proposed fix — in superdoc-redlines' sdt emitter:** give every emitted `w:sdtPr/w:id` a real
integer. The minimal proven repair for this file is `w:val=""` → `w:val="1"`; a generator should emit
a unique decimal id per `w:sdt` (Word's own files use large 32-bit integers). Never serialize the id
element with an empty value — the same rule the repo's check enforces.

**Fix target:** `superdoc-redlines` (generator bug — its output is schema-invalid and Word refuses it).

## Where the fixes belong and why

- **superdoc-redlines** — the redline's own `word/document.xml` is genuinely invalid: the TOC content
  control is stamped with an empty id. Whatever code path writes the `w:sdtPr` (likely a TOC template
  with an unbound id placeholder) must emit a concrete integer or omit `w:id` entirely (it is
  `minOccurs="0"`). Until then every document it emits with this TOC block fails to open in Word.
- **docx-validate — no change.** `word-sdt-id-empty` already exists
  (`src/scripts/office/checks/word-sdt-id.ts:19`, header comment citing "fixture 1830" — this very
  document), is registered for full `word-valid`/`strict` DOCX runs, and reports exactly this one
  site. Neither a validator gap nor a false positive; it is the credited detection.
- **Not a `repair` case:** the broken artefact is the tool's output, not a hand-authored source
  document, so no repair prescription is filed for the source.

## Verification

- **Validator reproduction (this run):**
  `pnpm exec tsx src/scripts/office/validate.ts really_broken/tool_redlines/repair_prompt/superdoc-redlines__01a11a49_1830.docx`
  → exit 1, exactly one ERROR line: `word-sdt-id-empty` on `word/document.xml`.
- **Sites and schema re-verified:** `word/document.xml:629` read directly (snippet above); XSD lines
  56-61 and 2050/2053 read directly (quoted above); sdt counts confirmed across all story parts;
  `xmllint` error at line 629 captured verbatim.
- **Established Word cause (proven assessment, carried in):** setting that single `w:id` from
  `w:val=""` to `w:val="1"` opens clean while the 140 GUID revision ids and eight nil font keys stay —
  `docs/really-broken/inspection-log.jsonl:125`. Both sides are pinned hash-bound in
  `tests/fixtures-all.manifest.json` (`open-error` / `clean-open`, Word 16.115), and the repaired
  fixture's only delta was verified this run to be `<w:id w:val="1"/>`.
- **For the generator fix:** re-emit the document with an integer sdt id and confirm a clean open
  under the repo's two-run Word-observation standard (two separate runs, verdicts read from each run's
  Word log); `tests/word-sdt-id-openability.test.ts` and `tests/checks-word-sdt-id.test.ts` already
  pin the check's behaviour on both fixtures.
