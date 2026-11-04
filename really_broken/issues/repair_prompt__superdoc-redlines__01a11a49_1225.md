# [superdoc-redlines 1225] repair-prompt: cell-less table row Word refuses to open

## Summary

`tool_redlines/repair_prompt/superdoc-redlines__01a11a49_1225.docx` is a superdoc-redlines
tracked-changes redline (sha256 `3745e490…f3a51f6`, verified identical here and in the sibling
corpus copy). All three reported violations reproduce against the unpacked package
(`really_broken/unpacked/tool_redlines__repair_prompt__superdoc-redlines__01a11a49_1225/`) and the
bundled `ISO-IEC29500-4_2016/wml.xsd`; every defect lives in `word/document.xml`.

**Word refuses the file without any recovery offer** — the live run's only dialog was
`Word experienced an error trying to open the file. Try these suggestions. …
(superdoc-redlines__01a11a49_1225.docx)`, OK only, no document afterwards, no crash report. By
this repo's rubric (AGENTS.md, *Real-Word probes*) that dialog is an **open error**, even though
the live result arrived labelled `repair-prompt` (the lane name). The manifest agrees on refusal.
Provenance: the repo's `really_broken/` has no `manifest.csv`; the row is quoted from the sibling
corpus `/Users/arthrod/temp/T/neurotic_docx_bench/maybe_future_corpus/really_broken/manifest.csv`.

## Word behaviour

- **Live (this ask):** open-error dialog quoted above; after OK no windows or documents remained;
  no crash report; Word quit cleanly.
- **Manifest row (sibling corpus):** `word:convert_failed` — "Word could not load it in 2 passes:
  document loaded empty (Word could not read it)".
- **Prior assessment — proven** (`docs/really-broken/assessments.jsonl` + `inspection-log.jsonl`,
  path `tool_redlines/repair_prompt/superdoc-redlines__01a11a49_1225.docx`, source `word-ab`):
  original bytes → Word `open-error`, validator invalid with `word-row-without-cell`; after
  inserting `<w:tc><w:p/></w:tc>` before `</w:tr>` on row `w14:paraId="54E79239"` → Word
  **clean-open**, `validate()` valid, `w:gridBefore w:val="2"` kept. That repair left both `tcPr`
  ordering defects untouched, so it doubles as direct evidence Word **tolerates** violations 2
  and 3 while refusing violation 1.

## Violation 1 — `word-row-without-cell` (error, the Word-blocking defect)

`word/document.xml:198-204` — a `w:tr` whose only child is `w:trPr`; no `w:tc`:

```xml
<w:tr w:rsidR="00464A37" w:rsidRPr="00FB3B7D" w:rsidTr="000D5DEC" w14:paraId="54E79239" w14:textId="77777777">
  <w:trPr>
    <w:hidden w:val="0"/>
    <w:trHeight w:val="1095"/>
    <w:gridBefore w:val="2"/>
  </w:trPr>
</w:tr>
```

**Rule.** `CT_Row` (`src/scripts/office/schemas/ISO-IEC29500-4_2016/wml.xsd:2417-2424`):

```xsd
<xsd:complexType name="CT_Row">
  <xsd:sequence>
    <xsd:element name="tblPrEx" type="CT_TblPrEx" minOccurs="0" maxOccurs="1"/>
    <xsd:element name="trPr" type="CT_TrPr" minOccurs="0" maxOccurs="1"/>
    <xsd:group ref="EG_ContentCellContent" minOccurs="0" maxOccurs="unbounded"/>
  </xsd:sequence>
```

`EG_ContentCellContent` (wml.xsd:2135) has `minOccurs="0"`, so **the XSD permits a cell-less
row** — raw xmllint prints nothing on this row (verified this session). The one-cell-per-row
requirement is Word's own, carried by the registered check `word-row-without-cell`
(`src/scripts/office/checks/tables.ts:20`), which reported exactly this one site.

**Fix — target `superdoc-redlines`.** The generator emitted a `gridBefore` spacer row (2 leading
grid columns, hidden, fixed height) with no cell. Every emitted `w:tr` must carry at least one
`w:tc` — Word's own serialisation gives such a row one empty cell after the skipped grid slots,
exactly the proven minimal repair.

## Violation 2 — `word-cell-margin-order` (warning; schema-invalid, Word tolerates)

`word/document.xml:244-249` — inside `w:tcMar` the margins are serialized `left, right, top,
bottom`:

```xml
<w:tcMar>
  <w:left w:w="0" w:type="dxa"/>
  <w:right w:w="0" w:type="dxa"/>
  <w:top w:w="0" w:type="dxa"/>      <!-- :247 -->
  <w:bottom w:w="0" w:type="dxa"/>
</w:tcMar>
```

**Rule.** `CT_TcMar` (`wml.xsd:2299-2307`) is an `xsd:sequence`:

```xsd
<xsd:complexType name="CT_TcMar">
  <xsd:sequence>
    <xsd:element name="top" type="CT_TblWidth" minOccurs="0" maxOccurs="1"/>
    <xsd:element name="start" type="CT_TblWidth" minOccurs="0" maxOccurs="1"/>
    <xsd:element name="left" type="CT_TblWidth" minOccurs="0"/>
    <xsd:element name="bottom" type="CT_TblWidth" minOccurs="0" maxOccurs="1"/>
    <xsd:element name="end" type="CT_TblWidth" minOccurs="0" maxOccurs="1"/>
    <xsd:element name="right" type="CT_TblWidth" minOccurs="0"/>
  </xsd:sequence>
</xsd:complexType>
```

`w:top` and `w:bottom` follow `w:right`; raw xmllint confirms `document.xml:247: element top: This
element is not expected`. The registered check `word-cell-margin-order`
(`src/scripts/office/checks/cell-margin-order.ts:107`) reports both out-of-order children on this
one site (2 findings, 1 site); its header documents Word reads margins by name, so the finding
stays advisory (cell-margin-order.ts:10-12). The proven repair file left this order untouched and
opened clean — direct tolerance evidence.

**Fix — target `superdoc-redlines`.** Serialize the six margin children in schema order:
`top, start, left, bottom, end, right`.

## Violation 3 — `tcPr` child order, **no check** (schema-invalid, Word tolerates)

`word/document.xml:242-256` — the same cell's `w:tcPr` stores `tcMar` (244-249) before
`tcBorders` (250-255):

```xml
<w:tcPr>
  <w:tcW w:w="3757" w:type="dxa"/>
  <w:tcMar>…</w:tcMar>
  <w:tcBorders>…</w:tcBorders>
</w:tcPr>
```

**Rule.** `CT_TcPrBase` (`wml.xsd:2321-2336`) requires `tcBorders` before `tcMar`:

```xsd
<xsd:element name="tcBorders" type="CT_TcBorders" minOccurs="0" maxOccurs="1"/>   <!-- :2327 -->
<xsd:element name="shd" type="CT_Shd" minOccurs="0"/>
<xsd:element name="noWrap" type="CT_OnOff" minOccurs="0"/>
<xsd:element name="tcMar" type="CT_TcMar" minOccurs="0" maxOccurs="1"/>           <!-- :2330 -->
```

Raw xmllint confirms: `document.xml:250: element tcBorders: This element is not expected.` The
validate CLI reports nothing for it — `word-cell-margin-order` judges only the six margin children
*inside* `tcMar`/`tblCellMar` (`cell-margin-order.ts:36-44`); `word-table-properties-misplaced`
enforces order only inside `tblPr`/`tblPrEx` (`word-table-properties.ts:113`, order parents
`["tblPr", "tblPrEx"]`) and covers `tblBorders`, not `tcBorders`; the sole `tcBorders` mention in
`checks/` is the vocabulary list (`wml-vocabulary.ts:571`). Word tolerates it: the proven repair
file kept this order and opened clean.

**Fix — target `docx-validate`.** Add an advisory-severity order check for the `CT_TcPrBase`
child sequence (a sibling of `word-cell-margin-order`, or an extension of
`word-table-properties-misplaced`'s order arm to `tcPr`), so a schema-invalid cell-property order
is at least reported. The generator should also stop emitting it (same fix as violation 2).

## Where the fixes belong and why

- **Violations 1 and 2 → `superdoc-redlines`.** The XML is genuinely defective and docx-validate
  already reports both correctly (error and warning) — no validator gap, no false positive.
  Violation 1 is the one Word refuses the file for.
- **Violation 3 → `docx-validate`.** The validator misses a real schema break (no checkCode
  exists). The producing tool shares one emission-order root cause with violation 2 — cell
  properties serialized schema-wrong (`tcMar` before `tcBorders`; margins `left, right, top,
  bottom`) — so one serializer fix in superdoc-redlines clears both.
- Not `repair`: this is a generator-produced redline, and the lane's own proven minimal repair
  (insert one `w:tc`) already restores Word and `validate()`.

## Verification

Run this session:

- `pnpm exec tsx src/scripts/office/validate.ts really_broken/tool_redlines/repair_prompt/superdoc-redlines__01a11a49_1225.docx --json`
  → `valid: false`, profile `word-valid`; issues: **error** `word-row-without-cell`,
  **warning** `word-cell-margin-order` ×2, plus `paraid-missing-element` ×3 (glossary),
  `xml-bom-leading` ×2, `info` `rel-ids-sequential` ×2 (not XSD violations; out of scope here).
- `xmllint --noout --schema src/scripts/office/schemas/ISO-IEC29500-4_2016/wml.xsd <unpacked>/word/document.xml`
  → exactly two sequence-order errors, `:247 element top` and `:250 element tcBorders`; nothing
  on the cell-less row (the other ~2024 lines are extension-attribute noise: `w14:paraId`/`textId`,
  GUID `w:id` on `w:ins`/`w:del`, `w:authorEmail`, `w:rsidRPr`).
- ElementTree scan of all 20 `word/**/*.xml` parts → exactly 1 cell-less `w:tr`, 1 disordered
  `w:tcMar`, 1 `tcMar`-before-`tcBorders` `w:tcPr`, all in `word/document.xml`.

For the docx-validate fix: a broken fixture carrying `tcMar` before `tcBorders` in a story part
must draw the new advisory, without regressing the tolerance controls
(`tests/fixtures/working/package.word-row-without-cell-repaired.docx`, whose quirk sits in
`word/styles.xml`, and the jubarte lane specimens). The generator fix is proven by this file's
assessment: insert `<w:tc><w:p/></w:tc>` before `</w:tr>` on row `w14:paraId="54E79239"`
(`gridBefore` kept) → Word clean-open and `validate()` valid.
