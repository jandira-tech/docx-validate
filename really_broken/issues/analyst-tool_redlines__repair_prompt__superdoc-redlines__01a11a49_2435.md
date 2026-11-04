# [superdoc-redlines 2435] repair-prompt: parts stored without [Content_Types].xml coverage

## Summary

`really_broken/tool_redlines/repair_prompt/superdoc-redlines__01a11a49_2435.docx`
(lane category `repair_prompt`, producing tool **superdoc-redlines**) is a 15-entry
WordprocessingML package whose `[Content_Types].xml` declares **13 `<Override>` elements and
zero `<Default>` elements**. Two stored parts are covered by no entry of either kind:

- `docProps/custom.xml` (557 bytes) — a custom-properties part holding `SuperdocVersion`
  (`1.10.0`) and `DocumentGuid`, the superdoc-redlines signature;
- `word/numbering.xml` (8,882 bytes) — a real `w:numbering` definitions part (2 `w:num`,
  2 `w:abstractNum`) that `word/styles.xml` genuinely uses (1 `w:numPr` element, 2 `w:numPr`
  string matches; `word/document.xml` has none).

OPC requires a content type for every part; with no Override and no Default, neither part
resolves one. **Live Word outcome matches the analyst and the lane label**: Word shows the
recovery dialog, and declining it yields an open error with no document window.

Re-verification this session ran on the repo copy (`shasum -a 256` =
`73d9e3652f27a533abfca2b814e61fb212a3c0eea734c41f66a1a4f1f05e376b`, identical to the hash the
analyst pinned). The ask's unpacked path `…/unpacked/analyst-tool_redlines__…__2435` does not
exist; I judged the sibling unpacking `…/unpacked/tool_redlines__repair_prompt__superdoc-redlines__01a11a49_2435/`.

## Word behaviour

- **Live result (given in this ask):** `repair-prompt` — dialog
  *"Word found unreadable content in superdoc-redlines__01a11a49_2435.docx. Do you want to
  recover the contents of this document?"*; clicking **No** was followed by *"Word experienced
  an error trying to open the file"* and no document opened. No crash report; Word quit
  cleanly. `agrees: true`.
- **Manifest row:** none — `really_broken/manifest.csv` does not exist in this repo yet
  (`ls` this session), and `manifestSays` in the live result is empty.
- **Prior assessment:** none — `grep -c "superdoc-redlines__01a11a49_2435"` returns **0** in
  both `docs/really-broken/assessments.jsonl` and `docs/really-broken/inspection-log.jsonl`.

## Violation 1 — `parts-without-content-type` (`ct-uncovered-part`, 2 sites)

**Offending XML** — `[Content_Types].xml` in full (read this session; 13 Override, 0 Default):

```xml
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Override PartName="/_rels/.rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Override PartName="/word/_rels/document.xml.rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Override PartName="/word/settings.xml" ContentType="…wordprocessingml.settings+xml"/>
  <Override PartName="/word/fontTable.xml" ContentType="…wordprocessingml.fontTable+xml"/>
  <Override PartName="/word/footer2.xml" ContentType="…wordprocessingml.footer+xml"/>
  <Override PartName="/word/header2.xml" ContentType="…wordprocessingml.header+xml"/>
  <Override PartName="/word/footer1.xml" ContentType="…wordprocessingml.footer+xml"/>
  <Override PartName="/word/header1.xml" ContentType="…wordprocessingml.header+xml"/>
  <Override PartName="/word/styles.xml" ContentType="…wordprocessingml.styles+xml"/>
  <Override PartName="/word/theme/theme1.xml" ContentType="…theme+xml"/>
  <Override PartName="/word/document.xml" ContentType="…document.main+xml"/>
  <Override PartName="/docProps/app.xml" ContentType="…extended-properties+xml"/>
  <Override PartName="/docProps/core.xml" ContentType="…core-properties+xml"/>
</Types>
<!-- no Override for /docProps/custom.xml or /word/numbering.xml; no Default at all -->
```

`grep -o 'PartName="/\(docProps/custom\|word/numbering\)\.xml"'` over the stream matches
nothing. Both parts are nonetheless stored in the package (directory listing this session).

**Schema rule** — `src/scripts/office/schemas/ecma/fourth-edition/opc-contentTypes.xsd:7-26`
(read this session):

```xsd
<xs:element name="Types" type="CT_Types"/>
<xs:element name="Default" type="CT_Default"/>
<xs:element name="Override" type="CT_Override"/>

<xs:complexType name="CT_Types">
  <xs:choice minOccurs="0" maxOccurs="unbounded">
    <xs:element ref="Default"/>
    <xs:element ref="Override"/>
  </xs:choice>
</xs:complexType>

<xs:complexType name="CT_Default">
  <xs:attribute name="Extension" type="ST_Extension" use="required"/>
  <xs:attribute name="ContentType" type="ST_ContentType" use="required"/>
</xs:complexType>

<xs:complexType name="CT_Override">
  <xs:attribute name="ContentType" type="ST_ContentType" use="required"/>
  <xs:attribute name="PartName" type="xs:anyURI" use="required"/>
</xs:complexType>
```

The XSD constrains only the stream's shape — it **cannot express** "every stored part must be
covered", which is why
`xmllint --noout --schema …/opc-contentTypes.xsd "[Content_Types].xml"` reports
`validates` (run this session) while the package is still broken. The operative rule is
**ECMA-376 Part 2 §10.1.2.4**: a part's content type resolves by exact `Override` match on the
part name, then by extension `Default`; with neither, the part has no content type, and OPC
(MCE/OPC package model) does not permit an untyped part. The repo anchors the same rule at
`src/scripts/office/validators/base.ts:1569-1572` (resolution order) and emits the error at
`base.ts:1686-1692`.

**Validator agreement (run this session):**
`pnpm exec tsx src/scripts/office/validate.ts really_broken/tool_redlines/repair_prompt/superdoc-redlines__01a11a49_2435.docx`
→ `valid: false`, profile `word-valid`, exactly **2 error-severity `ct-uncovered-part`**
findings (`docProps/custom.xml`, `word/numbering.xml`). Warnings: `rels-unreferenced` ×2 on the
same two parts (no relationship names either: `_rels/.rels` holds core-properties /
extended-properties / officeDocument; `word/_rels/document.xml.rels` holds styles, header1-2,
footer1-2, fontTable, settings, theme — **no numbering relationship either**), and
`paraid-missing-element` ×5 (advisory). Info: `rel-ids-sequential` ×2.

**Proposed fix** — superdoc-redlines writes both parts but never adds their content-types
entries (and never relates `numbering.xml`). It must, when writing these parts:

1. add `<Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>`;
2. add `<Override PartName="/docProps/custom.xml" ContentType="application/vnd.openxmlformats-officedocument.custom-properties+xml"/>`;
3. add the `numbering` relationship to `word/_rels/document.xml.rels`
   (`Type="…/relationships/numbering" Target="numbering.xml"`) — a Word-native package always
   relates its numbering part; and, conventionally, a `custom-properties` relationship from
   `_rels/.rels` to `docProps/custom.xml`.

A blanket `<Default Extension="xml" ContentType="application/xml"/>` would also cover both
parts, but the typed Overrides match what Word itself writes and keep part-type checks
meaningful; prefer the Overrides.

**Fix target:** `superdoc-redlines` (generator bug).

## Where the fixes belong and why

- **superdoc-redlines** — the redline's package is genuinely invalid: parts stored with no
  content type violate ECMA-376 Part 2 §10.1.2.4 and Word offers recovery (live result this
  ask). The generator's packaging step must emit the `[Content_Types].xml` entries (and the
  numbering relationship) for every part it writes. The `SuperdocVersion`/`DocumentGuid`
  custom part and the styles-referenced numbering part show this is its own output, not a
  passthrough of a broken source document.
- **Not docx-validate** — the validator already reports this exactly: `ct-uncovered-part`,
  error severity, correct parts, correct message ("Word prompts to repair such packages"),
  under the default `word-valid` profile. No missing check, no false positive.
- **Not repair** — the producing tool is a redline generator, not a broken source document.

## Verification

The live Word `repair-prompt` and the OPC rule are the evidence; **the causal link is not yet
proven by a controlled experiment** (no prior assessment exists for this file). Minimal-change
experiment to prove it: take these exact bytes, add **only** the two `Override` entries above
(plus the numbering relationship) to `[Content_Types].xml` / `document.xml.rels`, repack
without touching anything else, and open in Word **twice with logs** per the repo's probe
protocol — expect a clean open both runs. A one-at-a-time variant (only the `numbering.xml`
Override, then only the `custom.xml` Override) would tie the dialog to each part individually;
`docProps/custom.xml` is a metadata part and may turn out benign, while `word/numbering.xml`
is a word-part read from the main document's relationships, so it is the likelier sole blocker.

Repo-side regression: the existing `ct-uncovered-part` check already covers this shape; a
fixture cut from these bytes belongs under `tests/fixtures/broken/` named per convention
(e.g. `superdoc-redlines-2435.missing-content-type-overrides.docx`), asserting two
`ct-uncovered-part` errors for `docProps/custom.xml` and `word/numbering.xml`.
