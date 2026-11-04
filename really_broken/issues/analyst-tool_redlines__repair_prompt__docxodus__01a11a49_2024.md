# [docxodus 2024] repair-prompt: two theme relationships in document.xml.rels

## Summary

`tool_redlines/repair_prompt/docxodus__01a11a49_2024.docx` (sha256 `a7918bb8…26e9a10a3`, verified
this session) is a docxodus-produced tracked-changes redline — change records authored
`Open-Xml-PowerTools`, `w:date="2000-01-01T00:00:00Z"` — over a Thai-locale source document. The
**proven, Word-blocking defect**: `word/_rels/document.xml.rels` lists **two theme Relationships**
(`rId6` → `theme/theme1.xml`, `R3` → `/word/theme/P793018ea27585ffc216b2857999ebeef.xml`);
deleting either one opens the file clean. Three further `wml.xsd` sequence violations (bare `w:br`
inside `w:ins`, `numPr` after `spacing` ×5, `link` before `name` ×1) are genuinely invalid XML but
**not** Word-blocking. All four fix targets: **docxodus**.

## Word behaviour

- **Live result (ask-given):** outcome `repair-prompt` — dialog *"Word found unreadable content in
  docxodus__01a11a49_2024.docx. Do you want to recover…"*; answering **No** got *"Word experienced
  an error trying to open the file"* and no document window. `agrees: true`.
- **Manifest row:** `really_broken/manifest.csv` is **absent from disk this session** (no `*.csv`
  anywhere under `really_broken/`; deleted by the same external incident the analyst recorded).
  The ask's own block carries `"manifestSays": ""`. Not readable; not invented.
- **Prior proven assessment** (read this session): `docs/really-broken/assessments.jsonl:109`,
  mirrored at `docs/really-broken/inspection-log.jsonl:109` — `status: proven`, `word.original:
  "unreadable-content-warning"`. **Minimal change that opened clean: delete either theme
  Relationship.** `2024-keep-rId6.docx` (deletes `R3`) and `2024-keep-R3.docx` (deletes `rId6`)
  both open clean, both keep both theme parts and all four duplicate hyperlink relationships
  (`rId7`/`R1`, `rId8`/`R2`), and `validate()` reports no errors on either. Credited check:
  `word-duplicate-theme-relationship`.

## Violation 1 — `word-duplicate-theme-relationship` (Word-blocking, proven)

**Offending XML** — `word/_rels/document.xml.rels` (read this session; lines 6 and 13):

```xml
<Relationship Type="…/officeDocument/2006/relationships/theme" Target="theme/theme1.xml" Id="rId6"/>        <!-- line 6 -->
<Relationship Type="…/officeDocument/2006/relationships/theme" Target="/word/theme/P793018ea27585ffc216b2857999ebeef.xml" Id="R3"/>  <!-- line 13 -->
```

Both target parts exist and are real DrawingML themes (`theme1.xml` names "Office Theme";
`P793018ea….xml` names the Thai "ธีมของ Office" — the source's own theme, parked under a hash part
name beside the generator's fresh `theme1.xml`).

**Schema rule — `ecma/fourth-edition/opc-relationships.xsd`** (read this session, lines 10–25):
`CT_Relationships` holds `<xsd:element ref="Relationship" minOccurs="0" maxOccurs="unbounded"/>`
and `CT_Relationship` requires only `Target`/`Type`/`Id` — no Type uniqueness. Confirmed by
`xmllint --noout --schema …/opc-relationships.xsd` on the part: **"validates"**. The XSD permits
this shape; the broken rule is Word's one-theme-relationship-per-story-root behaviour, which the
validator already names: `src/scripts/office/checks/word-duplicate-relationships.ts:171`.

**Proposed fix (docxodus):** emit exactly **one** theme relationship per story root — when the
source package already has a theme part, reuse it instead of adding a second one (or drop the
original). Either deletion is proven sufficient; the orphaned theme part may stay.

## Violation 2 — bare `w:br` directly inside `w:ins` (xsd-error; not Word-blocking)

**Offending XML** — `word/document.xml:37-39` (read this session):

```xml
<w:ins w:author="Open-Xml-PowerTools" w:id="6" w:date="2000-01-01T00:00:00Z">
  <w:br/>
</w:ins>
```

**Schema rule — `ISO-IEC29500-4_2016/wml.xsd`** (read this session): `CT_RunTrackChange`
(line 997) allows only:

```xsd
<xsd:choice minOccurs="0" maxOccurs="unbounded">
  <xsd:group ref="EG_ContentRunContent"/>
  <xsd:group ref="m:EG_OMathMathElements"/>
</xsd:choice>
```

`EG_ContentRunContent` (line 2084) lists `customXml…r` + `EG_RunLevelElts` — no `br`. `br` is
declared in `EG_RunInnerContent` (line 1745: `<xsd:element name="br" type="CT_Br"/>`), reachable
only inside `CT_R`. Confirmed by `xmllint --noout --schema …/wml.xsd word/document.xml`:
`line 38: element br: This element is not expected.`

**Proposed fix (docxodus):** wrap the break in a run —
`<w:ins …><w:r><w:br/></w:r></w:ins>`.

## Violation 3 — `w:numPr` after `w:spacing` in five list paragraphs (xsd-error; not Word-blocking)

**Offending XML** — `word/document.xml` lines 443, 462, 481, 500, 519 (read this session; a sixth
`numPr` at line 537 has no preceding `spacing` and is fine):

```xml
<w:pPr>
  <w:spacing w:after="0"/>
  <w:numPr>
    <w:ilvl w:val="0"/>
    <w:numId w:val="2"/>
  </w:numPr>
```

**Schema rule — `wml.xsd` `CT_PPrBase`** (read this session, lines 1089–1126): a fixed
`xsd:sequence` in which `numPr` (line 1097) precedes `spacing` (line 1112). Confirmed by the same
xmllint run: `numPr: This element is not expected` at 443, 462, 481, 500, 519 (the "Expected is
one of ( ind, contextualSpacing, … )" list proves the parser had already consumed `spacing`).

**Proposed fix (docxodus):** emit `numPr` before `spacing` in inserted list paragraphs' `pPr`.

## Violation 4 — `w:link` before `w:name` in the Heading1 style (xsd-error; not Word-blocking)

**Offending XML** — `word/styles.xml:447-449` (read this session):

```xml
<w:style w:type="paragraph" w:styleId="Heading1">
  <w:link w:val="Heading1Char"/>
  <w:name w:val="heading 1"/>
```

**Schema rule — `wml.xsd` `CT_Style`** (read this session, lines 3233–3239): a fixed
`xsd:sequence` opening `<xsd:element name="name" … maxOccurs="1"/>` (line 3235) with `link`
(line 3239) fifth. Confirmed by `xmllint --noout --schema …/wml.xsd word/styles.xml`:
`line 449: element name: This element is not expected. Expected is one of ( autoRedefine, … )`.

**Proposed fix (docxodus):** insert `<w:name>` first when adding `<w:link>`/`Heading1Char` to an
existing style.

## Where the fixes belong and why

All four violations belong to **docxodus**, the producing tool:

- The theme duplicate is written by the generator (its own `theme1.xml` beside the preserved
  original); the XSD cannot catch it by design — Word behaviour, already covered by the registered
  check `word-duplicate-theme-relationship`
  (`src/scripts/office/checks/word-duplicate-relationships.ts:171`, tested in
  `tests/checks-word-duplicate-relationships.test.ts` with fixtures
  `broken/package.word-duplicate-theme.docx` / `working/package.word-duplicate-theme-removed.docx`).
- The three wml defects are genuinely invalid redline XML authored by the generator; xmllint
  against the bundled `wml.xsd` reproduces each site.
- **No docx-validate gap.** Verified by running the CLI this session
  (`pnpm exec tsx src/scripts/office/validate.ts really_broken/tool_redlines/repair_prompt/docxodus__01a11a49_2024.docx`):
  the default word-valid profile reports exactly the one theme ERROR; `--profile lenient` reports
  the three deduped XSD errors; `--profile strict` the theme error plus each XSD site. Word-valid's
  silence on the wml defects is correct for its purpose — the proven assessment shows Word opens
  these exact bytes clean once one theme relationship is removed.

xmllint noise ignored as such (extension attributes, not defects): `w14:anchorId`/`w14:editId` at
`document.xml:593` and `mc:Ignorable` at `styles.xml:2`.

## Verification

- **Word cause (already proven, `assessments.jsonl:109`):** two Word runs each on
  `2024-keep-rId6.docx` and `2024-keep-R3.docx` opened clean with all wml defects still present —
  the theme duplicate is the sole Word-blocking defect; both fix directions work.
- **Validator regression guard (run this session):** the three CLI invocations above reproduce the
  expected findings; the theme-check fixtures and suite pin the check's behaviour.
- **For docxodus:** regenerate keeping one theme relationship, `w:br` wrapped in `w:r`, `numPr`
  before `spacing`, `name` before `link`; then `--profile strict` should report zero errors and a
  double Word run should open clean without the recovery dialog.
