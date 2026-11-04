# [superdoc-redlines 2180] repair-prompt: 16 w:hyperlink r:ids name no relationship

## Summary

- Document: `really_broken/tool_redlines/repair_prompt/superdoc-redlines__01a11a49_2180.docx` — a repair-prompt redline produced by **superdoc-redlines** (216 tracked changes: 98 `w:ins`, 118 `w:del`, each wrapping link-styled runs), unpacked at `really_broken/unpacked/tool_redlines__repair_prompt__superdoc-redlines__01a11a49_2180/`.
- Live Word (this ask): **repair-prompt** — "Word found unreadable content in …Do you want to recover…"; answering No gave "Word experienced an error trying to open the file". No manifest row exists (`really_broken/manifest.csv` is absent from the working tree; `manifestSays` is empty), so the comparison is live-only.
- `docx-validate` (word-valid): `valid: false`, 16 error-severity `word-hyperlink-relationship` issues — exactly the Word-blocking defect, matching the 16 rIds of the proven assessment.

## Word behaviour

- Live run (this ask): recovery dialog, then open error; no crash report; outcome `repair-prompt`. No Word run was re-probed for the track-change deviation; its non-blocking status rests on the proven assessment.
- **Proven assessment** — `docs/really-broken/assessments.jsonl:136` (status `proven`; mirrored at `docs/really-broken/inspection-log.jsonl:136`): the 16 `w:hyperlink` `r:id` values (`rId7ecb4eb4`, `rId6266a4a1`, `rId56b19594`, `rId3d9e0512`, `rId3df29dad`, `rId2fcdbc5f`, `rId421fbaa0`, `rId58ca306a`, `rId02019e31`, `rId0846e6a5`, `rId175ed03a`, `rId41e94ebd`, `rId47667523`, `rId124cd329`, `rId53477b0c`, `rId1981b6b1`) are absent from the part relationships; **removing only those r:id attributes — every other byte unchanged — took Word from `unreadable-content-warning` to `clean-open`.** The docx-validate check naming it is `word-hyperlink-relationship`.

## Violation 1 — 16 `w:hyperlink` r:ids name no Relationship (the Word-blocking defect)

Offending XML — `word/document.xml` lines 335, 359, 383, 407, 479, 503, 527, 551, 575, 599, 623, 647, 671, 695, 2793, 4622 (all 16 hyperlinks in the part):

```xml
<w:hyperlink w:history="1" r:id="rId7ecb4eb4">   <!-- rels part defines only rId1..rId14 -->
  <w:del w:id="…" w:author="Live Bench" …> … </w:del>
</w:hyperlink>
```

`word/_rels/document.xml.rels` defines exactly `rId1`–`rId14` and **no hyperlink-type Relationship at all** (`grep -c hyperlink` → 0). Each of the 16 hex-suffixed ids is a dangling package-relationship reference. The part's other r:ids — `rId7`–`rId12` on header/footer references (document.xml:4850-4855) — are defined; not at issue.

Schema rule: `ISO-IEC29500-4_2016/wml.xsd:1296-1304` `CT_Hyperlink` declares `<xsd:attribute ref="r:id"/>` (line 1303), typed by `shared-relationshipReference.xsd` `ST_RelationshipId` — a plain `<xsd:restriction base="xsd:string"/>`, so an XSD walk cannot catch a dangling id; the binding obligation is the OPC semantics that an `r:id` value must name a Relationship in the part's relationships part (ECMA-376-1 §17.16.22; Part 2 §9.3). Word session 28 (cited in `src/scripts/office/checks/word-hyperlink.ts:16-22`) pinned the only accepted Type URIs.

**Proposed fix** (in superdoc-redlines): when emitting `w:hyperlink` with an `r:id`, also write the matching relationship into `word/_rels/document.xml.rels`:

```xml
<Relationship Id="rId7ecb4eb4"
  Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"
  Target="https://…" TargetMode="External"/>
```

— or, when no external target is known, omit the `r:id` (keep the styled run, or use `w:anchor` for internal links). Fix target: **superdoc-redlines**.

## Violation 2 — every tracked change carries a GUID `w:id` and an undeclared `w:authorEmail` (not Word-blocking)

Offending XML — all 98 `w:ins` and 118 `w:del` elements (first at `word/document.xml:36`; identical spelling throughout, `grep` found no non-GUID `w:id`):

```xml
<w:ins w:id="af355071-59c4-45c9-8937-bf3724d4b6e8" w:author="Live Bench"
       w:authorEmail="ai@example.com" w:date="2026-10-13T03:50:00.000Z">
```

Schema rule (`ISO-IEC29500-4_2016/wml.xsd`):

- lines 843-845, `CT_Markup`: `<xsd:attribute name="id" type="ST_DecimalNumber" use="required"/>`; lines 56-58, `ST_DecimalNumber` = `<xsd:restriction base="xsd:integer"/>` — a GUID string is not an integer.
- lines 846-853, `CT_TrackChange` extends `CT_Markup` with only `author` (required) and `date` (optional) — no `authorEmail` is declared there or anywhere in the bundled schemas (`grep -rn authorEmail src/scripts/office/schemas/` → no match).

Reproduced with `xmllint --noout --schema wml.xsd word/document.xml`: 216 × "'…' is not a valid value of the atomic type '…ST_DecimalNumber'" and 216 × "attribute '…authorEmail' is not allowed" (plus 248 `w14:paraId`/`w14:textId` errors — Microsoft-extension noise, excluded per the ask).

Not the Word-blocking defect: the proven assessment's clean-open left all 216 ins/del elements unchanged. The word-valid profile has no named check for it (the `--json` issue list names nothing of the kind); `--strict` reports all 432 as `xsd-error`.

**Proposed fix** (in superdoc-redlines): emit integer `w:id` values (as Word does) and drop `w:authorEmail`; if the author's email must survive, use an extension attribute in a properly declared, `mc:Ignorable`-listed namespace — never `w:`. Fix target: **superdoc-redlines**.

## Where the fixes belong and why

- **superdoc-redlines — both violations.** The redline XML is genuinely invalid: violation 1 is the proven Word-blocking defect, violation 2 is ISO-XSD-invalid emission by the same generator. The generator must stop emitting both shapes; no repair prescription is needed (this is not a broken source document).
- **docx-validate — no correctness change required.** Violation 1 is already caught exactly by `word-hyperlink-relationship` (`src/scripts/office/checks/word-hyperlink.ts`), 16/16 with the exact rIds. Violation 2 is surfaced by `--strict` as `xsd-error` (216 + 216 of the 433); word-valid correctly does not block it, since the proven assessment shows Word opens the file clean with all 216 ins/del present. Optional hardening, out of scope here: an advisory named check for non-integer tracked-change ids / undeclared `w:authorEmail`, so this generator-bug family is visible without `--strict`.

## Verification

- Validator, word-valid (ran in this ask): `pnpm exec tsx src/scripts/office/validate.ts really_broken/tool_redlines/repair_prompt/superdoc-redlines__01a11a49_2180.docx --json` → `valid: false`; 16 × error `word-hyperlink-relationship`; 1 `paraid-missing-element` + 2 `rels-unreferenced` warnings; 338 `run-props-redundant` + 1 `rel-ids-sequential` info.
- Validator, strict (ran in this ask): `--strict --json` → 433 × `xsd-error` = 216 GUID-id + 216 `authorEmail` + 1 unrelated `footnotePr`.
- XSD reproduction (ran in this ask): `xmllint --noout --schema src/scripts/office/schemas/ISO-IEC29500-4_2016/wml.xsd word/document.xml` → 216 + 216 (+ 248 w14 noise).
- Established Word cause (proven, `assessments.jsonl:136`): removing only the 16 hyperlink `r:id` attributes → clean-open. Minimal-change experiment for the generator fix: apply the relationship-adding variant above (not the removing variant) and re-open in Word — expected clean-open with working links. Violation 2 needs no Word experiment (non-blocking); after the fix, `--strict` must pass with no ST_DecimalNumber/authorEmail findings.
