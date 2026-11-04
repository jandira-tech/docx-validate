# [superdoc-redlines 1614] repair-prompt: dangling `w:lvlPicBulletId` with no `w:numPicBullet`

## Summary

`really_broken/tool_redlines/repair_prompt/superdoc-redlines__01a11a49_1614.docx` (sha256
`9bb46b8b7a93a3cf239483dec2ec2ad86e10e0527fd6470d04f4ca54528f5502`) is a redline produced by
**superdoc-redlines**: `word/document.xml` carries 282 `w:del` and 23 `w:ins` tracked changes, and
`word/numbering.xml` defines 14 `w:abstractNum` / 14 `w:num` entries. Two of those numbering levels
are picture-bullet levels — `<w:numFmt w:val="bullet"/>` with an empty `<w:lvlText w:val=""/>` and a
`<w:lvlPicBulletId w:val="0"/>` — but the part contains **no `w:numPicBullet` definition at all**
(`grep -c numPicBullet word/numbering.xml` = 0). Both references dangle.

- **Live Word outcome:** `repair-prompt` — Word shows *"Word found unreadable content in
  superdoc-redlines__01a11a49_1614.docx. Do you want to recover the contents of this document?"*;
  answering No yields *"Word experienced an error trying to open the file"* and no document window.
- **Manifest row:** `really_broken/manifest.csv` does not exist at the time of this writing, so no
  row could be read (the analyst's reproducibility notes record the same absence).

## Word behaviour

- **Live result (this run):** `repair-prompt`, as above; Word quit cleanly, no crash report.
- **Prior assessment:** PROVEN — `docs/really-broken/assessments.jsonl` (row keyed
  `tool_redlines/repair_prompt/superdoc-redlines__01a11a49_1614.docx`) and
  `docs/really-broken/inspection-log.jsonl` (source `word-ab`) record the minimal-change experiments:
  - original bytes → `unreadable-content-warning` (validator invalid, `word-numbering-picture-bullet`)
  - **delete both** `w:lvlPicBulletId` elements → **`clean-open`** (validator valid; `w:lvlText` untouched)
  - delete only the second (leave the first) → `unreadable-content-warning` (check still fires)
  - delete only the first (leave the second) → `unreadable-content-warning` (check still fires)

  Conclusion: each dangling reference alone is sufficient for Word to offer recovery; deleting both
  opens clean.

## Violation — `word-numbering-picture-bullet` / `dangling-lvlpicbulletid-no-numpicbullet`

**Part:** `word/numbering.xml` — 2 sites, both re-verified in the unpacked package
(`really_broken/unpacked/repair_prompt__superdoc-redlines__01a11a49_1614`).

**Offending XML** (line 119, inside `<w:abstractNum w:abstractNumId="1">`):

```xml
<w:lvl w:ilvl="0" w:tplc="92963006">
  <w:start w:val="1"/>
  <w:numFmt w:val="bullet"/>
  <w:lvlText w:val=""/>
  <w:lvlPicBulletId w:val="0"/>   <!-- dangles: no w:numPicBullet anywhere in this part -->
```

and identically at line 800, inside `<w:abstractNum w:abstractNumId="7">`
(`<w:lvl w:ilvl="0" w:tplc="D062F92E">`). The part defines no `w:numPicBullet` for id 0 — or any id.

**Schema rule** (ISO-IEC 29500-4:2016, bundled `src/scripts/office/schemas/ISO-IEC29500-4_2016/wml.xsd`):

```xml
<!-- wml.xsd:3144, inside CT_Lvl (declared at wml.xsd:3135) -->
<xsd:element name="lvlPicBulletId" type="CT_DecimalNumber" minOccurs="0"/>
<!-- wml.xsd:3192, inside CT_Numbering (declared at wml.xsd:3190) -->
<xsd:element name="numPicBullet" type="CT_NumPicBullet" minOccurs="0" maxOccurs="unbounded"/>
<!-- wml.xsd:3114, CT_NumPicBullet's key (complexType declared at wml.xsd:3109) -->
<xsd:attribute name="numPicBulletId" type="ST_DecimalNumber" use="required"/>
```

Both sides are XSD-optional, so a dangling reference is *schema-legal*: re-running
`xmllint --noout --schema src/scripts/office/schemas/ISO-IEC29500-4_2016/wml.xsd word/numbering.xml`
here reports **zero** findings mentioning `lvlPicBulletId` or `numPicBullet` (only the known
Microsoft-extension attribute noise: `mc:Ignorable`, `w15:restartNumberingAfterBreak`,
`w16cid:durableId`). The broken rule is the ECMA-376 referential constraint — a `w:lvlPicBulletId`
value must name a `w:numPicBullet/@numPicBulletId` **in the same numbering part** — which is exactly
what this repo's `word-numbering-picture-bullet` check implements
(`src/scripts/office/checks/word-numbering-picture-bullet.ts:20`, header citing fixture 1541).

**Proposed fix — in superdoc-redlines' numbering emitter:** keep the reference and its definition
paired. Either

1. emit the `<w:numPicBullet numPicBulletId="0">` (with its `w:pict`/`w:drawing`) that these two
   levels reference — the numbering skeleton looks copied from a Word-saved part
   (`w15:restartNumberingAfterBreak`, `w16cid:durableId` present), so the `w:numPicBullet` block was
   likely dropped in the copy; or
2. stop emitting `w:lvlPicBulletId` on those levels — the proven minimal repair, since deleting both
   references opens clean in Word with the empty `w:lvlText` left as-is (optionally give the level a
   real bullet glyph in `w:lvlText` so the list renders a visible marker).

**Fix target:** `superdoc-redlines` (generator bug — its output is referentially invalid and Word
refuses it).

## Where the fixes belong and why

- **superdoc-redlines** — the redline's own `word/numbering.xml` is genuinely invalid: two
  `w:lvlPicBulletId` values name a picture bullet the same part never defines. Whatever code path
  writes or copies `w:lvl` definitions must carry the matching `w:numPicBullet` along (option 1) or
  strip the reference (option 2). Until then every document it emits with these bullet levels
  triggers Word's recovery prompt.
- **docx-validate — no change.** The `word-numbering-picture-bullet` check already exists, is
  registered for full `word-valid`/`strict` DOCX runs, and reports exactly these two sites (neither a
  validator gap nor a false positive). It is the credited detection for this defect.
- **Not a `repair` case:** the broken artefact is the tool's output, not a hand-authored source
  document, so no repair prescription is filed for the source.

## Verification

- **Validator reproduction (this run):**
  `pnpm exec tsx src/scripts/office/validate.ts really_broken/tool_redlines/repair_prompt/superdoc-redlines__01a11a49_1614.docx`
  → exit 1, exactly two ERROR lines, both `word-numbering-picture-bullet` on `word/numbering.xml`
  (the JSON run's 437 issues include no other error-severity code; the rest are warnings/info such as
  `word-shading-color`, `word-cell-margin-order`, `paraid-missing-element`).
- **Sites re-verified:** `word/numbering.xml:119` and `:800` both `<w:lvlPicBulletId w:val="0"/>`;
  `grep -c numPicBullet word/numbering.xml` = 0. XSD lines 3109–3116, 3135/3144 and 3190–3196 read
  directly (quoted above).
- **Established Word cause (proven assessment, carried in):** deleting **both** `w:lvlPicBulletId`
  elements opens clean (`1614-drop-both.docx`, `clean-open`, validator valid); leaving **either one**
  still triggers the recovery prompt (`1614-leave-first.docx` / `1614-leave-second.docx`, both
  `unreadable-content-warning`) — `docs/really-broken/inspection-log.jsonl`.
- **For the generator fix:** re-emit the document with the pairing restored and confirm a clean open
  under the repo's two-run Word-observation standard (two separate runs, verdicts read from each
  run's Word log; `word-numbering-picture-bullet` fixture 1541 already pins the check's behaviour).
