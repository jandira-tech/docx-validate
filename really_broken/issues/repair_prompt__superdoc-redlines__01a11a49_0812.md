# [superdoc-redlines 0812] repair-prompt: dangling `w:lvlPicBulletId` names no `w:numPicBullet`

## Summary

`tool_redlines/repair_prompt/superdoc-redlines__01a11a49_0812.docx` (sha256 `1c1465090b1f178bebf09de9c9f6c0a5068a3c85d9386fc36ecc35cb8b289aae`, 24 XML parts, source artifact `https://jubarte.pro/live/artifacts/01a11a49-7e73-770c-997d-0013581d67e9/812/superdoc-redlines-redline.docx`) is a tracked-changes redline produced by **superdoc-redlines**. `word/numbering.xml` keeps a `w:lvlPicBulletId w:val="1"` reference while carrying **zero** `w:numPicBullet` definitions, so the picture-bullet reference names a target that does not exist.

- **Live Word outcome** (this run's fresh observation): `repair-prompt` — Word showed *"Word found unreadable content in superdoc-redlines__01a11a49_0812.docx. Do you want to recover the contents of this document?"*; answering No was followed by *"Word experienced an error trying to open the file"* and no document opened. No crash report; Word quit cleanly.
- **Manifest row** agrees in substance: the lane manifest records the earlier observation `soffice,Word could not load it in 2 passes: document loaded empty (Word could not read it)`. (Read from `/Users/arthrod/temp/T/neurotic_docx_bench/maybe_future_corpus/really_broken/manifest.csv`; `really_broken/manifest.csv` was absent from the working tree when I looked — this lane's files are being consumed by concurrent workers, and I restored the packed docx byte-identically from the same corpus before verifying.)

One violation kind, one site. Status: **confirmed** — re-verified against the bundled XSDs, the unpacked part, and this repo's validator (commands and outputs under *Verification*).

## Word behaviour

- **Live result**: recovery dialog, then open error after declining recovery (above). Outcome `repair-prompt`.
- **Manifest row** (`maybe_future_corpus/really_broken/manifest.csv`, grep `superdoc-redlines__01a11a49_0812`): earlier soffice observation "Word could not read it".
- **Prior assessment — PROVEN** (`docs/really-broken/assessments.jsonl:118`, `docs/really-broken/inspection-log.jsonl:118`, source `word-ab`): the original file draws `unreadable-content-warning`; deleting **only** that one `w:lvlPicBulletId` element (`0812-no-picbullet.docx`, no other byte changes) opens **clean** and `validate()` turns **valid**. The 1713 `w:delText` elements in `word/document.xml` stayed in that experiment and are explicitly **not** the cause (I count `w:delText` on 3426 tag lines of the unpacked `document.xml`, consistent with 1713 open/close pairs).

## The violation — dangling `w:lvlPicBulletId` (`word/numbering.xml:48`)

**Offending XML** — `w:lvl w:ilvl="0"` of `w:abstractNum w:abstractNumId="2"` (unpacked `word/numbering.xml:39-64`):

```xml
<w:abstractNum w:abstractNumId="2">
  <w:nsid w:val="00000004"/>
  <w:multiLevelType w:val="singleLevel"/>
  <w:tmpl w:val="00000004"/>
  <w:name w:val="WW8Num15"/>
  <w:lvl w:ilvl="0">
    <w:start w:val="1"/>
    <w:numFmt w:val="bullet"/>
    <w:lvlText w:val=""/>
    <w:lvlPicBulletId w:val="1"/>   <!-- line 48: names picture bullet 1 -->
    ...
```

`grep -c numPicBullet word/numbering.xml` → **0**. The id `1` is defined nowhere in the part.

**Schema rule** (`src/scripts/office/schemas/ISO-IEC29500-4_2016/wml.xsd`, lines read this session):

- `wml.xsd:3144` — `CT_Lvl`'s sequence: `<xsd:element name="lvlPicBulletId" type="CT_DecimalNumber" minOccurs="0"/>`
- `wml.xsd:3192` — `CT_Numbering`'s sequence: `<xsd:element name="numPicBullet" type="CT_NumPicBullet" minOccurs="0" maxOccurs="unbounded"/>`
- `wml.xsd:3109-3114` — `<xsd:complexType name="CT_NumPicBullet">` … `<xsd:attribute name="numPicBulletId" type="ST_DecimalNumber" use="required"/>`

A `lvlPicBulletId` value may only name the `numPicBulletId` of a `CT_NumPicBullet` in the same `CT_Numbering`. That cross-reference is **semantic**: `wml.xsd` declares no `xs:key`/`xs:keyref`/`xs:unique` (`grep -c` → 0), so raw XSD validation passes the part — I confirmed `xmllint --noout --schema …/wml.xsd word/numbering.xml` prints `validates` (exit 0). The rule is enforced by this repo's check **`word-numbering-picture-bullet`** (`src/scripts/office/checks/word-numbering-picture-bullet.ts:47-61`), which fired as the run's only ERROR.

**Corroboration that the generator dropped the definitions** (not that the reference is stray): `word/_rels/numbering.xml.rels` still relates `rId1 → media/image1.png` and `rId2 → media/image2.png`, and both images are stored (15×15 bullet-sized PNGs). A numbering part references images **only** through a `w:numPicBullet`'s `pict`/`drawing` child (`wml.xsd:3110-3113`); `grep r:id|r:embed|r:link word/numbering.xml` → 0. The relationships and media survived; the `w:numPicBullet` definitions that consumed them did not. Also note `abstractNum 2` is not even mapped by any `w:num` (numIds 1-9 map abstractNumIds 11-18, 14, 17) — Word refuses the file on the dangling reference in the definition part regardless of reachability, which matches the check's part-level judgement.

**Proposed fix** — in **superdoc-redlines**' numbering serialization: when writing/copying `word/numbering.xml`, keep the `w:numPicBullet` definitions (with their `pict`/`drawing` children and `numPicBulletId`s) alongside any `w:lvlPicBulletId` references, since it already preserves their media parts and relationships; or, if picture-bullet definitions are deliberately stripped, strip/rewrite the `w:lvlPicBulletId` references in the same pass (e.g. emit a plain text bullet `w:lvlText` instead). The two must be written or dropped **together**.

**Fix target**: `superdoc-redlines` (generator bug — invalid redline XML).

## Where the fixes belong and why

- **superdoc-redlines** — the one defect. The redline's `word/numbering.xml` is semantically invalid OOXML (reference with no definition), and Word refuses the whole file for it. The media parts and rels the definitions would use are still in the package, so the generator's numbering-part writer dropped the definitions while keeping a reference to them.
- **Not docx-validate** — the check `word-numbering-picture-bullet` exists and is correct: my full `validate.ts` run reports exactly this one ERROR (exit 1) and names the fix; no missing check, no false positive.
- **Not `repair`** — the lane category is `tool_redlines` (generator output), not a broken source document. For rescuing the existing artifact, the proven minimal repair is deleting the single `w:lvlPicBulletId` at `word/numbering.xml:48`.

## Verification

All run this session against the restored byte-identical copy (sha256 verified equal after copy; the lane directory had again lost the packed file and pre-unpacked dir):

1. `pnpm exec tsx src/scripts/office/unpack.ts really_broken/tool_redlines/repair_prompt/superdoc-redlines__01a11a49_0812.docx really_broken/unpacked/repair_prompt__superdoc-redlines__01a11a49_0812 --merge-runs false --simplify-redlines false` → `Unpacked … (24 XML files)`.
2. `pnpm exec tsx src/scripts/office/validate.ts really_broken/tool_redlines/repair_prompt/superdoc-redlines__01a11a49_0812.docx` → exit 1, one ERROR: *"A numbering level references picture bullet 1, and this part has no w:numPicBullet with that id. Microsoft Word offers to recover the file. Remove the w:lvlPicBulletId, or add the matching picture bullet."* — matching `word-numbering-picture-bullet.ts:53-61`.
3. `xmllint --noout --schema src/scripts/office/schemas/ISO-IEC29500-4_2016/wml.xsd <unpacked>/word/numbering.xml` → `validates`, exit 0 (raw XSD is blind; the rule is semantic).
4. Site reads: `grep -n lvlPicBulletId word/numbering.xml` → line 48 only; `grep -c numPicBullet` → 0; `grep -c "r:id\|r:embed\|r:link" word/numbering.xml` → 0; `numbering.xml.rels` lists the two image relationships; both PNGs stored.
5. **Established Word cause (proven assessment, carried in)**: deleting the single `w:lvlPicBulletId` (`0812-no-picbullet.docx`) → Word `clean-open`, validator `valid` with no error codes; the 1713 `w:delText` elements stayed. Regression home for this repo: the existing check plus a fixture shaped like this part (reference, no definition, orphaned image rels) — `word-numbering-picture-bullet` already rejects it, so no new check is needed.
