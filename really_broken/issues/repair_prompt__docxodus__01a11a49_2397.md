# [docxodus 2397] Eight dangling `w:lvlPicBulletId` references, zero `w:numPicBullet` definitions

## Summary

`tool_redlines/repair_prompt/docxodus__01a11a49_2397.docx` (sha256
`0a4d44614d5ffaf54939799b274cafb0cd9b91dd9e8bdf4c913ca48f82601e73`) is a redlined
document produced by **docxodus** (repair_prompt lane). It is a standard package:
18 XML parts plus one JPEG, one numbering part (`word/numbering.xml`) with 18
`w:abstractNum` definitions.

The numbering part carries **eight** `<w:lvlPicBulletId w:val="0"/>` references — one
each on the level-0 bullet `w:lvl` of `w:abstractNum` ids 21, 9, 12, 4, 35, 1, 15 and 0
(unpacked `word/numbering.xml` lines 11, 125, 239, 353, 467, 581, 695, 1089) — and
**zero** `w:numPicBullet` elements anywhere in the part (verified: `grep -c numPicBullet`
returns 0). Every picture-bullet reference dangles.

- **Live Microsoft Word result:** `repair-prompt` — "Word found unreadable content in
  docxodus__01a11a49_2397.docx. Do you want to recover the contents of this document?";
  answering No yields "Word experienced an error trying to open the file" and no document
  window. No crash report; Word quit cleanly.
- **docx-validate** (`pnpm exec tsx src/scripts/office/validate.ts <file>`, run this
  session): exit 1 with **exactly 8 error-severity issues**, all code
  `word-numbering-picture-bullet` on `word/numbering.xml`. Validator and Word agree.

The manifest row named by the lane (`really_broken/manifest.csv`) does not exist in this
checkout (`ls` reports "No such file or directory"); the live-Word record itself carries
`manifestSays: ""`. No prior assessment exists for this file: `grep
"docxodus__01a11a49_2397"` over `docs/really-broken/assessments.jsonl` and
`inspection-log.jsonl` returns no row.

## Word behaviour

- **Live (this run, single reported observation):** recovery dialog, then open error
  after No. Outcome `repair-prompt`, `agrees: true` with the validator's blocking errors.
- **Manifest:** absent from this checkout (see above) — nothing recorded there.
- **Prior assessment for this file:** none.
- **The defect kind is proven on four sibling corpus files** with the same shape, all
  status `proven` in `docs/really-broken/assessments.jsonl`:
  - line 79, `docxodus__01a11a49_1541.docx` (same tool, one ref `w:val="0"`):
    original `unreadable-content-warning`; after removing only that element
    (`lvlPicBulletId_removed`) → `clean-open`. "Removing the numbering relationship also
    opens clean… Removing only that element, and leaving the relationship, opens clean."
  - line 118, `superdoc-redlines__01a11a49_0812.docx` (one ref `w:val="1"`): deleting the
    element → `clean-open`.
  - line 124, `superdoc-redlines__01a11a49_1614.docx` (two refs): deleting **both** →
    `clean-open`; leaving **one** → still `unreadable-content-warning`. "One reference is
    enough."
  - line 129, `superdoc-redlines__01a11a49_2142.docx` (three refs): deleting all three →
    `clean-open`.

  So for this defect kind the Word cause and the fix are established: a dangling
  `w:lvlPicBulletId` makes Word offer recovery; one dangling reference suffices; all of
  them must go. For **this exact file** the cause is not yet proven — the live result and
  the sibling record are the evidence, and the minimal-change experiment below would prove
  it.

## Violation 1 — `word-numbering-picture-bullet` (8 sites, `word/numbering.xml`)

**Offending XML** (first site, unpacked `word/numbering.xml:7-12`, inside
`w:abstractNum w:abstractNumId="21"`; the other seven sites are identical in shape):

```xml
<w:lvl w:ilvl="0" w:tplc="08090007">
  <w:start w:val="1"/>
  <w:numFmt w:val="bullet"/>
  <w:lvlText w:val=""/><!-- w:val holds U+F0B7 (bytes EF 82 B7), the Symbol-font bullet glyph; it is not empty -->
  <w:lvlPicBulletId w:val="0"/>
  <w:lvlJc w:val="left"/>
```

The part defines no `<w:numPicBullet>` at all, so `w:val="0"` names a picture bullet that
does not exist. (Note: the incoming analyst snippet rendered `w:lvlText` as empty; the
actual bytes are `EF 82 B7`, a private-use bullet glyph — which is why deleting the
picture reference leaves a working text-bullet level.)

**Schema rule** — `src/scripts/office/schemas/ISO-IEC29500-4_2016/wml.xsd`:

- `wml.xsd:3144` (inside `CT_Lvl`, the referencing element):
  ```xml
  <xsd:element name="lvlPicBulletId" type="CT_DecimalNumber" minOccurs="0"/>
  ```
- `wml.xsd:3110-3115` (`CT_NumPicBullet`, the definition that is missing):
  ```xml
  <xsd:complexType name="CT_NumPicBullet">
    <xsd:choice>
      <xsd:element name="pict" type="CT_Picture"/>
      <xsd:element name="drawing" type="CT_Drawing"/>
    </xsd:choice>
    <xsd:attribute name="numPicBulletId" type="ST_DecimalNumber" use="required"/>
  </xsd:complexType>
  ```
- `wml.xsd:3189-3196` (`CT_Numbering`, where definitions would live):
  ```xml
  <xsd:complexType name="CT_Numbering">
    <xsd:sequence>
      <xsd:element name="numPicBullet" type="CT_NumPicBullet" minOccurs="0" maxOccurs="unbounded"/>
      ...
  ```

Because `numPicBullet` is `minOccurs="0"`, the XSD cannot see the dangling reference —
raw `xmllint --noout --schema wml.xsd` on this part flags only Microsoft extension
attributes (`p14:restartNumberingAfterBreak`, `p16cid:durableId`), never this defect. It
is a **referential** violation, which is exactly why docx-validate carries the dedicated
check `src/scripts/office/checks/word-numbering-picture-bullet.ts` (id at line 21, issue
emitted at lines 53-61; its doc comment cites the Word-proven fixture 1541).

**Proposed fix** — in the generator: when emitting (or copying) a `w:lvl` that carries
`w:lvlPicBulletId`, either emit the `w:numPicBullet` definition with the matching
`w:numPicBulletId` in the same part, or drop the `w:lvlPicBulletId` element. For this
document, deleting all eight `<w:lvlPicBulletId w:val="0"/>` elements is the proven-shape
repair: each affected level already has `w:numFmt w:val="bullet"` and a non-empty
`w:lvlText` glyph, so the level remains a well-formed text-bullet level.

**Fix target:** **docxodus** (generator bug — it kept the references while dropping every
`w:numPicBullet` definition).

## Where the fixes belong and why

- **docxodus** — owns this defect. The numbering part was generated with picture-bullet
  references whose definitions were never written. The `w:val="0"` id and the eight-fold
  repetition across `w:abstractNum` ids 21/9/12/4/35/1/15/0 indicate the source document
  used one shared picture bullet (id 0) that docxodus dropped wholesale while keeping the
  `w:lvlPicBulletId` hooks. Sibling `docxodus__01a11a49_1541` shows the same signature
  (one dangling ref, no definitions), so this is a recurring generator behaviour, not a
  one-off corruption.
- **docx-validate** — no change. The existing `word-numbering-picture-bullet` check
  reports all eight sites at error severity and its verdict matches the live Word refusal;
  nothing is missed and nothing is misreported.
- **repair** — not applicable as a fix target (the broken source document is not the
  case here); the document-level repair is the deletion described above, recorded as the
  verification experiment.

## Verification

1. **Re-verify (done this session):** `grep -o '<w:lvlPicBulletId[^/]*/>' word/numbering.xml`
   → exactly 8 `<w:lvlPicBulletId w:val="0"/>`; `grep -c numPicBullet word/numbering.xml` → 0.
   `pnpm exec tsx src/scripts/office/validate.ts really_broken/tool_redlines/repair_prompt/docxodus__01a11a49_2397.docx`
   → exit 1, exactly 8 error lines, all `word-numbering-picture-bullet`.
2. **Minimal-change experiment to prove the cause for this file:** unpack the docx, delete
   the eight `<w:lvlPicBulletId w:val="0"/>` elements from `word/numbering.xml` (no other
   byte changes), repack, and open in Word twice with logs per the repo's probing
   standard. Expected: `clean-open` in both runs (the sibling record 1614/2142 shows
   deleting *only some* refs keeps the recovery dialog — all eight must go). If any
   dangling ref is left in place, the prompt must persist.
3. **Regression guard for the generator:** a docxodus test that renders a redline from a
   source document containing a picture bullet list, then asserts the output's numbering
   part satisfies: every `w:lvlPicBulletId/@w:val` ∈ {`w:numPicBullet/@w:numPicBulletId`}
   of the same part — or simply asserts `word-numbering-picture-bullet` yields no issues
   on the output.
