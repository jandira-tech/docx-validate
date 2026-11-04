# [docxodus 2259] repair-prompt: w:delText inside twenty w:moveFrom move sources

## Summary

- **File**: `really_broken/tool_redlines/repair_prompt/docxodus__01a11a49_2259.docx` (93,735 bytes,
  sha256 `f404c8fbd36a296646cd2ff12ecaea899051dab22b0de45025a42f7486f8d981`, verified this session).
- **What it is**: a docxodus-produced tracked-changes redline (lane `tool_redlines/repair_prompt`,
  label "docxodus 2259"). Every revision stamp in the file reads `w:author="Open-Xml-PowerTools"`,
  `w:date="2000-01-01T00:00:00Z"` — generator-emitted revision markup.
- **Live Word outcome**: **repair-prompt** — "Word found unreadable content in
  docxodus__01a11a49_2259.docx. Do you want to recover the contents of this document?" (Yes/No);
  choosing No was followed by "Word experienced an error trying to open the file" and no document
  window. No crash report; Word quit cleanly.
- **Manifest**: `really_broken/manifest.csv` does not exist in the working tree (checked this
  session — `really_broken/` holds only `issues/`, `tool_redlines/`, `unpacked/`), and the live
  result's `manifestSays` is empty, so there is no manifest row to compare against.
- **One error kind, twenty sites**, all in `word/document.xml`: deletion-marked text (`w:delText`)
  inside move sources (`w:moveFrom`). One site is enough to make Word refuse the file.

## Word behaviour

- **Live run** (given with this ask): repair-prompt, as above. `agrees: true`.
- **Prior assessment** — status `proven`, source `word-ab`
  (`docs/really-broken/assessments.jsonl:144`, `docs/really-broken/inspection-log.jsonl:144`,
  both grepped and read this session):
  - original → `unreadable-content-warning`;
  - `2259-move-plain.docx` (all twenty `w:delText` inside `w:moveFrom` renamed to `w:t`, nothing
    else changed) → `clean-open`, and `validate()` then reports valid with no error codes;
  - `2259-leave-first.docx` (first left as `w:delText`, nineteen renamed) → still
    `unreadable-content-warning`.
  - The 1,171 `w:delText` outside `w:moveFrom` stayed untouched in every variant.

## Violation — `word-move-from-deleted-text` (20 sites, `word/document.xml`)

Offending XML (`word/document.xml:157-165`; the other nineteen sites are the same shape at lines
167, 177, 187, …, 347 — `w:id` 15 through 34):

```xml
<w:moveFrom w:author="Open-Xml-PowerTools" w:id="15" w:date="2000-01-01T00:00:00Z">
  <w:r>
    <w:rPr>
      <w:b/>
      <w:bCs/>
    </w:rPr>
    <w:delText xml:space="preserve">
    </w:delText>
  </w:r>
</w:moveFrom>
```

Schema rule (all lines read this session in `src/scripts/office/schemas/ISO-IEC29500-4_2016/wml.xsd`):

- `wml.xsd:3407` — `<xsd:element name="moveFrom" type="CT_RunTrackChange"/>` (inside the run-level
  choice alongside `ins`, `del`, `moveTo`).
- `wml.xsd:997-1006` —
  ```xml
  <xsd:complexType name="CT_RunTrackChange">
    <xsd:complexContent>
      <xsd:extension base="CT_TrackChange">
        <xsd:choice minOccurs="0" maxOccurs="unbounded">
          <xsd:group ref="EG_ContentRunContent"/>
          <xsd:group ref="m:EG_OMathMathElements"/>
        </xsd:choice>
      </xsd:extension>
    </xsd:complexContent>
  </xsd:complexType>
  ```
- `wml.xsd:2084` — `EG_ContentRunContent` includes `<xsd:element name="r" type="CT_R"/>`.
- `wml.xsd:1745-1750` — `EG_RunInnerContent` declares **both** spellings for a run's text:
  `<xsd:element name="t" type="CT_Text"/>` (1748) and `<xsd:element name="delText" type="CT_Text"/>`
  (1750).

**Why this is a defect despite being schema-valid.** The XSD chain above makes
`w:moveFrom > w:r > w:delText` legal XML — confirmed this session by running
`xmllint --noout --schema src/scripts/office/schemas/ISO-IEC29500-4_2016/wml.xsd <unpacked>/word/document.xml`:
2,818 validity errors, every one of them `w14:paraId`/`w14:textId` extension-attribute noise on
`w:p`/`w:tr`, and **zero** mentioning `moveFrom` or `delText`. The defect is semantic:
`w:delText` is deletion run content whose home is `w:del`; the runs a move source wraps hold
ordinary text (`w:t`). Word treats the combination as unreadable content — the proven
minimal-change experiment (rename to `w:t` → clean open; leave one → still warns) establishes
this beyond the schema's silence.

**Proposed fix.** Inside each of the twenty `w:moveFrom`, rename the `w:delText` to `w:t`,
leaving the 1,171 `w:delText` outside `w:moveFrom` untouched (deletion content in the body is
Word-tolerated and is not this defect). Generator-side: the code path that emits move-source run
content must write `w:t` — docxodus already does exactly that in this very file for the move
*targets*: all four `w:moveTo` elements (verified this session) hold `w:t` runs, so the bug is
specific to the move-source arm.

**Fix target**: `docxodus`.

## Where the fixes belong and why

- **docxodus (generator bug — the fix belongs here).** The redline's XML is genuinely broken
  output: live Word refuses it, and the proven experiment isolates the twenty `w:delText`-in-
  `w:moveFrom` sites as sufficient and necessary for the refusal. The generator stamped its move
  sources with deletion text while its own move targets carry `w:t`; it should emit `w:t` for the
  runs it wraps in `w:moveFrom`, matching its `w:moveTo` arm.
- **docx-validate — no detection gap.** The named check `word-move-from-deleted-text`
  (`src/scripts/office/checks/word-deleted-text.ts:129-157`, helper `moveFromDeletedText`
  at `:76-97`, behaviour documented at `:16-18` citing fixture 0726) already reports exactly
  these twenty sites as error-severity, and `word-blocking.ts:294` makes it Word-blocking.
  Verified this session by running the validator (below). Nothing to add or change.
- **repair — not applicable.** The source document is not the patient; the generator produced the
  invalid markup.

## Verification

- **Validator (ran this session)**:
  `pnpm exec tsx src/scripts/office/validate.ts --json really_broken/tool_redlines/repair_prompt/docxodus__01a11a49_2259.docx`
  → exit 1, `valid: false`, issue codes: `word-move-from-deleted-text` [error] ×20,
  `style-default-missing` [warning] ×4, `paraid-missing-element` [warning] ×1,
  `rel-ids-sequential` [info] ×2. Only the first is error-severity; the plain-text run prints
  twenty "A move source contains deleted text…" errors.
- **Site census (ran this session**, `python3 -I` ElementTree over the unpacked part**)**:
  20 `w:moveFrom`, each holding exactly one `w:r` with one `w:delText`; 1,191 `w:delText` in
  `word/document.xml` (20 inside `w:moveFrom`, 1,171 outside); `word/footer1.xml` has one
  `w:delText` and no `w:moveFrom` (not a site).
- **Established Word cause (proven assessment, carried in above)**: renaming all twenty
  `w:delText` → `w:t` opens clean and `validate()` turns valid with no error codes; leaving one
  still warns. That both pins the cause and validates the check's prescription.
- **For the generator fix**: a docxodus regression test that renders a move revision and asserts
  every text child of `w:moveFrom`/`w:moveTo` runs is `w:t` (never `w:delText`). For this repo:
  this file (or a twenty-site synthetic variant) fits `tests/fixtures/broken/` as a
  `word-move-from-deleted-text` specimen pinning the 20-error output.
- The non-blocking advisories are not part of the refusal: the proven clean-open variant kept
  them unchanged.
