# [docxodus 0162] bare w:delText / w:delInstrText inside text-box stories makes Word offer recovery

## Summary

`tool_redlines/repair_prompt/docxodus__01a11a49_0162.docx` is a docxodus-produced
tracked-changes redline (corpus id `docxodus__01a11a49_0162`, lane category
`repair_prompt`). The reported defect: inside one `mc:AlternateContent` text box in
`word/document.xml`, the `mc:Choice` (`wps:txbx`) branch's `w:txbxContent` story carries
deleted run content (`w:delText` / `w:delInstrText`) in plain `w:r` runs with no `w:del`
ancestor inside the story — the construct that makes Word offer to recover the file.
The `mc:Fallback` VML twin carries the same elements, which the repo's check deliberately
exempts.

**Status of this issue: disputed — for absence of the specimen, not because the defect
class is in doubt.** The fixture is not in the working tree: `really_broken/tool_redlines/
repair_prompt/` holds 30 fixtures (ids 1951–2435 and others; `ls` this session), a repo-wide
`find … -name "*_0162*"` returns nothing, and neither is there an unpacked copy
(`really_broken/unpacked/` holds 32 `tool_redlines__repair_prompt__*` directories, none
`0162`). `really_broken/manifest.csv` does not exist either. So the 12 reported sites in
`word/document.xml` could not be re-read, and the validator could not be re-run on the
bytes. Everything else below — the schema rule, the registered check, and Word evidence for
this exact shape — was re-verified this session and holds.

## Word behaviour

- **Live Word result (this session's lane): `failed`, `agrees: false`.** "The file does not
  exist … so Word was never opened." The probe could not disagree with anything; it never
  ran. `manifestSays` is empty because no manifest row exists.
- **Prior assessment: PROVEN** — `docs/really-broken/assessments.jsonl:35` (mirrored at
  `docs/really-broken/inspection-log.jsonl:35`), path
  `tool_redlines/repair_prompt/docxodus__01a11a49_0162.docx`:
  > word/document.xml contains 2 w:txbxContent elements with 24 w:delText elements between
  > them. Microsoft Word offers to recover the file. … Inside those 2 w:txbxContent elements
  > only, change w:delText to w:t. No other part changes.
  > `"word": {"original": "unreadable-content-warning", "text_box_delText_changed_to_w_t": "clean-open"}`

  The same package still contains 1376 `w:delText` and 5 `w:delInstrText` outside those text
  boxes (including headers/footers); they stay in the repair and Word tolerates them.
- **Independent hash-bound Word evidence for the same shape** (not the same bytes —
  `tests/fixtures-all.manifest.json`, all Word 16.115): the broken specimens
  `broken/package.word-textbox-choice-deleted-text.docx` (deleted text in the `mc:Choice`
  text box — this exact shape) and `broken/package.word-textbox-deleted-instruction.docx`
  (`w:delInstrText` variant) both open with `unreadable-content-warning`; the working
  controls `working/package.word-textbox-deleted-text-removed.docx`,
  `working/package.word-textbox-deleted-text-shape-repaired.docx` (repaired → `clean-open`),
  `working/package.word-textbox-fallback-deleted-text.docx` (Fallback tolerated →
  `clean-open`) and `working/package.word-textbox-inner-deletion.docx` (`w:del` inside the
  box tolerated → `clean-open`) pin the rule's boundaries.

## Violation 1 — `word-textbox-deleted-text` / `bare-deltext-in-choice-textbox`

**Part:** `word/document.xml` (Choice-box story reported at lines 24833–24990; Fallback twin
at 25020–25177). **Sites: 12** (10 `w:delText` + 2 `w:delInstrText` in the `mc:Choice` box;
the Fallback's identical 12 are exempted by the check). *Line numbers and counts are the
analyst's; I could not re-read them — the fixture is absent.*

Offending XML as reported (document.xml line 24845, first of 12 in the Choice box):

```xml
<w:txbxContent>                       <!-- 24833, inside mc:Choice/w:drawing/wps:wsp/wps:txbx -->
  <w:p>
    <w:r>
      <w:rPr><w:sz w:val="16"/></w:rPr>
      <w:delText xml:space="preserve">© State of Western Australia</w:delText>
      <!-- run has no w:del ancestor inside this text box -->
    </w:r>
    …
    <w:delInstrText xml:space="preserve">DOCPROPERTY CommencementYear</w:delInstrText>  <!-- 24858 -->
```

**Schema rule — verified by me against the bundled XSD this session.** This is *not* an
XSD-validity defect: `EG_RunInnerContent` declares deleted text legal in any run, so the
schema cannot express the rule. `src/scripts/office/schemas/ISO-IEC29500-4_2016/wml.xsd:1745`:

```xml
<xsd:group name="EG_RunInnerContent">
  <xsd:choice>
    <xsd:element name="br" type="CT_Br"/>
    <xsd:element name="t" type="CT_Text"/>
    <xsd:element name="contentPart" type="CT_Rel"/>
    <xsd:element name="delText" type="CT_Text"/>       <!-- line 1748 -->
    <xsd:element name="instrText" type="CT_Text"/>
    <xsd:element name="delInstrText" type="CT_Text"/>  <!-- line 1750 -->
```

The intended containment is `wml.xsd:3406` — `<xsd:element name="del"
type="CT_RunTrackChange" minOccurs="0"/>` — whose content model (`CT_RunTrackChange`,
`wml.xsd:997-1006`) hosts `EG_ContentRunContent`. Deleted run content belongs inside
`w:del`; inside a text-box story Word additionally refuses *bare* deleted text at open time.
The operative constraint is Word's runtime behaviour, anchored here by the proven
assessment and by the repo's hash-bound fixtures above.

**Why the repo check owns it** (verified): `src/scripts/office/checks/word-deleted-text.ts:100`
registers `word-textbox-deleted-text` (error severity: "A text box contains deleted text.
Microsoft Word offers to recover the file."). One walk (`word-deleted-text.ts:48-73`) resets
the deleted flag when entering `w:txbxContent` (lines 57-59), sets it inside a `w:del` of
that story (60-61), and exempts `mc:Fallback` (62) — Word reads the sibling `mc:Choice`
(corpus fixture 2054, clean-open). Registered in `checks/index.ts:132`; per AGENTS.md note 17
the word rules run under word-valid and strict only, which matches the analyst's reported
runs: lenient → "All validations PASSED!"; `--profile word-valid` → exactly 12 ERROR lines,
all this code, all `word/document.xml`. **Not re-run this session** (fixture absent).

**Proposed fix.**
- *Artifact (proven minimal change, assessments.jsonl:35):* inside those 2 `w:txbxContent`
  elements only, rename `w:delText` → `w:t` (and `w:delInstrText` → `w:instrText`) — Word
  then opens clean. Everything outside the boxes, including the tolerated 1376 + 5
  body/header/footer deletions, stays.
- *Generator (the real fix):* docxodus must not emit bare deleted-run content inside
  `w:txbxContent`. When a tracked deletion lands in a text box, either wrap the runs in a
  `w:del` inside that story (tolerated: `working/package.word-textbox-inner-deletion.docx`,
  clean-open) or emit ordinary `w:t` / `w:instrText` there. The Fallback twin must be kept
  in sync with whatever the Choice branch carries.

**Fix target: docxodus** — the redline's XML is schema-legal but Word-refused; the producing
tool must change.

## Where the fixes belong and why

- **docxodus (generator bug).** The generator wrote tracked-deletion markup into text-box
  stories without the `w:del` wrapper the containment model intends. No schema change can
  catch it; only the emitter can stop producing it.
- **Not docx-validate.** The validator already detects exactly this: registered check
  `word-textbox-deleted-text` with unit tests (`tests/checks-word-deleted-text.test.ts`,
  `tests/word-textbox-deleted-text.test.ts`), broken/working fixture pairs, and hash-bound
  Word 16.115 observations. No missing check, no false positive.
- **Not a source-document repair target** (per this lane's ask): the document is tool
  output. The minimal-change experiment above is recorded as the *verification* repair, not
  as the deliverable fix.
- **Bookkeeping (no code change):** the fixture and its unpacked copy are missing from
  `really_broken/` while `docs/really-broken/assessments.jsonl:35` records it as proven, and
  the lane's `manifest.csv` does not exist. Re-materializing the fixture would let the site
  claim be confirmed or refuted here.

## Verification

1. **Re-materialize** `really_broken/tool_redlines/repair_prompt/docxodus__01a11a49_0162.docx`
   (from the corpus the assessment was recorded against), then run
   `pnpm exec tsx src/scripts/office/validate.ts really_broken/tool_redlines/repair_prompt/docxodus__01a11a49_0162.docx --profile word-valid`
   — expect exactly 12 error issues, code `word-textbox-deleted-text`, all in
   `word/document.xml` (Choice box only). **Not run this session: the file does not exist.**
2. **Minimal-change experiment (already proven, assessments.jsonl:35):** rename the 24
   deleted-text-class elements to `w:t` / `w:instrText` inside the 2 text boxes only →
   `clean-open` (original: `unreadable-content-warning`). Per repo methodology, two separate
   Word runs of the same bytes, verdicts read from each run's log.
3. **Existing in-repo coverage pinning the defect class:** the unit tests above plus
   `tests/fixtures-all.manifest.json`'s hash-bound Word observations
   (`unreadable-content-warning` for the broken shapes, `clean-open` for the repaired,
   Fallback-tolerated and inner-`w:del` controls) — all read this session.
