# [docxodus 2300] repair-prompt: bare w:delText in footer text boxes triggers Word recovery

## Summary

- **Document**: `really_broken/tool_redlines/repair_prompt/docxodus__01a11a49_2300.docx`
  (sha256 `fd274a732fc9daab280cee93db1fcf4841801532d0a93cc58648a534ff7764b8`, verified this
  session with `shasum -a 256`). A docxodus / Open-Xml-PowerTools redline of a Slovak municipal
  regulation. Each of the three footers carries a classification text box ("Internal data, not
  suitable for public channels.") whose deleted text became bare `w:delText` in the text-box story.
- **Defect (3 sites, all re-verified)**: in `word/footer1.xml`, `word/footer2.xml` and
  `word/footer3.xml`, one `w:delText` sits inside the `mc:Choice` `wps:txbx` → `w:txbxContent`
  with **no `w:del` inside that text box** wrapping it. The only enclosing `w:del` opens at
  footer-body level, outside the box — a `w:del` outside the box does not cover the story.
- **Live Word vs manifest**: live Word showed the recovery dialog; answering **No** was followed
  by "Word experienced an error trying to open the file" and no document window
  (`outcome: repair-prompt`, `agrees: true`). `manifestSays` is empty and
  `really_broken/manifest.csv` does not exist on disk (checked this session).
- **Validator agreement**: `word-textbox-deleted-text` fires once per footer at the default
  profile — exactly the three sites the recovery prompt corresponds to.

## Word behaviour

- **Live result this run** (from the workflow's Word probe): recovery prompt → open error → no
  document; no crash report; Word quit cleanly. `agrees: true`.
- **Prior assessment for this file**: none — `grep docxodus__01a11a49_2300` over
  `docs/really-broken/assessments.jsonl` and `inspection-log.jsonl` returns 0 rows (checked this
  session). The exact-file cause is therefore not individually proven; the evidence below is the
  class-level proof in four sibling docxodus specimens, all status `proven`.
- **Class-level proof** (`docs/really-broken/assessments.jsonl`):
  - **line 29 — 0011**: VML `v:textbox` in a footer with one `w:delText` (or `w:delInstrText`)
    → recovery; removing the deleted text from the box → `clean-open`; the body's 10,018
    `w:delText` are not the cause.
  - **line 30 — 0044**: this file's exact shape — `mc:Choice` `wps:txbx` + `mc:Fallback`
    `v:textbox` twins each holding `w:delText` → recovery; repairing the Choice copy opens
    clean; repairing only the Fallback copy still warns.
  - **line 62 — 0948**: 20 `w:delText` + 4 `w:delInstrText` inside 2 `w:txbxContent` →
    recovery; renaming to `w:t`/`w:instrText` inside those boxes only → `clean-open`.
  - **line 111 — 2054**: four `mc:Choice` boxes each with one bare `w:delText` → recovery;
    repairing only the Choice boxes → `clean-open` while the `mc:Fallback` twins keep theirs.
- **This repo's check** (`word-textbox-deleted-text`,
  `src/scripts/office/checks/word-deleted-text.ts:99-127`; header `:6-19` cites fixtures
  0011/0044/2054/0948) encodes exactly this rule. It is Word-blocking
  (`validators/word-blocking.ts:293`), explained at `validators/word-error-explanations.ts:97-98`.

## Each violation

All three sites share one anatomy (verified by reading the parts and by an independent Python
walk replicating the check's state machine): the enclosing `w:del` wraps the drawing-carrying
run at footer-body level, the `mc:Choice` `wps:txbx` box's inner text became `w:delText`, and no
`w:del` opens inside that `w:txbxContent`.

### 1. `word/footer1.xml:77` — bare `w:delText` in the Choice text box (`word-textbox-deleted-text`)

```xml
<!-- footer1.xml:10 — the only w:del opens OUTSIDE the text box, at footer-body level -->
<w:del w:author="Open-Xml-PowerTools" w:id="1101" w:date="2000-01-01T00:00:00Z">
  <w:r><mc:AlternateContent><mc:Choice Requires="wps">
    <w:drawing><wp:anchor …>…<wps:wsp>…<wps:txbx>
      <w:txbxContent>                     <!-- line 57 -->
        <w:p …>…<w:r>…<w:rPr>…</w:rPr>
          <w:delText>Internal data, not suitable for public channels.</w:delText>
          <!-- line 77: no w:del anywhere inside this txbxContent -->
```

`</w:del>` closes at line 135, after the whole `mc:AlternateContent` — the `del` is outside the
`txbxContent`.

### 2. `word/footer2.xml:255` — same shape, enclosing `w:del` id 1069 at line 188

Identical anatomy: `w:del` (id 1069) opens at line 188 wrapping the run; `w:txbxContent` spans
235–259; the bare `w:delText` is at line 255; `</w:del>` at 313.

### 3. `word/footer3.xml:77` — same shape, enclosing `w:del` id 1081 at line 10

Identical to footer1: `w:del` (id 1081) at line 10, `w:txbxContent` 57–81, bare `w:delText` at
line 77, `</w:del>` at 135. All three sites carry check code `word-textbox-deleted-text`.

### Schema rule — and why this is not an XSD error

Read this session from `src/scripts/office/schemas/ISO-IEC29500-4_2016/wml.xsd`:

```xml
<!-- wml.xsd:1745-1752 -->
<xsd:group name="EG_RunInnerContent">
  <xsd:choice>
    <xsd:element name="t" type="CT_Text"/>
    <xsd:element name="delText" type="CT_Text"/>   <!-- permitted in ANY w:r -->
    <xsd:element name="delInstrText" type="CT_Text"/>
…
<!-- wml.xsd:3378-3381 -->
<xsd:complexType name="CT_TxbxContent">
  <xsd:group ref="EG_BlockLevelElts" minOccurs="1" maxOccurs="unbounded"/>
</xsd:complexType>
<xsd:element name="txbxContent" type="CT_TxbxContent"/>
```

XSD 1.0 cannot express the co-occurrence constraint, so the construct is schema-valid and the
defect is a **Word-runtime tracked-change rule**: deleted text in a text-box story must be
covered by a `w:del` inside that same text box (`word-deleted-text.ts:14`, `:57-64` — the walk
resets `deleted = false` on entering `txbxContent`, line 59). Confirmed this session with
`xmllint --noout --schema …/wml.xsd …/word/footer1.xml`: only known mc/w14 noise, **nothing
about the `w:delText`**.

### Tolerated twins (not violations — excluded by design)

Each footer's `mc:Fallback` VML `v:textbox` holds the same `w:delText` (footer1:123, footer2:301,
footer3:123); the check excludes `mc:Fallback` copies (`word-deleted-text.ts:62`), matching
2054's proof that Word uses the sibling `mc:Choice`. footer2's body-level `w:delInstrText` /
`w:delText` (ids 1070–1074, 1078) and the 336 body-level `w:delText` in `word/document.xml` all
sit inside proper `w:del` elements (0 offenders there).

### Proposed fix (per site)

docxodus must not emit bare `w:delText` inside a `w:txbxContent`. For each of the three Choice
text boxes, either **(a)** wrap the in-box runs in a `w:del` inside that `w:txbxContent`
(preserves the deletion marking; documented tolerated at `word-deleted-text.ts:13`), or
**(b)** rename the in-box `w:delText` → `w:t` — the recorded class-proven minimal repair
(0948/2054: renaming inside the Choice boxes only → `clean-open`; Fallback twins may stay).

**Fix target**: `docxodus` (generator bug), all three sites.

## Where the fixes belong and why

- **docxodus (generator)**: the redline wrapped the drawing-carrying run in a footer-body
  `w:del` and converted the text box's inner text to `w:delText`, but the text-box story needs
  its own in-box `w:del`. This is the only code that must change; the same generator defect is
  already recorded for siblings 0011/0044/0070/0948/2054.
- **docx-validate — no gap**: the check exists, is Word-blocking, fires at the default profile,
  reports exactly the three sites and no tolerated twin, and agrees with the live repair-prompt
  outcome.
- **repair — not applicable**: generator output, not a broken source document.

## Verification

1. **Validator reproduction (run this session)**:
   `pnpm exec tsx src/scripts/office/validate.ts really_broken/tool_redlines/repair_prompt/docxodus__01a11a49_2300.docx`
   → exit 1, exactly 3 ERROR lines, all `word-textbox-deleted-text`, one per footer part.
2. **Independent site walk (run this session)**: a Python script replicating the check's state
   machine over the unpacked tree found exactly one offender per footer (path ending
   `…/Choice/drawing/anchor/graphic/graphicData/wsp/txbx/txbxContent/p/r/delText`) and 0 in
   `document.xml`, `header1–4.xml`, `footnotes.xml`, `endnotes.xml` — matching the CLI exactly.
3. **Schema non-error (run this session)**: `xmllint --noout --schema …/wml.xsd` on footer1
   reports no `delText` finding (only known mc/w14 noise), confirming the behavioural-rule
   character.
4. **Minimal-change experiment for this exact file (not yet run — no live Word here)**: in the
   three `mc:Choice` `w:txbxContent` only, rename `w:delText` → `w:t` (Fallback twins
   untouched) and open in Word twice. Expected: clean open, by 2054/0948's class proof. Until
   this (or the A/B twin: wrapping each in-box run in a `w:del` inside the box) runs, the
   exact-file cause rests on the class proof plus the live repair-prompt and exact site
   agreement. *(Not run this session: a live Word open of a repaired variant; `--verify-word`.)*
