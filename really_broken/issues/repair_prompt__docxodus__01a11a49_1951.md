# [docxodus 1951] repair-prompt: bare w:delText inside text boxes makes Word offer recovery

## Summary

`tool_redlines/repair_prompt/docxodus__01a11a49_1951.docx` (sha256 `315615c10288cc354b8d71c8a32f0ccc75f1b0a548501dd303282132a4a3b5b1`, matching the origin manifest row) is a tracked-changes redline produced by **docxodus** (all revision marks authored `Open-Xml-PowerTools`, fetched from `https://jubarte.pro/live/artifacts/01a11a49-…/1951/docxodus-redline.docx`). It is a letterhead/memo document whose letterhead blocks are DrawingML/VML text boxes.

- **Live Word outcome (this run):** `repair-prompt` — Word showed *"Word found unreadable content in docxodus__01a11a49_1951.docx. Do you want to recover the contents of this document?"*; answering **No** was followed by *"Word experienced an error trying to open the file"* and no document window ever opened.
- **Manifest row** (`really_broken/manifest.csv`, origin copy at `neurotic_docx_bench/maybe_future_corpus/`): class `word:corrupt`, `Word could not load it in 2 passes: document loaded empty (Word could not read it)` — a cruder two-pass loader observation of the same refusal; the live run pins the precise dialog. The two agree (`agrees: true`).
- **docx-validate** (word-valid profile, run for this issue — see Verification): exit 1, **29 ERROR lines, all code `word-textbox-deleted-text`** (3 `word/document.xml`, 18 `word/footer1.xml`, 8 `word/footer2.xml`).

The dominant defect: the generator marked text inside text-box stories as deleted (`w:delText`) while the `w:del` wrapper stays **outside** the box around the drawing run. A text box is its own story; bare `w:delText` inside it makes Word refuse the whole part.

## Word behaviour

- **Live result:** recovery prompt, then open error; no crash report; Word quit cleanly (above).
- **Prior assessment: PROVEN** — `docs/really-broken/assessments.jsonl` and `inspection-log.jsonl` (source `word-ab`), check `word-textbox-deleted-text`. Experiment matrix, each row verified in Word:
  - original → `unreadable-content-warning`
  - repair footer1+footer2, leave `word/document.xml` unrepaired (`1951-leave-document.docx`) → still warns
  - repair document+footer2, leave `word/footer1.xml` unrepaired (`1951-leave-footer1.docx`) → still warns
  - repair document+footer1, leave `word/footer2.xml` unrepaired (`1951-leave-footer2.docx`) → still warns
  - **inside the text boxes of all three parts, rename `w:delText` → `w:t` (`1951-txbx-all.docx`) → `clean-open`, validator then valid.** The 61 out-of-box `w:delText` in `word/document.xml` stayed and the file opened clean — they are Word-tolerated.

So the Word cause is proven and each of the three parts is independently sufficient to trigger the prompt.

## Each violation

All three violations are the same defect in different parts; the check is `word-textbox-deleted-text` (`src/scripts/office/checks/word-deleted-text.ts:99-127`). Its header comment documents the Word rule with fixtures: *"w:delText or w:delInstrText inside w:txbxContent makes Word offer recovery when no w:del inside that text box wraps it… A w:del inside the text box is tolerated (fixture 0948). A w:del outside the text box does not cover the text-box story."* (`word-deleted-text.ts:7-14`). The walker confirms this: entering `txbxContent` resets the deleted flag (`word-deleted-text.ts:57-59`) and `mc:Fallback` content is skipped (`:62`), so the 29 errors are the `mc:Choice`/effective-view sites; the Fallback twins are counted below as sites.

### 1. `word/document.xml` — 6 in-box `w:delText` (3 flagged + 3 Fallback twins)

4 `w:txbxContent` elements (2 `mc:AlternateContent` × Choice-`wps:txbx` + Fallback-`v:textbox`), 6 in-box `w:delText` (Choice boxes: 2 + 1; Fallback boxes: 2 + 1). Zero `w:del` inside any box; the `w:del` (id=98 and id=102, author `Open-Xml-PowerTools`) wraps the drawing run **outside** the box, on the chain `txbxContent < wsp < graphicData < graphic < anchor < drawing < Choice < AlternateContent < … < w:del`. First box (real bytes, whitespace compacted):

```xml
<w:txbxContent>
  <w:p w14:paraId="01DD47AF" w14:textId="5921FAC1">
    <w:pPr><w:pStyle w:val="returnaddress"/><w:spacing w:line="240" w:lineRule="auto"/></w:pPr>
    <w:r><w:delText>Project Development Branch </w:delText></w:r>
  </w:p>
  …
</w:txbxContent>
<!-- context: the whole box sits inside <w:del w:author="Open-Xml-PowerTools" w:id="98"><w:r><mc:AlternateContent><mc:Choice Requires="wps"><w:drawing>… -->
```

### 2. `word/footer1.xml` — 36 in-box `w:delText` (18 flagged + 18 Fallback twins)

2 `w:txbxContent` (letterhead address block), 18 + 18 in-box `w:delText`; the part's total `w:delText` count is 36 — **every deleted text in this part is inside a box**. No `w:del` inside either box; the `w:del` (id=214) wraps the run outside. First box:

```xml
<w:txbxContent>
  <w:p w14:paraId="197430CB" w14:textId="615531BA">
    <w:pPr><w:pStyle w:val="returnaddressbottom"/><w:jc w:val="left"/></w:pPr>
    <w:r><w:delText>4201 East Arkansas Ave. </w:delText></w:r>
    <w:r><w:delText>, </w:delText></w:r>
    <w:r><w:delText xml:space="preserve">4 </w:delText></w:r>
    <w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:delText>th </w:delText></w:r>
  </w:p>
  …
</w:txbxContent>
```

### 3. `word/footer2.xml` — 16 in-box `w:delText` (8 flagged + 8 Fallback twins)

Same shape: 2 boxes, 8 + 8 in-box `w:delText`, all of the part's deleted text, `w:del` (id=232) outside the boxes. Sample:

```xml
<w:txbxContent>
  <w:p w14:paraId="7556801E" w14:textId="77777777">
    <w:pPr><w:pStyle w:val="returnaddressbottom"/></w:pPr>
    <w:r><w:delText>555 Street Address, Room 555, Denver, CO 55555-5555 </w:delText></w:r>
    …
  </w:p>
</w:txbxContent>
```

### Schema rule — and why this is behavioural, not an XSD error

Quoted from the bundled `src/scripts/office/schemas/ISO-IEC29500-4_2016/wml.xsd`:

```xml
<xsd:group name="EG_RunInnerContent">          <!-- wml.xsd:1745 -->
  <xsd:choice>
    <xsd:element name="t" type="CT_Text"/>     <!-- :1748 -->
    <xsd:element name="delText" type="CT_Text"/> <!-- :1750 -->
    …
</xsd:group>

<xsd:complexType name="CT_TxbxContent">        <!-- wml.xsd:3378 -->
  <xsd:group ref="EG_BlockLevelElts" minOccurs="1" maxOccurs="unbounded"/>
</xsd:complexType>
<xsd:element name="txbxContent" type="CT_TxbxContent"/> <!-- :3381 -->
```

`EG_BlockLevelElts` → `w:p` → `CT_R` → `EG_RunInnerContent` (`wml.xsd:1782-1786`), so **the schema deliberately permits `w:delText` inside a text box**; this construct draws no `xsd-error`. The rule being broken is Word's story boundary — a deletion in the main story does not cover a text-box story — which is exactly what the repo's `word-textbox-deleted-text` check encodes and Word's recovery prompt proves. No `w:delInstrText` exists anywhere in the package, and no other part (headers, footnotes, endnotes) holds a text box.

### Proposed fix — target: **docxodus** (all three violations)

When marking deletions, docxodus descends into `w:txbxContent` (and the VML `v:textbox` twin) of a deleted drawing run and renames the box's `w:t` to `w:delText`, leaving the `w:del` wrapper outside the box. It must not do that. Either:

1. keep `w:t` inside `w:txbxContent`/`v:textbox` untouched (the whole drawing is deleted by the outer `w:del`; the box's story needs no deletion marks) — this is the **proven** repair for this file; or
2. if deleted box text must be shown as such, wrap those runs in their **own** `w:del` *inside* the box story (own id, same author/date) — tolerated by Word per fixture 0948 (`word-deleted-text.ts:13`), though not re-proven on this specific file.

The 61 out-of-box `w:delText` in `word/document.xml` are correct redline markup and must be left alone.

## Where the fixes belong and why

- **docxodus** — the redline's XML is genuinely invalid-for-Word: the generator (Open-Xml-PowerTools-based, per the revision author) wrote deletion marks into a story that its `w:del` does not cover. All three parts come from one code path, so one generator change fixes all 58 in-box sites (29 Choice + 29 Fallback twins).
- **Not docx-validate** — the validator already reports every flagged site with the right code and severity (29 errors, exit 1); no missing check, no false positive, nothing to change.
- **Not `repair`** — the source document is not at fault; this lane judges the tool's output, and the output is what Word refuses.

## Verification

- **Validator re-run (this issue):** `pnpm exec tsx src/scripts/office/validate.ts really_broken/tool_redlines/repair_prompt/docxodus__01a11a49_1951.docx` → exit 1, exactly 29 ERROR lines all `word-textbox-deleted-text` (3 / 18 / 8 by part); matches the prior assessment's `validator_error_codes`.
- **Independent XML walk (this issue):** a Python ElementTree walk mirroring `word-deleted-text.ts` semantics over the freshly unpacked parts (`pnpm exec tsx src/scripts/office/unpack.ts … --merge-runs false --simplify-redlines false`) reproduced every count above: 6/36/16 in-box `w:delText`, 0 in-box `w:del`, the outside `w:del` ids 98/102/214/232, 61 out-of-box `w:delText` in `document.xml`, 0 `w:delInstrText`, text boxes in exactly the three parts.
- **Minimal-change experiment (already proven, `inspection-log.jsonl`):** renaming in-box `w:delText` → `w:t` in all three parts (`1951-txbx-all.docx`) → Word `clean-open` and docx-validate valid; each single-part-left-unrepaired variant still shows the recovery prompt, so a partial generator fix is not enough. For a docxodus-side fix, the acceptance test is: regenerate this redline, confirm 0 `word-textbox-deleted-text` findings and a clean open, with the out-of-box deletions intact.
- Note on the analyst log: two early bare validator runs in the analyst's session printed `All validations PASSED!`; the run for this issue (git tree at `2a98048d`) reproduced the 29-error result, as did the analyst's own hash-verified re-runs, so the 29-error result stands.
