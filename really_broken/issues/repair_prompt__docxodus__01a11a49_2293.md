# [docxodus 2293] repair-prompt: bare `w:delText` in 370 text-box sites + `xml:space` on `w:r`

## Summary

- Specimen: `really_broken/tool_redlines/repair_prompt/docxodus__01a11a49_2293.docx` — a docxodus redline of a French-language document, sha256 `426506983b5f3a51131677797d79d8c626b96d2be6edc29a5cf607e16675d259`, restored byte-identical to the corpus copy (`cmp` clean, this session). Tracked-change inventory of `word/document.xml` (my ElementTree count): 741 `w:delText`, 23 `w:del`, 296 `w:ins w:author="Open-Xml-PowerTools"` (0 in other parts; the analyst note's "320" is off — actual is 296).
- Live Microsoft Word (result given in the ask): **repair-prompt** — "Word found unreadable content in docxodus__01a11a49_2293.docx. Do you want to recover…?"; No → "Word experienced an error trying to open the file", no document opened, no crash report.
- Manifest: no row exists. `really_broken/manifest.csv` does not exist in the repo (searched; the only manifest is `tests/lane/jubarte-0.13.0/manifest.csv`) and the live-result payload carries `manifestSays: ""`. So there is no manifest verdict to agree or disagree with; the live outcome is the sole Word evidence.
- docx-validate, default word-valid profile (run this session): exit 1, 740 output lines = 370 `ERROR` lines on `word/document.xml`, 0 WARN/INFO lines in the text output. `--json` codes: `word-textbox-deleted-text` ×370, plus advisory `paraid-missing-element` ×5, `xml-bom-leading` ×2, info `tracked-change-ids-regenerated`, info `rel-ids-sequential`. Lenient profile: exit 1, one `xsd-error` + one `xsd-summary` ("word/document.xml: 1 new error(s)") for the `xml:space`-on-`w:r` defect.

## Word behaviour

- Live result (from the ask): repair-prompt, dialog text quoted above; clicking No produced the open-error dialog; Word quit cleanly.
- Prior assessment for **this file**: none. 0 hits for `docxodus__01a11a49_2293` in `docs/really-broken/assessments.jsonl` (144 rows) and `docs/really-broken/inspection-log.jsonl`.
- The Word cause of the textbox defect is nonetheless proven **for this exact producer shape** by sibling rows in `docs/really-broken/assessments.jsonl` (all status `proven`):
  - fixture **0044**: `mc:Choice` wps:txbx + `mc:Fallback` v:textbox twin, 25 `w:delText` each → `both_text_boxes_delText_changed_to_w_t` = **clean-open**; body delText tolerated in the repair; either copy alone still warned.
  - fixture **2054**: repairing only the four Choice boxes = **clean-open** while the Fallback boxes kept their delText; repairing only the Fallbacks still warned (Word reads the Choice).
  - fixture **0011**: VML textbox `w:delText`/`w:delInstrText` warns; removing it opens clean; 10,018 body `w:delText` tolerated.
- The repo check `src/scripts/office/checks/word-deleted-text.ts` (lines 8–17) cites exactly these fixtures. For this file the chain is therefore: shape-proven, file-unproven — the minimal-change experiment below closes that gap.

## Violation 1 — `word-textbox-deleted-text`: 370 bare `w:delText` inside Choice-side text boxes

**Reproduced (this session).** `word/document.xml:2139–2156`, inside the first `mc:AlternateContent` (line 2092):

```xml
<mc:AlternateContent>
  <mc:Choice Requires="wps">
    …<wps:txbx>
      <w:txbxContent>
        <w:p>…<w:r>
          <w:rPr><w:rFonts w:ascii="Arial Narrow" w:hAnsi="Arial Narrow"/><w:b/><w:bCs/></w:rPr>
          <w:delText>CARACTÉRISTIQUES PHYSIQUES
          </w:delText>
        </w:r></w:p>
```

- My independent ancestry-tracking walk over the unpacked part: 17 `mc:AlternateContent` (line 2092 ff.), each `mc:Choice Requires="wps"` (wps:txbx) + `mc:Fallback` (v:textbox) twin; the 17 Choice-side boxes hold **370 `w:delText`** with **0 `w:del`** inside any box; the 17 Fallback twins hold the mirrored 370 (exempt — Word reads the Choice, fixture 2054). The walk matches the validator's 370 exactly. `word/footer1.xml`'s 3 `w:delText` sit outside any textbox (tolerated body content, per 0011).
- Schema rule — this is **not** an XSD error. Bundled `src/scripts/office/schemas/ISO-IEC29500-4_2016/wml.xsd`, group `EG_RunInnerContent` (line 1746 ff.): `<xsd:element name="delText" type="CT_Text"/>` (line 1750) — `delText` is legal run content anywhere; the schema cannot express Word's requirement that a text-box story carry its own `w:del`. The rule is the Word-behaviour check `word-deleted-text.ts:100` (`word-textbox-deleted-text`), Word-proven on this shape by 0044/2054/0011.
- **Proposed fix — target `docxodus`:** never emit `w:delText`/`w:delInstrText` inside a `w:txbxContent` run unless a `w:del` wraps it within that text box. Either (a) emit `w:t` in the text-box story — the proven repair (0044: delText→w:t opens clean); keep the Fallback twin consistent with the Choice; or (b) wrap the runs in a `w:del` inside the box (tolerated, fixture 0948).

## Violation 2 — `xml:space="preserve"` on 81 `w:r` elements (`CT_R` does not declare it)

**Reproduced (this session).** `word/document.xml:125–128`, inside `<w:ins w:author="Open-Xml-PowerTools" w:id="16" …>` (line 124):

```xml
<w:r xml:space="preserve">
  <w:t>
  </w:t>
</w:r>
```

- Raw part counts: 81 `<w:r xml:space="preserve">` (first at line 14); 408 `xml:space` attributes total, the other 327 on `w:t` (135) and `w:delText` (192) where `CT_Text` declares it — only the 81 on `w:r` are illegal.
- Schema rule, quoted from the bundled `wml.xsd`:
  - `CT_R` (lines 1782–1790): content is `<xsd:group ref="EG_RPr" …/>` + `<xsd:group ref="EG_RunInnerContent" …/>`; its only attributes are `rsidRPr`, `rsidDel`, `rsidR` (lines 1787–1789) — no `xml:space`, no `anyAttribute`.
  - `CT_Text` (lines 1738–1742): `<xsd:attribute ref="xml:space" use="optional"/>` (line 1741) — the sole declaration.
- Commands run: `xmllint --noout --schema …/wml.xsd word/document.xml` → 99 lines: **81 × "attribute '{http://www.w3.org/XML/1998/namespace}space' is not allowed" on `r`**, 17 × `mc:AlternateContent` "not expected" (MCE artifact — raw xmllint does no branch selection; the repo validates after markup-compatibility preprocessing and its lenient schema walk reports only the `xml:space` error), 1 summary line. `validate.ts --profile lenient --json`: one `xsd-error` ("…attribute …space… is not allowed") + one `xsd-summary`.
- Word effect for this file: **unproven** — no A/B evidence; the live repair-prompt is attributed to the shape-proven textbox defect. The attribute is schema-invalid regardless (strict emits each site as its own `xsd-error`).
- **Proposed fix — target `docxodus`:** drop `xml:space` from `w:r` elements; if whitespace preservation is intended, put it on the `w:t`/`w:delText` child, where `CT_Text` declares it.

## Where the fixes belong and why

- **`docxodus` (both violations):** the redline generator wrote tracked-change markup Word refuses (bare text-box `w:delText`, its shape Word-proven) and schema-invalid run attributes (`xml:space` on `w:r`). Two generator changes: text-box story emission (emit `w:t` or wrap in `w:del` inside the box) and the run serializer (never place `xml:space` on `w:r`).
- **Not docx-validate:** the check exists and fired exactly (`word-textbox-deleted-text` ×370, correct part, correct Choice/Fallback/`w:del` exemptions); the lenient/strict schema walk already reports the `CT_R` attribute error; word-valid runs no XSD by documented design (AGENTS.md notes 4 and 15), so the default profile emitting no code for violation 2 is intended behaviour, not a gap.
- **Not repair:** this is generator output, not a broken source document.

## Verification

1. Repro (run this session, from the repo root):
   - `pnpm exec tsx src/scripts/office/validate.ts really_broken/tool_redlines/repair_prompt/docxodus__01a11a49_2293.docx` → exit 1, 370 ERROR lines (740 with detail lines); `--json` codes as in Summary.
   - `pnpm exec tsx src/scripts/office/validate.ts … --profile lenient --json` → exit 1, `xsd-error` ×1 + `xsd-summary` ×1.
   - `xmllint --noout --schema src/scripts/office/schemas/ISO-IEC29500-4_2016/wml.xsd really_broken/unpacked/tool_redlines__repair_prompt__docxodus__01a11a49_2293/word/document.xml` → 99 lines (81 / 17 / 1 as above).
2. Minimal-change experiment to prove this file's cause (it has no assessment row): copy the docx, rewrite the 370 Choice-side text-box `w:delText` to `w:t` — leave the Fallback twins, the body tracked changes and the 81 `xml:space` untouched — and open in Word twice with logs per the repo's two-run rule. Expected **clean-open** (this is exactly 2054's `choice_boxes_repaired`). If it still warns, strip the 81 `xml:space` from `w:r` in a second specimen to isolate violation 2's contribution.
3. On a clean-open result, record the row in `docs/really-broken/assessments.jsonl`; if the repaired specimen joins `tests/fixtures/`, name it per the `<subject>.<comment-or-error>.docx` convention and refresh the manifest.
