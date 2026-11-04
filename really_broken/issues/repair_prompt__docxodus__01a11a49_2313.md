# [docxodus 2313] repair-prompt: w:delText inside a wps text box no w:del covers

## Summary

`tool_redlines/repair_prompt/docxodus__01a11a49_2313.docx` is a docxodus redline (tracked-changes output, author `Open-Xml-PowerTools`) over a UK programme-specification document. Live Microsoft Word refuses it: the recovery dialog `Word found unreadable content in docxodus__01a11a49_2313.docx` appears, and answering No yields `Word experienced an error trying to open the file` with no document window (live outcome `repair-prompt`, run 2026-11-04).

`docx-validate` (default word-valid, command `pnpm exec tsx src/scripts/office/validate.ts really_broken/tool_redlines/repair_prompt/docxodus__01a11a49_2313.docx --json`, run during this assessment) reports exactly **1 error**: `word-textbox-deleted-text` on `word/document.xml`. No manifest row exists for this path — `really_broken/manifest.csv` is absent from the repo (grep: no such file), so `manifestSays` is empty; the live result and the validator agree.

One violation, re-verified from the unpacked part, the check source and the bundled XSD: a `w:delText` inside the `w:txbxContent` of a DrawingML text box, not covered by any `w:del` within that text-box story. The placement is **XSD-legal** — the block is behavioural, and the producing tool must change.

## Word behaviour

- **Live (this run, from the ask):** recovery dialog, then open error after No; no crash report; Word quit cleanly. Outcome `repair-prompt`, `agrees: true`.
- **Manifest / prior assessment:** none. `really_broken/manifest.csv` does not exist; grep of `docs/really-broken/assessments.jsonl` and `inspection-log.jsonl` for `docxodus__01a11a49_2313` returns no row. The cause is therefore **not yet proven for this exact file**.
- **Family evidence (same tool, same lane, proven):** `docs/really-broken/assessments.jsonl` rows 29 and 30:
  - row 29 (`docxodus__01a11a49_0011.docx`): a VML `v:textbox` holding one `w:delText` (or one `w:delInstrText`) makes Word offer recovery; removing it opens clean; 10,018 body-level `w:delText` are *not* the cause.
  - row 30 (`docxodus__01a11a49_0044.docx`): the **same shape as this file** — `mc:Choice` wps text box plus `mc:Fallback` VML twin. Proven verdicts: original warns; changing `w:delText`→`w:t` in **both** `w:txbxContent` copies opens clean; changing only the VML copy, or removing only one branch, still warns.

## Violation — `word-textbox-deleted-text` (the only one; confirmed)

**Offending XML** — `word/document.xml` (unpacked at `really_broken/unpacked/repair_prompt__docxodus__01a11a49_2313`, restored this session from the byte-identical bench copy):

```xml
<w:del w:author="Open-Xml-PowerTools" w:id="1079" ...>   <!-- :12254, body — deletes the whole shape -->
  <w:r>
    <mc:AlternateContent>
      <mc:Choice Requires="wps">                          <!-- :12263 -->
        <w:drawing>…<wps:wsp>…<wps:txbx>
          <w:txbxContent>                                 <!-- :12308 -->
            <w:p w14:paraId="3E3DD5A2" …>
              <w:r>…<w:delText>BEng (Hons) Civil Engineering achieved
                </w:delText>                              <!-- :12321 — the defect -->
              …
            </w:r>
          </w:txbxContent>                                <!-- closes :12325; zero w:del inside -->
      </mc:Choice>
      <mc:Fallback>                                       <!-- :12347 -->
        <w:pict>…<v:textbox><w:txbxContent>               <!-- :12356 -->
          …<w:delText>BEng (Hons) Civil Engineering achieved</w:delText>  <!-- :12369, tolerated -->
```

Verified this session: `grep -c 'w:del '` over both `w:txbxContent` spans (12308–12325, 12356–12373) returns **0** — no deletion wraps the text-box story. The twin in `mc:Fallback` is excluded by the check's `inFallback` tolerance because Word reads the sibling `mc:Choice`. Every other story part is clean of the shape (`header1–3.xml`, `footnotes.xml`, `endnotes.xml`: no `txbxContent`; `footer1.xml` has 34 `w:delText` + 4 `w:delInstrText`, all body-level, which row 0011 proved Word tolerates).

**Schema rule.** `w:delText` inside a run is legal anywhere a run goes — `src/scripts/office/schemas/ISO-IEC29500-4_2016/wml.xsd:1745`:

```xml
<xsd:group name="EG_RunInnerContent">
  <xsd:choice>
    <xsd:element name="t" type="CT_Text"/>
    <xsd:element name="delText" type="CT_Text"/>      <!-- :1750 — ordinary run content -->
    …
```

and `wml.xsd:3378`:

```xml
<xsd:complexType name="CT_TxbxContent">
  <xsd:group ref="EG_BlockLevelElts" minOccurs="1" maxOccurs="unbounded"/>
</xsd:complexType>
<xsd:element name="txbxContent" type="CT_TxbxContent"/>
```

So the XSD does not confine `w:delText` to `w:del`. Confirmed empirically: `--strict --json` (run this session) emits **zero `xsd-error` findings** (7 `xsd-no-schema` info; the 2 `paraid-missing-element` errors under strict are the w14 extension-attribute noise, 72 `w:p` + 22 `w:tr`). The Word block is the behavioural check `word-textbox-deleted-text` — `src/scripts/office/checks/word-deleted-text.ts:99-127`, walk at `:48-75`: entering `w:txbxContent` resets `deleted` (`:58`), which encodes the rule stated at `:14` — *"A w:del outside the text box does not cover the text-box story."* The check fired exactly once on this file, at this site (the fallback twin at 12369 is tolerated by `inFallback`).

**Proposed fix (generator).** In docxodus's redline emitter: when the whole text box is deleted by the enclosing body-level `w:del` (as here — `w:id="1079"` deletes the shape run), the runs inside `w:txbxContent` must carry plain `w:t` / `w:instrText`, not `w:delText` / `w:delInstrText`, in **both** the `mc:Choice` and `mc:Fallback` copies. If a future case deletes only text *inside* a surviving text box, the emitter must instead open a `w:del` within that `w:txbxContent` story.

**Fix target:** `docxodus` (the producing tool). The redline's XML is schema-valid but behaviourally rejected by Word; the generator must change.

## Where the fixes belong and why

- **docxodus** — the single violation is a generator defect: it propagated the body deletion into the nested text-box story as `w:delText`, in both compatibility branches. docx-validate already detects it (`word-textbox-deleted-text`, error, Word-blocking), so no validator change is needed and there is no false positive to dispute.
- **Not docx-validate** — re-verified: the check, the XSD lines and the site all reproduce; the strict run confirms schema legality, so this is not an XSD matter.
- **Not repair** — this is generated redline output, not a broken source document; no repair prescription beyond the generator change.

## Verification

1. **Minimal-change experiment to prove the cause for this file** (not yet proven — no assessment row exists): rename `w:delText` → `w:t` at `word/document.xml:12321` **and** `:12369` (both branches; row 0044 shows changing one copy still warns), repack, and re-run `pnpm exec tsx src/scripts/office/validate.ts <repaired.docx> --verify-word`. Expected: the single `word-textbox-deleted-text` error disappears and Word opens clean. Record the two-run Word observation per `docs/really-broken/` practice.
2. **Existing coverage** — the rule is already pinned by `tests/word-textbox-deleted-text.test.ts`, `tests/checks-word-deleted-text.test.ts` and fixtures `tests/fixtures/broken/package.word-textbox-deleted-text.docx` and `package.word-textbox-choice-deleted-text.docx` (the mc:Choice/Fallback shape). A docxodus-side regression test should assert no `w:delText`/`w:delInstrText` survives inside any emitted `w:txbxContent` unless a `w:del` opens inside that story.
3. **This session's runs** (for reproduction): default profile → 1 error `word-textbox-deleted-text`, 2 `xml-syntax` warnings (`customXml/item2.xml`, relative xmlns URI `aecd4273-0d56-430f-bd52-977836de9101`), 2 `paraid-missing-element` warnings, 2 `rel-ids-sequential` info — none Word-blocking; `--strict` → same textbox error, paraid escalated to error (extension noise), zero `xsd-error`.
