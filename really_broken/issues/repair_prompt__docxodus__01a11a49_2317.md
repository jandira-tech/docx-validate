# [docxodus 2317] repair-prompt: bare `w:delText` in text-box stories makes Word refuse the file

## Summary

`tool_redlines/repair_prompt/docxodus__01a11a49_2317.docx` (sha256
`71e52fdd5b120f6f8ade48d1a2c337f579e8a18024e1339c9241d6f1c3424457`, verified against the
manifest row) is a docxodus redline: the generator marked whole paragraphs — including runs
holding `mc:AlternateContent` text boxes — as deleted (`w:author="Open-Xml-PowerTools"`), and
inside each text box's own story (`w:txbxContent`) emitted bare `w:delText` / `w:delInstrText`
runs with no `w:del` inside that story. Word refuses to load the file; docx-validate's
`word-textbox-deleted-text` check reports it exactly (59 errors: 52 `word/document.xml`,
7 `word/footer1.xml`). Both violations re-verified against the unpacked XML and bundled XSDs;
both are the same generator bug.

Live Word vs manifest: the live run saw the recovery dialog *"Word found unreadable content in
docxodus__01a11a49_2317.docx"* (**No** → *"Word experienced an error trying to open the file"*,
no document window, no crash report, clean quit) — `repair-prompt`. The manifest row (corpus
copy `/Users/arthrod/temp/T/neurotic_docx_bench/maybe_future_corpus/really_broken/manifest.csv:155`;
the repo tree has no `manifest.csv`) records `word:corrupt` / *"Word could not load it in 2
passes: document loaded empty (Word could not read it)"* — the same refusal, coarser wording.

## Word behaviour

- **Live result (this ask):** `repair-prompt` — recovery dialog, then open error; nothing opened.
- **Manifest:** `word:corrupt`; "document loaded empty (Word could not read it)".
- **Prior assessment for this exact file:** none — `docs/really-broken/assessments.jsonl` and
  `docs/really-broken/inspection-log.jsonl` have **zero** rows for this path (grep run this
  session). The cause is therefore **not yet proven for this exact file**; the evidence is the
  live Word refusal plus the check, and the defect *kind* is proven on four siblings of the
  same docxodus family (`0011`, `0044`, `0070`, `0135`, records read this session in
  `inspection-log.jsonl`): in each, *"inside those `w:txbxContent` elements only, change
  `w:delText` to `w:t`"* made Word open clean, leaving all body `w:delText` in place. Sibling
  `0135` is this file's exact shape (`word/document.xml` + `word/footer1.xml` both affected).
- **docx-validate run (this session):** `pnpm exec tsx src/scripts/office/validate.ts
  really_broken/tool_redlines/repair_prompt/docxodus__01a11a49_2317.docx` (plus a `--json`
  rerun) → `valid: false`, profile `word-valid`, **59 errors `word-textbox-deleted-text`**
  (52 `word/document.xml`, 7 `word/footer1.xml`); side findings, all non-blocking: 2×
  `paraid-missing-element`, 1× `xml-bom-leading`, 1× `rel-ids-sequential` (info).

## Violation 1 — `word/document.xml`: 52 bare `w:delText` in the Choice text box

The deleted drawing run's text-box story carries deleted-text elements no `w:del` inside that
story covers. `w:del w:id="1018"` sits at `document.xml:8751`; the `mc:Choice Requires="wps"`
`w:txbxContent` spans **8797–9222** and holds 52 runs whose only run content is a bare
`w:delText` — zero `w:del` anywhere inside that `w:txbxContent` (verified by walk).

```xml
<w:del w:author="Open-Xml-PowerTools" w:id="1018" w:date="2000-01-01T00:00:00Z">  <!-- :8751 -->
  <w:r><w:rPr><w:b/><w:noProof/><w:sz w:val="20"/></w:rPr>
    <mc:AlternateContent>
      <mc:Choice Requires="wps">
        <w:drawing>… <wps:txbx>
          <w:txbxContent>                                                  <!-- :8797 -->
            <w:p w14:paraId="76169B6C" w14:textId="77777777">…</w:p>
            <w:r><w:rPr><w:b/><w:color w:val="000000"/></w:rPr>
              <w:delText>Please                                            <!-- :8812 -->
              </w:delText>
            </w:r>                                                          <!-- ×52, no w:del in this story -->
          </w:txbxContent></wps:txbx>…</w:drawing>
      </mc:Choice>
      <mc:Fallback>… v:textbox/w:txbxContent (:9241–9666), same 52 w:delText …</mc:Fallback>
```

**Schema rule.** This placement is *not* a particle error — no bundled XSD forbids it. The
cited declarations define the elements involved
(`src/scripts/office/schemas/ISO-IEC29500-4_2016/wml.xsd`):

- `wml.xsd:1745-1752` — `EG_RunInnerContent` admits `delText` in **any** run:
  ```xsd
  <xsd:group name="EG_RunInnerContent">
    <xsd:choice>
      <xsd:element name="br" type="CT_Br"/>
      <xsd:element name="t" type="CT_Text"/>
      <xsd:element name="delText" type="CT_Text"/>
      <xsd:element name="instrText" type="CT_Text"/>
      <xsd:element name="delInstrText" type="CT_Text"/>
  ```
- `wml.xsd:3378-3381` — `CT_TxbxContent` is only block-level content, no run-content
  restriction: `<xsd:complexType name="CT_TxbxContent"> <xsd:group ref="EG_BlockLevelElts"
  minOccurs="1" maxOccurs="unbounded"/> </xsd:complexType>`

Verified this session: `xmllint --noout --schema …/wml.xsd word/document.xml` reports only
Microsoft-extension-attribute noise (`w14:paraId/textId`, `wp14:anchorId/editId`) — **no
particle error for the `delText` placement**. The break is Word's *behavioural* rule — deleted
text inside a text-box story must be wrapped in a `w:del` within that story; a `w:del` around
the drawing run does not cover it — exactly what `word-textbox-deleted-text` encodes
(`src/scripts/office/checks/word-deleted-text.ts:48-73`: the deleted flag resets at each
`w:txbxContent`; `mc:Fallback` copies are excluded).

**Proposed fix (generator):** in docxodus's redline generation, when a run inside a
`w:txbxContent` story is part of a deletion, either (a) open a `w:del` *inside* that
`w:txbxContent` wrapping those runs (same author/id/date as the outer deletion — tolerated per
fixture 0948), or (b) keep `w:t`/`w:instrText` inside the text-box story and let the enclosing
run-level `w:del` carry the deletion — the form the proven sibling repairs converge on. Never
emit bare `w:delText`/`w:delInstrText` inside `w:txbxContent`. **Fix target: docxodus.**

## Violation 2 — `word/footer1.xml`: 5 bare `w:delText` + 2 bare `w:delInstrText`

Same shape in the deleted footer's text box. `w:del w:id="1333"` at `footer1.xml:12`; the
Choice `w:txbxContent` spans **54–181** with 5 `w:delText` (e.g. "Page" at :70) and 2
`w:delInstrText` ("PAGE" at :87, :149), no `w:del` inside it; Fallback copy at :200–327.

```xml
<w:del w:author="Open-Xml-PowerTools" w:id="1333" …>                        <!-- footer1.xml:12 -->
  <w:r>…<mc:AlternateContent><mc:Choice Requires="wps"><w:drawing>… <wps:txbx>
      <w:txbxContent>                                                        <!-- :54 -->
        …<w:r><w:rPr><w:rFonts w:ascii="Calibri"/><w:sz w:val="24"/></w:rPr>
            <w:delText xml:space="preserve">Page</w:delText></w:r>            <!-- :70 -->
        …<w:r>…<w:delInstrText xml:space="preserve">PAGE</w:delInstrText></w:r><!-- :87, :149 -->
      </w:txbxContent>…</wps:txbx></w:drawing></mc:Choice>
    <mc:Fallback>… same 5 + 2 in v:textbox/w:txbxContent (:200–327) …</mc:Fallback>
```

**Schema rule:** identical to Violation 1 (`EG_RunInnerContent` at `wml.xsd:1745-1752` admits
`delText`/`delInstrText` in any run; `CT_TxbxContent` at `wml.xsd:3378-3381` places no
restriction). xmllint on `word/footer1.xml` shows the same attribute noise plus one
`mc:AlternateContent: This element is not expected` (raw ISO wml.xsd does not model MCE
branches; the repo validator applies markup compatibility before schema work) — again no
particle error for the placement. **Proposed fix / target:** same as Violation 1 — **docxodus**.

## Where the fixes belong and why

- **docxodus (both violations).** The redline's XML is schema-legal but Word-refused; the
  generator chose the markup (`w:author="Open-Xml-PowerTools"` stamps the deletion pass). The
  fix is in the deletion logic around text-box stories, as above.
- **Not docx-validate.** The validator already detects and counts this defect correctly
  (`word-textbox-deleted-text`, 52 + 7 errors at error severity under word-valid) — no missing
  check, no false positive.
- **Not repair.** This is a generated redline, not a broken source document; the source fed to
  docxodus is not implicated. The minimal-change experiment below is the *verification*, not a
  prescription for a source document.

## Verification

1. **Re-verify (done this session):** the validator command above reproduces 59
   `word-textbox-deleted-text` errors (52 + 7); an independent Python re-implementation of the
   check's walk over the unpacked parts (`python3 -I /tmp/check2317.py <document.xml>
   <footer1.xml>`) finds exactly 52 and 7 (5 `delText` + 2 `delInstrText`) hits, with **zero**
   `w:del` inside all four `w:txbxContent` subtrees. Counts reconcile: `document.xml` has 975
   `delText`/`delInstrText` elements = 52 Choice box + 52 Fallback copy + 871 tolerated body
   sites; `footer1.xml` has 14, all inside the two boxes.
2. **Minimal-change experiment to prove the cause for this exact file** (unproven as of this
   issue — no assessment row for this path): inside the two `mc:Choice` `w:txbxContent`
   elements only (`document.xml:8797-9222`, `footer1.xml:54-181`), rename every `w:delText` →
   `w:t` and `w:delInstrText` → `w:instrText`, nothing else; expect a clean open. The
   `mc:Fallback` copies may stay (tolerated per fixture 2054 while the Choice is selected); if
   it still warns, extend the rename to the Fallback copies (sibling 0044 showed either copy
   alone suffices when it is the one Word reads). Two runs, logs read, per the Word-probe rule.
3. **docxodus regression guard:** generate a redline deleting a paragraph containing a text
   box; assert no bare `w:delText`/`w:delInstrText` inside any `w:txbxContent` — the existing
   `word-textbox-deleted-text` check on the output is the oracle.
