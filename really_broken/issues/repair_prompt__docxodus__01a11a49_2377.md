# [docxodus 2377] repair-prompt: unwrapped delText in footer text-box story

## Summary

`really_broken/tool_redlines/repair_prompt/docxodus__01a11a49_2377.docx` (sha256
`a77421dc77fc5045b7581f1d0a42f91e2650c61bd39ecba8573eabf7e395bb3c`, byte-identical to its
master in `neurotic_docx_bench/maybe_future_corpus`) is a docxodus tracked-changes redline: the
generator (author `Open-Xml-PowerTools` on every marker) deleted a footer's page-number line,
wrapping the run that carries a text box in `w:del` **outside** the text box while emitting
`w:delText`/`w:delInstrText` runs **inside** the text-box story with no covering `w:del` there.

- Live Word (this run): **repair-prompt** — "Word found unreadable content in
  docxodus__01a11a49_2377.docx. Do you want to recover…?"; answering No gave "Word experienced
  an error trying to open the file" and no document window. `agrees: true`.
- Manifest: `really_broken/manifest.csv` does not exist in this checkout (verified `ls`); the
  ask's `manifestSays` field is empty. No manifest row could be read.
- Repo validator: `valid: false`, exactly **6 error issues, all `word-textbox-deleted-text` in
  `word/footer1.xml`**; everything else warning/info (`paraid-missing-element`,
  `style-default-missing`, `rels-empty-part`, `word-cell-margin-order`, `rel-ids-sequential`).

## Word behaviour

- Live result (ask-provided): recovery dialog, then open error on No; no crash report; Word quit
  cleanly. Outcome `repair-prompt`.
- Prior assessment: **none for this file** — `grep -c "docxodus__01a11a49_2377"` returns `0` in
  both `docs/really-broken/assessments.jsonl` and `docs/really-broken/inspection-log.jsonl`.
- Family proven causes (same defect, sibling docxodus files):
  - **0044** (identical `mc:Choice` `wps:txbx` + `mc:Fallback` `v:textbox` shape, 25+25
    elements): either copy alone still warns; renaming `w:delText` → `w:t` inside **both**
    `w:txbxContent` elements → **clean-open**.
  - **0011** (VML `v:textbox` in `word/footer1.xml`, the same part as here): one `w:delText` or
    one `w:delInstrText` alone makes Word offer recovery; removing them opens clean; a text box
    with only field characters opens clean; the body's 10018 `w:delText` are not the cause.
  - Fixture 0948 (cited in the check's header, `src/scripts/office/checks/word-deleted-text.ts:13`):
    a `w:del` **inside** the text box is tolerated.

## Each violation

### `word-textbox-deleted-text` — `word/footer1.xml`, 6 sites (effective view)

The footer's only paragraph holds `w:del` id **1039** (`word/footer1.xml:17`) wrapping a run
whose `mc:AlternateContent` (`:22`) offers the same text box twice: `mc:Choice Requires="wpg"`
→ `w:drawing` → `wps:txbx` → `w:txbxContent` (`:58-179`) and `mc:Fallback` → `w:pict` →
`v:textbox` → `w:txbxContent` (`:193-314`). Inside each story (verified by regex over both
blocks: `del=0 delText=5 delInstrText=1` each):

```xml
<w:del w:author="Open-Xml-PowerTools" w:id="1039" ...>   <!-- footer body story, outside the box -->
  <w:r>...
    <mc:AlternateContent><mc:Choice Requires="wpg"><w:drawing>...<wps:txbx><w:txbxContent>
      <w:p>...
        <w:r>...<w:delText xml:space="preserve">—</w:delText></w:r>          <!-- :82 -->
        <w:r>...<w:delText xml:space="preserve"> </w:delText></w:r>          <!-- :91 -->
        <w:r>...<w:fldChar w:fldCharType="begin"/></w:r>
        <w:r>...<w:delInstrText xml:space="preserve">PAGE \* MERGEFORMAT</w:delInstrText></w:r>  <!-- :108 -->
        <w:r>...<w:fldChar w:fldCharType="separate"/></w:r>
        <w:r>...<w:delText xml:space="preserve">1</w:delText></w:r>          <!-- :125 -->
        <w:r>...<w:fldChar w:fldCharType="end"/></w:r>
        <w:r>...<w:delText xml:space="preserve">—</w:delText></w:r>          <!-- :151 -->
      </w:p></w:txbxContent></wps:txbx></w:drawing></mc:Choice>
    <mc:Fallback><w:pict><v:textbox><w:txbxContent> … same 5 w:delText + 1 w:delInstrText …
      </w:txbxContent></v:textbox></w:pict></mc:Fallback></mc:AlternateContent></w:r></w:del>
```

No `w:del` exists inside either `w:txbxContent` (verified). Note an analyst-report correction:
the wrapper is id **1039** (`:17`); id 1041 is the paragraph-mark deletion in `w:pPr/w:rPr`
(`:14`) and id 1040 wraps an empty run (`:322`) — the analyst's snippet attributed the wrapper
to 1041.

**Schema rule — this is XSD-legal; the breach is Word's behaviour.**
`src/scripts/office/schemas/ISO-IEC29500-4_2016/wml.xsd:1745-1752` (`EG_RunInnerContent`):

```xsd
<xsd:group name="EG_RunInnerContent">
  <xsd:choice>
    ...
    <xsd:element name="t" type="CT_Text"/>
    <xsd:element name="delText" type="CT_Text"/>
    <xsd:element name="instrText" type="CT_Text"/>
    <xsd:element name="delInstrText" type="CT_Text"/>
```

and `wml.xsd:3378-3380` (`CT_TxbxContent`):

```xsd
<xsd:complexType name="CT_TxbxContent">
  <xsd:group ref="EG_BlockLevelElts" minOccurs="1" maxOccurs="unbounded"/>
</xsd:complexType>
```

`delText`/`delInstrText` are ordinary run inner content — the XSD never ties them to a covering
`w:del`. Confirmed by run: `xmllint --noout --schema …/wml.xsd word/footer1.xml` fails **only**
on `mc:Ignorable` (line 2) and `mc:AlternateContent` (line 22) — markup-compatibility noise the
repo preprocesses before XSD — with no delText finding. The defect is that a text box is a
separate story: `w:del` id 1039 covers the drawing run in the footer story but cannot reach
inside `w:txbxContent`, so the deleted-text elements sit uncovered in their own story, which
Word refuses to open.

**Proposed fix (docxodus generator):** when a deletion covers a text box, emit tracked
deletions *within the text-box story* — wrap the runs in `w:del` inside each `w:txbxContent`
(tolerated, fixture 0948) — or, since the outer `w:del` on the drawing run already tracks the
whole box as deleted, emit plain `w:t`/`w:instrText` inside both `w:txbxContent` copies
(proven clean-open by 0044). The generator must apply the change to **both** the `mc:Choice`
and `mc:Fallback` copies.

**Fix target:** docxodus.

## Where the fixes belong and why

- **docxodus** — the redline's XML is genuinely defective for Word: the generator emitted
  `w:delText`/`w:delInstrText` inside a text-box story with no `w:del` in that story, in both
  the Choice and Fallback copies. Word's recovery offer (live, this run) proves it.
- **docx-validate — no change.** The check exists and fires correctly:
  `src/scripts/office/checks/word-deleted-text.ts:99-127` (`word-textbox-deleted-text`), whose
  walk (`:48-73`) resets the deleted flag on entering `w:txbxContent` (`:57-59`) — encoding
  exactly "a `w:del` outside the text box does not cover the text-box story" (`:14`) — and
  skips `mc:Fallback` (`:62-64`), hence 6 errors for the Choice's 5 `w:delText` + 1
  `w:delInstrText`, matching the live Word outcome. No false positive, no gap.
- Not the `repair` lane: this is generator output, not a broken source document. For any
  already-produced copy, the document-level repair is 0044's proven change: rename the 5
  `w:delText` → `w:t` and the 1 `w:delInstrText` → `w:instrText` inside **both**
  `w:txbxContent` elements of `word/footer1.xml`.

## Verification

- Ran (this session): `pnpm exec tsx src/scripts/office/validate.ts
  really_broken/tool_redlines/repair_prompt/docxodus__01a11a49_2377.docx --json` →
  `valid: false`, 6 × error `word-textbox-deleted-text` (`word/footer1.xml`); all other issues
  warning/info. Matches the analyst's run.
- Ran: `xmllint --noout --schema src/scripts/office/schemas/ISO-IEC29500-4_2016/wml.xsd
  really_broken/unpacked/repair_prompt__docxodus__01a11a49_2377/word/footer1.xml` → only the
  two mc findings above; no delText finding (schema legality).
- Unpacked with `pnpm exec tsx src/scripts/office/unpack.ts … --merge-runs false
  --simplify-redlines false` (21 XML files); `word/document.xml` has 0 `w:txbxContent`, so its
  517 body `w:delText` are tolerated deletions outside this defect.
- Cause for this exact file is **not yet proven** (no prior assessment). Minimal-change
  experiment that would prove it, per the repo's two-run Word protocol: in a copy, rename the 5
  `w:delText` → `w:t` and 1 `w:delInstrText` → `w:instrText` inside both `w:txbxContent`
  elements of `word/footer1.xml`, repack, open in Word twice — expect clean-open both runs
  (0044's identical experiment on the identical shape returned clean-open).
- Existing coverage: `tests/fixtures/broken/package.word-textbox-choice-deleted-text.docx` and
  siblings pin this check in `tests/checks-word-deleted-text.test.ts`.
