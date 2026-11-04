# [docxodus 2306] repair-prompt: w:delText inside a header text box Word refuses to load

## Summary

`tool_redlines/repair_prompt/docxodus__01a11a49_2306.docx` is a docxodus redline (every
tracked change stamped `w:author="Open-Xml-PowerTools"`, `w:date="2000-01-01T00:00:00Z"`)
of a Polish church-document header. Live Microsoft Word shows the recovery dialog
"Word found unreadable content in docxodus__01a11a49_2306.docx. Do you want to recover
the contents of this document?"; answering No is followed by "Word experienced an error
trying to open the file" and no document window. The manifest row
(`really_broken/manifest.csv`, restored corpus copy at
`~/temp/T/neurotic_docx_bench/maybe_future_corpus/really_broken/manifest.csv`) agrees:
`word:corrupt`, "Word could not load it in 2 passes: document loaded empty (Word could not
read it)", harvest 20261013_091613, sha256 `b40dd996f5eff1dc4a3f063e5870cf20e055aac381443b3c35fc65bdfd01915e`
(re-verified with `shasum -a 256` this session).

This repo's validator reports the dominant defect exactly once:

```
$ pnpm exec tsx src/scripts/office/validate.ts really_broken/tool_redlines/repair_prompt/docxodus__01a11a49_2306.docx
ERROR [word/header1.xml]: A text box contains deleted text. Microsoft Word offers to
  recover the file. Remove the w:delText or w:delInstrText from the text box.
exit 1
```

The `--json` run (`valid: false`) adds only non-blocking findings: warnings
`paraid-missing-element` (`word/document.xml`) and `xml-bom-leading`
(`word/_rels/header1.xml.rels`), infos `rel-ids-sequential` on both `.rels` parts.

## Word behaviour

- **Live result (ask-supplied):** recovery dialog, then open error on No; outcome
  `repair-prompt`; no crash report; Word quit cleanly. Agrees with the manifest's
  `word:corrupt`.
- **Prior assessment:** none for this exact path (`grep "tool_redlines/repair_prompt/docxodus__01a11a49_2306"
  docs/really-broken/assessments.jsonl docs/really-broken/inspection-log.jsonl` → exit 1,
  re-run this session). The cause is therefore **not yet proven for this file** — but it is
  proven for 20+ same-generator siblings carrying the identical defect
  (`assessments.jsonl`, status `proven`: docxodus 0044, 0070, 0110, 0135, 0162, 0164, 0214,
  0238, 0347, 0420, 0479, …). Row 0044 is the same shape as this file (drawingML
  `wps:txbx` in `mc:Choice` + VML twin in `mc:Fallback`): changing the Choice-branch
  `w:delText` to `w:t` (Fallback removed) opened clean; changing only the Fallback copy
  still warned; both changed opened clean; body `w:delText` left in place stayed tolerated.
  Row 0070: changing `w:delText` to `w:t` inside the 10 `w:txbxContent` elements only →
  clean open.

## Violation 1 — `word-textbox-deleted-text` (`word/header1.xml`, 1 site)

**Offending XML** (the stored part is one minified line; structure quoted from
`xmllint --format` of the raw bytes, `/tmp/header1_pretty_2306.xml` lines 227–343):

```xml
<w:del w:author="Open-Xml-PowerTools" w:id="1137" w:date="2000-01-01T00:00:00Z"> <!-- outside the text box -->
  <w:r>
    <mc:AlternateContent>
      <mc:Choice Requires="wps">
        <w:drawing>…<wps:wsp>…<wps:txbx>
            <w:txbxContent>
              <w:p w14:paraId="327F664C" w14:textId="77777777">
                <w:pPr><w:pStyle w:val="Logo-Mark"/></w:pPr>
                <w:r><w:delText xml:space="preserve"> </w:delText></w:r>
              </w:p>
            </w:txbxContent>
        </wps:txbx>…</wps:wsp>…</w:drawing>
      </mc:Choice>
      <mc:Fallback><w:pict>…<v:textbox><w:txbxContent>
        …<w:r><w:delText xml:space="preserve"> </w:delText></w:r>…
      </w:txbxContent></v:textbox>…</w:pict></mc:Fallback>
    </mc:AlternateContent>
  </w:r>
</w:del>
```

No `w:del` exists inside either `w:txbxContent`; the `w:del id=1137` wraps the run from
outside the text-box story. This is the header's only text box (`grep txbxContent` across
`word/*.xml` hits `header1.xml` alone: 2 elements, Choice + Fallback twin). Deleted text
elsewhere — `word/document.xml` 203 `w:delText`, `word/footer2.xml` 5, `word/footer1.xml`
1 `w:delInstrText`, none inside a text box — is tolerated per the proven rows.

**Schema rule — the element is LEGAL, so this is Word-behavioural, not XSD-grammatical.**
`src/scripts/office/schemas/ISO-IEC29500-4_2016/wml.xsd:1745`:

```xsd
<xsd:group name="EG_RunInnerContent">
  <xsd:choice>
    …
    <xsd:element name="delText" type="CT_Text"/>   <!-- line 1750: legal w:r child -->
```

and `wml.xsd:3378`:

```xsd
<xsd:complexType name="CT_TxbxContent">
  <xsd:group ref="EG_BlockLevelElts" minOccurs="1" maxOccurs="unbounded"/>
</xsd:complexType>
<xsd:element name="txbxContent" type="CT_TxbxContent"/>
```

`xmllint --noout --schema …/wml.xsd word/header1.xml` does report "fails to validate",
but every error is the known noise class (mc:Ignorable, w14:paraId/textId,
wp14:anchorId/editId, mc:AlternateContent and wp14:sizeRelH unexpected under raw wml.xsd);
`| grep -c delText` → **0** errors touch the `w:delText`. Correction to the analyst notes:
the raw part is minified — 0 of its 6 `wp:posOffset` values contain whitespace and xmllint
raises 0 posOffset errors; that claim was an artifact of a pretty-printed copy.

**Proposed fix (for the document):** inside the `mc:Choice` branch's
`wps:txbx > w:txbxContent`, rename `<w:delText xml:space="preserve"> </w:delText>` to
`<w:t xml:space="preserve"> </w:t>` and keep the outer `w:del id=1137` as-is — it deletes
the whole drawing run at the header level. The `mc:Fallback` twin needs no change for Word
(the check exempts it; fixture 2054: Word reads the sibling Choice), though changing both
is also proven clean (0044).

**Fix target: docxodus** (generator bug).

## Where the fixes belong and why

- **docxodus** — the generator marks a run that contains a text box as deleted by wrapping
  it in `w:del` from outside, then rewrites the text inside the text-box story to
  `w:delText`. A text box is a separate story: a `w:del` outside it does not cover content
  inside it, and Word refuses the file. When docxodus deletes a run containing a text box,
  it must emit `w:t` inside the `w:txbxContent` (or place the `w:del` inside that story).
  Same bug class as proven siblings 0044/0070/…, same generator stamp.
- **docx-validate — no change.** The existing check `word-textbox-deleted-text`
  (`src/scripts/office/checks/word-deleted-text.ts:100`) reports this defect exactly once,
  on the right part, with the evidenced semantics: entering `txbxContent` resets the
  deleted flag (lines 57–59, "a w:del outside the text box does not cover the text-box
  story") and `mc:Fallback` content is exempt (lines 62–64). Verified against this file:
  one error, one site (the Choice branch; the Fallback twin is the exempt copy).
- **Not "repair"** — this is a generated redline artifact, not a broken source document.

## Verification

1. **Minimal-change experiment (would prove the cause for this exact file; not yet run):**
   rename only the Choice-branch `w:delText` in `word/header1.xml` to `w:t`, repack, open
   in Word twice with logs (repo standard: verdict read from each run's log; both runs must
   agree). Expected: clean open, matching proven siblings 0044
   (`fallback_removed_and_choice_delText_changed_to_w_t: clean-open`) and 0070
   (`text_box_delText_changed_to_w_t: clean-open`).
2. **Validator regression (already passing this session):**
   `pnpm exec tsx src/scripts/office/validate.ts really_broken/tool_redlines/repair_prompt/docxodus__01a11a49_2306.docx`
   → exit 1 with exactly one ERROR (`word-textbox-deleted-text`, `word/header1.xml`); the
   repaired file from step 1 should exit 0 with no error.
3. **Generator-side regression:** re-run the docxodus redline for this source artifact and
   assert no `w:delText`/`w:delInstrText` inside any `w:txbxContent` lacking a `w:del`
   within that text box.
