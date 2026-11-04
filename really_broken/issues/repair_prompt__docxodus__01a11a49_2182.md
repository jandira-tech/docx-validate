# [docxodus 2182] repair-prompt: bare `w:delText` in a text box makes Word offer recovery

## Summary

`really_broken/tool_redlines/repair_prompt/docxodus__01a11a49_2182.docx` (sha256
`fcd72a949bbc269210a23273fe684455db63b55fb7970f1bbc517119fc48132b`) is a tracked-changes
redline produced by **docxodus** (deletions authored `Open-Xml-PowerTools`, dated
`2000-01-01T00:00:00Z`): `word/document.xml` carries 121 `w:del`, 254 `w:ins`, one
`mc:AlternateContent` text box (line 4167), 110 `w:delText` and 2 `w:txbxContent`
stories (Choice + VML Fallback twin).

**Live Word outcome:** repair-prompt — Word showed "Word found unreadable content in
docxodus__01a11a49_2182.docx. Do you want to recover the contents of this document?";
answering No was followed by "Word experienced an error trying to open the file" and no
document opened. **Manifest:** `word:corrupt` ("document loaded empty (Word could not
read it) | harvest 20261013_044337") — both readings reject the same bytes; the live run
(`agrees: true`) is the sharper one. Note: `really_broken/manifest.csv` does not exist
in this repo; the row was found at
`neurotic_docx_bench/maybe_future_corpus/really_broken/manifest.csv:137` (external
bench), whose sha256 matches the repo copy byte-for-byte.

## Word behaviour

- **Live run (this session, per the ask):** repair-prompt, as quoted above.
- **Prior assessment — PROVEN** (`docs/really-broken/assessments.jsonl:140`,
  `docs/really-broken/inspection-log.jsonl:140`, source `word-ab`):
  - original bytes → Word `unreadable-content-warning`; validator invalid,
    error code `word-textbox-deleted-text`.
  - minimal-change experiment `2182-choice-plain.docx` — renamed **only** the
    `mc:Choice` text box's `w:delText` to `w:t` (Fallback `w:delText` and all 109 body
    `w:delText` untouched) → Word **clean-open** and validator **valid**.

## Violation 1 — `word-textbox-deleted-text` (word/document.xml) — CONFIRMED

### Offending XML

`word/document.xml:4161-4223` — a body-level deletion wraps the run that carries the
whole text box, but the box's own story holds a bare `w:delText` with **no `w:del`
inside that story** (verified by parse: each of the 2 `w:txbxContent` stories contains
exactly 1 `w:delText` and 0 `w:del`):

```xml
<w:del w:author="Open-Xml-PowerTools" w:id="382" w:date="2000-01-01T00:00:00Z">  <!-- body story -->
  <w:r><w:rPr>…</w:rPr>
    <mc:AlternateContent>                                          <!-- :4167 -->
      <mc:Choice Requires="wps"><w:drawing><wp:anchor …>…<wps:txbx><w:txbxContent>
          <w:p w14:paraId="5ED9C7C5" w14:textId="684ED14B">…
            <w:r><w:rPr>…</w:rPr><w:delText>1B
              </w:delText>                                          <!-- :4218 — bare, no w:del in this story -->
          </w:p></w:txbxContent></wps:txbx>…
      </mc:Choice>
      <mc:Fallback><w:pict><v:shape …><v:textbox>
        <w:txbxContent>… identical bare w:delText …</w:txbxContent> <!-- twin; not Word-read -->
      </v:textbox></v:shape></w:pict></mc:Fallback>
    </mc:AlternateContent>
  </w:r>
</w:del>
```

### Schema rule (why the XSD alone cannot flag it)

`src/scripts/office/schemas/ISO-IEC29500-4_2016/wml.xsd:1745-1750` — `delText` is
declared as generic run inner content, legal inside any `w:r`:

```xsd
<xsd:group name="EG_RunInnerContent">
  <xsd:choice>
    <xsd:element name="br" type="CT_Br"/>
    <xsd:element name="t" type="CT_Text"/>
    <xsd:element name="contentPart" type="CT_Rel"/>
    <xsd:element name="delText" type="CT_Text"/>
```

`wml.xsd:3378-3381` — a text box is its own block-level story; `wml.xsd:3399-3406`
(`EG_RunLevelElts`) declares the deleted-run wrapper a story provides:

```xsd
<xsd:complexType name="CT_TxbxContent">
  <xsd:group ref="EG_BlockLevelElts" minOccurs="1" maxOccurs="unbounded"/>
</xsd:complexType>
<xsd:element name="txbxContent" type="CT_TxbxContent"/>
…
<xsd:element name="ins" type="CT_RunTrackChange" minOccurs="0"/>
<xsd:element name="del" type="CT_RunTrackChange" minOccurs="0"/>
```

Read together: the XSD permits `delText` in any run — including one inside a text-box
story — so the schema **cannot** express "delText in a text box must sit under a `w:del`
of that text box's story". That rule is Word's story scoping, encoded by this repo's
check `word-textbox-deleted-text` (`src/scripts/office/checks/word-deleted-text.ts`):
its walk resets the deleted state at each `w:txbxContent` boundary, tolerates `w:del`
inside the box and `mc:Fallback` boxes, and reports a bare `w:delText`/`w:delInstrText`
— exactly this site. The body `w:del` (id 382) wraps the run *outside* the box and does
not cover the box's story; Word offers recovery.

### Proposed fix — target: **docxodus** (generator)

When docxodus marks a run containing a text box as deleted, it converts the text inside
the box's story to `w:delText` while recording the deletion only in the body story. A
`w:txbxContent` is a separate story: its deletions must be recorded inside it. For each
deleted text box, the generator should either:

- wrap the box's runs in their own `w:del` (own `w:id`, same author/date) inside the
  `w:txbxContent` — the check tolerates this (fixture 0948) and it preserves the
  tracked-deletion semantics; or
- emit `w:t` in the box's story if the box's content should not be tracked as deleted —
  the variant the proven experiment validated (clean-open).

Apply the same treatment to the `mc:Fallback` twin for branch consistency (the Fallback
is not Word-blocking — Word reads the Choice — but the branches should not diverge).

## Violation 2 — `mc:AlternateContent` in `w:r` — recorded for completeness, NOT a defect

Raw `xmllint --noout --schema …/wml.xsd word/document.xml` reports, besides 160
`w14:paraId`/`w14:textId` attribute errors and 4 `w14:checkbox` element errors
(Microsoft wordml-2010 extension noise), exactly one further error:
`document.xml:4167: Element '{…markup-compatibility/2006}AlternateContent': This element
is not expected` — `EG_RunInnerContent` (wml.xsd:1745-1779) has no MCE alternative; the
ISO schema set predates MCE wiring at this run slot. Standard MCE markup Word supports;
this repo's validator accepts it after markup-compatibility preprocessing and the
word-valid run reports no issue for it. No fix, no check needed.

## Where the fixes belong and why

- **docxodus** — violation 1 is a genuine generator defect: the redline XML it emits
  makes Word refuse the file. The fix above belongs in docxodus's redline generation.
- **docx-validate — no change.** The check `word-textbox-deleted-text` already reports
  this defect correctly (error, word-valid/strict, story-scoped), and the advisory
  findings alongside it (4× `xml-bom-leading` warnings — the check's own message notes a
  UTF-8 BOM is "Permitted by the XML spec (and emitted by Microsoft Office)" — and 1×
  `rel-ids-sequential` info) are correctly non-blocking. No gap, no false positive.
- **repair — not applicable** (the source document is not the broken party here).

## Verification

Run during this session (all reproduced the analyst report):

1. `pnpm exec tsx src/scripts/office/validate.ts really_broken/tool_redlines/repair_prompt/docxodus__01a11a49_2182.docx`
   → exit 1, one ERROR: `word-textbox-deleted-text` on `word/document.xml`.
2. Same command with `--json` → `valid: false`, `metadata.profile: "word-valid"`;
   issues: 4× warning `xml-bom-leading`, 1× info `rel-ids-sequential`, 1× error
   `word-textbox-deleted-text`.
3. Same command with `--profile lenient` → "All validations PASSED!", exit 0 (the Word
   rules run only under word-valid/strict — expected, not a gap).
4. `xmllint --noout --schema src/scripts/office/schemas/ISO-IEC29500-4_2016/wml.xsd
   <unpacked>/word/document.xml` → 160 paraId/textId + 4 checkbox + 1 AlternateContent
   errors; **nothing about `delText`** — the defect is story-scoped Word behaviour, not
   an XSD break.
5. Cause is **proven**, not merely suspected: the recorded minimal-change experiment
   (`2182-choice-plain.docx`, Choice `w:delText`→`w:t` only) opened clean in Word and
   passed the validator (`assessments.jsonl:140`). Re-proving a docxodus fix means
   regenerating the redline, getting a clean open in two Word runs (the repo's probe
   standard) and `validate.ts` reporting no error.
