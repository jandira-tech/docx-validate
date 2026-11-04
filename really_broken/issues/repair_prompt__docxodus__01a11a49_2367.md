# [docxodus 2367] repair-prompt: w:delText inside VML text box without in-story w:del

## Summary

`tool_redlines/repair_prompt/docxodus__01a11a49_2367.docx` is a docxodus redline
(sha256 `c9b3a119…e2a915`; the analyst recovered the bytes from the source corpus at
`/Users/arthrod/temp/T/neurotic_docx_bench/maybe_future_corpus/really_broken/…` after the
working tree lacked them — hash verified identical). Live Word offers recovery and then
fails to open the file. docx-validate's `word-valid` profile reports 13 error-severity
`word-textbox-deleted-text` findings — the only error-severity findings — agreeing with
Word. The defect is invisible to the raw XSD: it is a WordprocessingML **scoping** rule
(deleted text in a text-box story must be wrapped by a `w:del` inside that story), which
this repo carries as the named check `word-textbox-deleted-text`
(`src/scripts/office/checks/word-deleted-text.ts:99-127`).

## Word behaviour

- **Live result (this run):** `repair-prompt` — dialog *"Word found unreadable content in
  docxodus__01a11a49_2367.docx. Do you want to recover the contents of this document?"*;
  clicking No was followed by *"Word experienced an error trying to open the file"* and no
  document window. No crash report; Word quit cleanly.
- **Manifest:** the repo's `really_broken/manifest.csv` does not exist. The source corpus
  manifest row (`neurotic_docx_bench/maybe_future_corpus/really_broken/manifest.csv`) says
  category `repair_prompt`, tool `docxodus`, `d1_failed_in word:corrupt`,
  `local_word_error` *"Word could not load it in 2 passes: document loaded empty (Word
  could not read it)"* — consistent with the live outcome.
- **Prior assessment:** none. `docs/really-broken/assessments.jsonl` and
  `inspection-log.jsonl` exist but contain no row for this file (grep for
  `docxodus__01a11a49_2367` returns 0 in both). **The cause is therefore not yet proven
  for this exact file**; the evidence is the schema analysis, the check's proven fixture
  refusals, and the live repair prompt. See Verification for the minimal-change experiment.

## Violation 1 — `word-textbox-deleted-text` (13 sites, `word/document.xml`)

**Offending XML** — one VML text box (`v:rect id="Rectangle 2"` at line 4352 →
`v:textbox` at 4354 → `w:txbxContent` spanning lines 4355–5017). The only covering
deletion wraps the `w:pict` **from outside** the text box:

```xml
<w:del w:author="Open-Xml-PowerTools" w:id="400" w:date="2000-01-01T00:00:00Z">  <!-- line 4342: OUTSIDE the text box -->
  <w:r>
    <w:pict>
      <v:rect id="Rectangle 2" o:spid="_x0000_s1513" …>                          <!-- line 4352 -->
        <v:textbox style="mso-next-textbox:#Rectangle 2" …>                      <!-- line 4354 -->
          <w:txbxContent>                                                       <!-- line 4355 -->
            <w:p>…
              <w:r>
                <w:rPr>…</w:rPr>
                <w:delText xml:space="preserve">Professionalization of Exercise
                </w:delText>                                                    <!-- line 4397, site 1 -->
              </w:r>
```

13 `w:delText` runs sit in this one `w:txbxContent` at lines 4397, 4409, 4421, 4464,
4506, 4517, 4528, 4555, 4567, 4608, 4638, 4916, 5004. Inside lines 4355–5017 there is
**no** `w:del` or `w:ins` element of any form (verified: `grep -o "<w:del[ />]"` over the
range returns nothing) and no `mc:Fallback` anywhere in the document (`grep -c` = 0).

**Schema rule (re-read this session).** The raw XSD permits this shape, which is exactly
why the validator needs the named check:

```xml
<xsd:complexType name="CT_TxbxContent">                <!-- wml.xsd:3378-3380 -->
  <xsd:group ref="EG_BlockLevelElts" minOccurs="1" maxOccurs="unbounded"/>
</xsd:complexType>
<xsd:element name="txbxContent" type="CT_TxbxContent"/> <!-- wml.xsd:3381 -->

<xsd:group name="EG_RunInnerContent">                   <!-- wml.xsd:1745 -->
  <xsd:choice>
    …
    <xsd:element name="delText" type="CT_Text"/>        <!-- wml.xsd:1750 -->
```

`delText` is ordinary run inner content (`CT_R`, wml.xsd:1782-1786, references the group
at 1785), and `CT_TxbxContent` freely allows block content, so an XSD walk reports
nothing. Confirmed: `xmllint --noout --schema
src/scripts/office/schemas/ISO-IEC29500-4_2016/wml.xsd` over the unpacked
`word/document.xml` yields 210 errors, **all** `w14:paraId`/`w14:textId` attribute noise
(filtering those leaves only the summary line) — zero content-model findings.

The rule actually broken is the WordprocessingML scoping requirement, encoded in
`src/scripts/office/checks/word-deleted-text.ts:7-14`: `w:delText`/`w:delInstrText`
inside `w:txbxContent` with no `w:del` inside that text box makes Word offer recovery —
proven by fixture 0011 (VML text box) and 0044 (DrawingML `wps:txbx`); tolerated are
fixture 2054 (`mc:Fallback` text box), fixture 0948 (`w:del` inside the text box), and
plain body `w:delText`. Line 14 states it directly: *"A `w:del` outside the text box does
not cover the text-box story."* That is precisely this file's shape: `w:del id=400`
closes at line 5025, after `</w:pict>`.

**Tolerated sites, correctly not flagged (re-verified):** the document's other 546
`w:delText` elements sit inside body `w:del` wrappers; `footer1.xml`'s single `w:delText`
(line ~96) sits inside `w:del w:id="1062"` in the footer body — not a text box. No
`w:moveFrom` exists anywhere (document, footer1, footer2: grep 0), so
`word-move-from-deleted-text` does not apply.

**Proposed fix — in docxodus:** when the redline marks text-box content as deleted, wrap
the deleted runs in a `w:del` **inside** the `w:txbxContent` (each `<w:del w:id=…
w:author=… w:date=…>` around its `<w:r>`), matching the tolerated fixture-0948 shape
(repo working fixture `tests/fixtures/working/package.word-textbox-inner-deletion.docx`);
or, when the deletion is already expressed by the outer `w:del` around the `w:pict` run,
emit ordinary `w:t` inside the story instead of `w:delText`. Either shape passes the
check; bare in-story `w:delText` does not.

**Fix target:** `docxodus` (generator bug).

## Where the fixes belong and why

- **docxodus** — the redline's XML is genuinely Word-invalid: the generator (its `w:del`
  carries `w:author="Open-Xml-PowerTools"`) emitted bare `w:delText` runs inside the
  text-box story while expressing the deletion only through the outer `w:del` around the
  `w:pict`. Word refuses this; the generator must change as prescribed above.
- **docx-validate — no change.** The named check already reports all 13 sites at error
  severity under `word-valid`, in exactly the part Word reads, matching the live repair
  prompt; this is a correct detection, not a gap or false positive.
- **Not `repair`** — this is a tool-produced redline, not a broken source document.

## Verification

Checked this session (commands run, output observed):

1. `pnpm exec tsx src/scripts/office/validate.ts
   really_broken/tool_redlines/repair_prompt/docxodus__01a11a49_2367.docx --profile
   word-valid --json` → `valid: false`; issue census: 13 × error `word-textbox-deleted-text`
   (all `word/document.xml`), 5 × warning `paraid-missing-element`, 2 × warning
   `xml-bom-leading`, 1 × info `rel-ids-sequential` — reproducing the analyst's counts;
   the error findings are the only blocking ones.
2. `xmllint --noout --schema src/scripts/office/schemas/ISO-IEC29500-4_2016/wml.xsd
   really_broken/unpacked/tool_redlines__repair_prompt__docxodus__01a11a49_2367/word/document.xml`
   → 210 errors, all `w14:paraId`/`w14:textId` noise; no `delText`/`txbxContent`
   content-model error — the defect is not a raw XSD break.
3. Existing suite: `tests/word-textbox-deleted-text.test.ts` already pins VML and
   DrawingML refusals and the tolerated shapes (removed / inner-deletion / fallback).

**Minimal-change experiment to prove the cause for this file (not yet run — no proven
assessment exists):** in a copy, rename the 13 bare `w:delText` elements inside
`word/document.xml` lines 4355–5017 to `w:t` (or wrap each in a `w:del` inside the story),
leaving every other byte identical, then run two Word probes per the repo's
check-twice-with-logs rule. A clean open in both runs proves the in-textbox `w:delText`
is the blocking defect; only then record the observation in the manifest.
