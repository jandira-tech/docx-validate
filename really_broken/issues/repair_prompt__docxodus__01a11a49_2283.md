# [docxodus 2283] repair-prompt: uncovered `[trash]/*.dat` parts and text-box `w:delText`

File: `really_broken/tool_redlines/repair_prompt/docxodus__01a11a49_2283.docx`
(sha256 `bab3360a36f4f829a029486660d6ebfcdc24fefbc1636319831de9a7802ba178`, staged and verified).
Unpacked: `really_broken/unpacked/tool_redlines__repair_prompt__docxodus__01a11a49_2283/`.

## Summary

A docxodus redline output ("repair_prompt" lane; tracked-change author strings say
`Open-Xml-PowerTools`, the generator's lineage). Word refuses it with the recovery
dialog; the package carries two independent Word-blocking defects, both produced by the
generator: two stored parts `[trash]/0000.dat` and `[trash]/0001.dat` with no content type
(`ct-uncovered-part`, 2 sites), and a bare `w:delText` inside a `wps:txbx` text-box story
of `word/header2.xml` (`word-textbox-deleted-text`, 1 site).

`pnpm exec tsx src/scripts/office/validate.ts <file>` (run this session) exits 1 with
exactly those three ERROR lines; `--json` adds only non-blocking warnings/info
(`paraid-missing-element`, `rels-unreferenced` for both `[trash]` parts, `xml-syntax`,
`xml-bom-leading`, `rel-ids-sequential`). Both codes are Word-blocking
(`src/scripts/office/validators/word-blocking.ts:241,293`).

## Word behaviour

- **Live (this session's harness):** `repair-prompt` — "Word found unreadable content in
  docxodus__01a11a49_2283.docx. Do you want to recover…"; No → "Word experienced an error
  trying to open the file", no document window, no crash report.
- **Manifest** (`really_broken/manifest.csv` at the bench corpus,
  `/Users/arthrod/temp/T/neurotic_docx_bench/maybe_future_corpus/really_broken/`):
  `word:corrupt`, "Word could not load it in 2 passes: document loaded empty (Word could
  not read it) | harvest 20261013_091613". Agrees with the live result.
- **Prior assessment for this file: none.** `grep -c "docxodus__01a11a49_2283"` over
  `docs/really-broken/assessments.jsonl` and `inspection-log.jsonl` returns 0 for each
  (run this session). So the cause is **not yet proven for this file** — but both defect
  classes are proven on corpus siblings and pinned fixtures, below.

## Violation 1 — `ct-uncovered-part` (2 sites): `[trash]/0000.dat`, `[trash]/0001.dat`

**Offending state.** `zipinfo -1` lists both parts as stored entries. `[Content_Types].xml`
(read in full this session) declares only:

```xml
<Default Extension="png" ContentType="image/png"/>
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Default Extension="jpeg" ContentType="image/jpeg"/>
<!-- no Extension="dat" Default; no Override names /[trash]/0000.dat or /[trash]/0001.dat -->
```

**Rule.** OPC (ECMA-376 Part 2 §10.1.2.4): every part shall have a content type, resolved
Override-first then Default. `src/scripts/office/schemas/ecma/fourth-edition/opc-contentTypes.xsd:11-21`
(read this session):

```xml
<xs:complexType name="CT_Types">
  <xs:choice minOccurs="0" maxOccurs="unbounded">
    <xs:element ref="Default"/>
    <xs:element ref="Override"/>
  </xs:choice>
</xs:complexType>
<xs:complexType name="CT_Default">
  <xs:attribute name="Extension" type="ST_Extension" use="required"/>
  <xs:attribute name="ContentType" type="ST_ContentType" use="required"/>
</xs:complexType>
```

The `.dat` extension matches neither arm for either part.

**Sibling Word cause (proven, not this file).** `docs/really-broken/assessments.jsonl`
(row for `tool_redlines/crashes_word/superdoc-redlines__01a11a49_1812.docx`, landed with
commit `deb626d9` "Treat a package part with no content type as a Word repair"): the
sibling's only uncovered part was `[trash]/0000.dat`; Word outcomes were
`original: unreadable-content-warning`, `that_part_deleted: clean-open`,
`extra_note_txt_with_no_content_type: unreadable-content-warning`,
`extra_note_xml_covered_by_the_xml_default: clean-open`. One undeclared part is enough;
an unreferenced part is not the cause.

**Proposed fix — target: docxodus.** Do not write the `[trash]/*.dat` discard parts
(235 and 1631 bytes, binary starting `FF FF FF FF` — the PowerTools DocumentBuilder
trash-can convention) into the saved redline: nothing references them (both drew
`rels-unreferenced`), so delete them from the package. If they must be retained, add
`<Default Extension="dat" ContentType="application/octet-stream"/>`.

## Violation 2 — `word-textbox-deleted-text` (1 site): `word/header2.xml`

**Offending XML** (`word/header2.xml`, read this session; line numbers from the unpacked part):

```xml
<w:del w:author="Open-Xml-PowerTools" w:id="998" ...>   <!-- :10 — header-story del, outside the text box -->
  <w:r><mc:AlternateContent><mc:Choice Requires="wps">
    <w:drawing>...<wps:txbx>
      <w:txbxContent>                                      <!-- :58 — the text box's own story begins; del state resets -->
        <w:p w14:paraId="1EC0BB46" ...>
          <w:r><w:rPr>...</w:rPr>
            <w:delText>Hindi
            </w:delText>                                   <!-- :73 — no w:del inside this story -->
          </w:r>
        </w:p>
      </w:txbxContent>
    </wps:txbx>...</mc:Choice>
  ...</w:r>
</w:del>
```

The `mc:Fallback` `v:textbox` twin (`:96-128`) carries the same `w:delText` (`:119`); the
check does not count it as a site (`src/scripts/office/checks/word-deleted-text.ts:62-64`
excludes `mc:Fallback`), but the proven sibling shows Word refuses either copy alone, so
the repair must fix both.

**Rule.** The text box is its own block-level story: `wml.xsd:3378-3380` —
`<xsd:complexType name="CT_TxbxContent"><xsd:group ref="EG_BlockLevelElts" …/></xsd:complexType>`.
`delText` is plain run inner content (`wml.xsd:1745-1750`, `EG_RunInnerContent`:
`<xsd:element name="delText" type="CT_Text"/>`), and deleted run content belongs under
`w:del` (`CT_RunTrackChange`, `wml.xsd:997-1005`; `<xsd:element name="del" type="CT_RunTrackChange"/>`,
`wml.xsd:3406`). I ran raw `xmllint --noout --schema wml.xsd word/header2.xml` myself: the
delText draws **no** XSD finding (0 mentions) — the raw schema admits it in any run, so the
refusal is Word's semantic rule that a text-box story's deleted text be wrapped in a `w:del`
inside that story, which this repo encodes as the named check
(`src/scripts/office/checks/word-deleted-text.ts:6-14`: a `w:del` outside the text box does
not cover the text-box story; entering `txbxContent` resets the del state, `:57-59`).

**Sibling Word cause (proven, not this file).** `assessments.jsonl` row for
`tool_redlines/repair_prompt/docxodus__01a11a49_0044.docx`: same Choice-`wps:txbx` +
Fallback-`v:textbox` twin shape ("Either copy is enough for Word to offer recovery");
outcomes included `fallback_removed_and_choice_delText_changed_to_w_t: clean-open` and
`both_text_boxes_delText_changed_to_w_t: clean-open`, while `only_the_vml_delText_changed_to_w_t`
still warned. The bundled fixtures (`broken/package.word-textbox-deleted-text.docx`, `-shape`,
`-choice-deleted-text`) carry hash-bound `unreadable-content-warning` Word observations in
`tests/fixtures-all.manifest.json`.

**Proposed fix — target: docxodus.** When redlining the deletion of a run that contains a
text box, keep the text inside `w:txbxContent` as plain `w:t` in **both** the Choice and the
Fallback copies (the outer `w:del` already records the deletion of the run holding the
drawing); a `w:del` inside the text box is the tolerated alternative (fixture 0948).

## Where the fixes belong and why

Both violations are generator output defects, so both belong to **docxodus**; docx-validate
needs no change (both codes exist, reproduce exactly, and are Word-blocking; no false
positives — the Fallback twin is correctly not counted, and the body/footer `delText`
elements correctly draw nothing). Not a broken source document, so no repair prescription
beyond the document edits above.

## Verification

Minimal-change experiments on a copy of this file, each opened in Word twice with logs
(the repo's two-run rule), would prove the causes this file lacks an assessment for:
(1) delete `[trash]/0000.dat` and `[trash]/0001.dat` from the ZIP — sibling 1812 predicts
clean-open; (2) rename the two bare `w:delText` to `w:t` inside `word/header2.xml`'s two
`w:txbxContent` stories — sibling 0044 predicts clean-open; run them separately, then
together, to attribute the prompt. Portable check: `pnpm exec tsx
src/scripts/office/validate.ts <repaired copy>` exits 0 with no ERROR lines (run this
session on the original: exit 1, the three ERROR lines above).
