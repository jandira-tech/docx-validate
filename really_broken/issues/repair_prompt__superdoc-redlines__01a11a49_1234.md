# [superdoc-redlines 1234] repair-prompt: hyperlink r:ids name no relationship

## Summary

`tool_redlines/repair_prompt/superdoc-redlines__01a11a49_1234.docx` (1,734,121 bytes,
sha256 `53186d5e8455b28988b058953e4b7824c4a45687d87dc41a962a20dba0758808`) is a
tracked-changes redline produced by **superdoc-redlines** (repair_prompt lane, jubarte
artifact 1234): 1,168 `w:del` and 4 `w:ins` over a Lithuanian ministry-of-education
order (references to orders "V-445", "V-689"). Microsoft Word refuses to open it — the
live probe saw the recovery dialog, and answering No was followed by the open error
and no document window. This repo's validator reports exactly the defect that causes
it: **3 `word-hyperlink-relationship` errors on `word/document.xml`**, one per
`w:hyperlink` whose `r:id` names no Relationship (I re-ran the validator this
session; exit 1, invalid). Live Word, the corpus manifest row, and the proven in-repo
assessment all agree: Word refuses this file, and the dangling hyperlink references
are the established cause.

## Word behaviour

- **Live result (this session's probe):** `repair-prompt`. Word showed
  `Word found unreadable content in superdoc-redlines__01a11a49_1234.docx. Do you
  want to recover the contents of this document?` (Yes/No); clicking No was followed
  by `Word experienced an error trying to open the file` and no document window
  opened. No crash report; Word quit cleanly. The live record carries `agrees: true`.
- **Manifest:** this repo has no `really_broken/manifest.csv` (the ask's path does
  not exist — verified). The source corpus manifest
  (`neurotic_docx_bench/maybe_future_corpus/really_broken/manifest.csv`, line 112)
  records `word:corrupt` — "Word could not load it in 2 passes: document loaded
  empty (Word could not read it)" — for the same sha256 as the copy in this repo
  (re-verified after copy). Both records are refusals.
- **Prior assessment: PROVEN** (`docs/really-broken/assessments.jsonl` line 123 and
  `docs/really-broken/inspection-log.jsonl` line 123, status `proven`). The
  minimal-change Word A/B experiment established the cause: **removing only the
  three `r:id` attributes from the `w:hyperlink` start tags — every other byte of
  every part unchanged — turns `unreadable-content-warning` into `clean-open`**.
  The assessment credits the repo check `word-hyperlink-relationship` with naming
  it. (The live run spells the refusal `repair-prompt` and the assessment
  `unreadable-content-warning`; both describe the same "found unreadable content"
  recovery dialog.)

## Violation: `hyperlink-rid-names-no-relationship` (3 sites, `word-hyperlink-relationship`)

**Offending XML** — `word/document.xml` lines 4835, 5292 and 62633 (verified by
reading the unpacked part, `really_broken/unpacked/
tool_redlines__repair_prompt__superdoc-redlines__01a11a49_1234/`). These are the only
three `w:hyperlink` elements in the whole package (no other part has one; the
`hyperlink` strings in `word/settings.xml`, `word/glossary/settings.xml` and
`word/styles.xml` are `w:clrMapping` attribute text and a style reference, not
elements). Each wraps a deleted cross-reference:

```xml
<w:hyperlink w:history="1" r:id="rId5b26125e">          <!-- line 4835 -->
  <w:del w:id="9256b8be-…" w:author="Live Bench" …>
    <w:r>
      <w:rPr><w:rStyle w:val="Hyperlink"/> …</w:rPr>
      <w:delText>V-689
      </w:delText>
    </w:r>
  </w:del>
</w:hyperlink>
<!-- likewise r:id="rId4b0463e5" at line 5292, r:id="rId3c8e1759" at line 62633 -->
```

**The dangle** — `word/_rels/document.xml.rels` declares exactly `rId1`–`rId15` and
nothing else: endnotes, three headers, three footers, fontTable, footnotes, settings,
styles, stylesWithEffects, theme, webSettings, one image. No relationship of any
hyperlink Type exists, and `rId5b26125e`, `rId4b0463e5`, `rId3c8e1759` match no
declared Id. So every `w:hyperlink` `r:id` in the document dangles.

**Schema rule** — `CT_Hyperlink` in `src/scripts/office/schemas/ISO-IEC29500-4_2016/
wml.xsd:1296-1304` declares `<xsd:group ref="EG_PContent" minOccurs="0"
maxOccurs="unbounded"/>`, optional `tgtFrame`/`tooltip`/`docLocation`/`history`/
`anchor` attributes, and `<xsd:attribute ref="r:id"/>`; and `r:id` in
`…/shared-relationshipReference.xsd:8-11` is:

```xml
<xsd:simpleType name="ST_RelationshipId">
  <xsd:restriction base="xsd:string"/>
</xsd:simpleType>
<xsd:attribute name="id" type="ST_RelationshipId"/>
```

`ST_RelationshipId` is a plain string, so the XSD alone cannot catch the dangle —
`r:id` is nonetheless *defined* as a relationship reference and must resolve to a
`Relationship` in the part's relationships part. I confirmed the XSD-invisibility
myself: `xmllint --noout --schema wml.xsd word/document.xml` exits 3 with 2680
findings, **zero** of which mention `hyperlink` (the 2680 are redline-tool noise the
assessment's experiment proves Word tolerates: GUID `w:id`s, undeclared
`w:authorEmail`/`w:rsidRPr` — `CT_TrackChange` at `wml.xsd:846-852` declares only
`author`/`date` — and `w:left` ordering). Only the dedicated check
`word-hyperlink-relationship` (`src/scripts/office/checks/word-hyperlink.ts:83`,
whose header documents Word session 28 / #614 pinning the two accepted hyperlink
Types) and the proven assessment anchor the cause.

**Proposed fix** — in **superdoc-redlines**: when wrapping content in a
`w:hyperlink`, either mint the backing relationship in the part's `.rels`
(`Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"`,
`TargetMode="External"`, `Target` = the carried-over URL, with a fresh sequential
`rIdN` that matches the existing convention), or — as the proven experiment
prescribes for repairing this artifact — emit no `r:id` at all (a `w:hyperlink` with
no `r:id` is left alone by the check and by Word). Fix target:
**superdoc-redlines**.

## Where the fixes belong and why

- **superdoc-redlines (generator bug):** the three Ids are visibly synthesized
  (`rId` + 8 hex chars, unlike the package's `rId1`–`rId15` convention) and were
  written into `w:hyperlink` start tags without ever adding Relationship entries.
  The generator deletes hyperlink-marked text ("V-689") but keeps a `w:hyperlink`
  wrapper pointing at a relationship it never minted.
- **docx-validate: no change.** The check exists, fires exactly 3× here with the
  correct message and remediation text, and the live Word result plus the proven
  assessment confirm its verdict. Raw XSD cannot see the defect
  (`ST_RelationshipId` is `xsd:string`), which is precisely why this dedicated
  check exists.
- **repair: not applicable as target** — the source document is not the broken
  party; the defect is introduced by the redline tool. For repairing the artifact
  itself, the prescription is the proven minimal change: delete the three `r:id`
  attributes, nothing else.

## Verification

Ran this session (all against the copy at
`really_broken/tool_redlines/repair_prompt/superdoc-redlines__01a11a49_1234.docx`,
sha256 re-verified = `53186d5e…758808`):

- `pnpm exec tsx src/scripts/office/validate.ts really_broken/tool_redlines/repair_prompt/superdoc-redlines__01a11a49_1234.docx --json`
  → `valid: false`; exactly **3 error issues**, all `word-hyperlink-relationship`
  on `word/document.xml` naming `rId5b26125e`, `rId4b0463e5`, `rId3c8e1759`; the
  other 1,690 issues are warnings/info — none error-severity.
- `xmllint --noout --schema src/scripts/office/schemas/ISO-IEC29500-4_2016/wml.xsd
  really_broken/unpacked/tool_redlines__repair_prompt__superdoc-redlines__01a11a49_1234/word/document.xml`
  → exit 3, 2680 findings, `grep -ci hyperlink` → **0**.
- Read `word/_rels/document.xml.rels` in full: only `rId1`–`rId15`, no hyperlink
  Type; grep confirmed exactly 3 `<w:hyperlink ` elements, all in `word/document.xml`.
- Carried established cause (assessment line 123, status `proven`): remove only the
  three `r:id` attributes → Word `clean-open` (original: `unreadable-content-
  warning`). Any *new* Word observation would need the repo's two-runs-with-logs
  bar; the recorded experiment already carries it.

Suggested regression guard for the tool: a fixture pair — this artifact as-is
(3 errors, Word refusal) and the r:id-stripped twin (validator clean, clean-open),
e.g. under `tests/fixtures/eigen/` with a manifest Word observation, plus a
generator-side test that every emitted `w:hyperlink` `r:id` has a matching
Relationship in the same part's `.rels`.
