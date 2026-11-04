# [docxodus 0070] failed: bare w:delText inside text boxes triggers Word recovery

## Summary

- **Document**: `tool_redlines/repair_prompt/docxodus__01a11a49_0070.docx` — a docxodus-produced
  redline (lane category `repair_prompt`, label "docxodus 0070"). Reported defect: 38 bare
  `w:delText` elements inside 10 `w:txbxContent` text-box stories in `word/document.xml`
  (19 in the three `mc:Choice` `wps:txbx` boxes Word renders + 19 in their `mc:Fallback`
  `v:textbox` twins), which makes Word offer to recover the file.
- **Status: disputed — the artifact is missing from the working tree, so the flagged sites
  could not be independently re-verified.** Verified in this session:
  - `really_broken/tool_redlines/repair_prompt/` contains docxodus specimens 1951–2397 but
    **no `docxodus__01a11a49_0070.docx`**; a repo-wide `find … -name "*0070*"` returns nothing;
    no unpacked `really_broken/unpacked/*0070*` directory exists (the unpacked dirs are named
    `tool_redlines__repair_prompt__docxodus__01a11a49_NNNN`; none ends in `0070`).
  - The analyst's exact command now fails:
    `pnpm exec tsx src/scripts/office/validate.ts really_broken/tool_redlines/repair_prompt/docxodus__01a11a49_0070.docx`
    → `Error: really_broken/tool_redlines/repair_prompt/docxodus__01a11a49_0070.docx does not exist`.
  - `really_broken/` is untracked (`git status` → `?? really_broken/`), so the file was never
    committed and cannot be recovered from git.
- **Live Word vs manifest**: the live Word check could not open the file ("The file does not
  exist … Word was never opened", `outcome: failed`, `agrees: false`). `really_broken/manifest.csv`
  also does not exist, so the manifest row quoted in the ask (category `repair_prompt`,
  `d1_failed_in soffice`, local_word_error "Word could not load it in 2 passes: document loaded
  empty") **could not be cross-checked**. What does exist is the proven-cause record
  (`docs/really-broken/assessments.jsonl:31`, mirrored at `inspection-log.jsonl:31`), which
  independently corroborates the analyst's anatomy.

## Word behaviour

- **Proven assessment** (`docs/really-broken/assessments.jsonl:31`, status `proven`):
  `word/document.xml` contains 10 `w:txbxContent` elements with 38 `w:delText` elements between
  them; original Word verdict `unreadable-content-warning`. Minimal change that opened clean:
  *inside those 10 `w:txbxContent` elements only*, rename `w:delText` to `w:t`, changing nothing
  else — the 961 other `w:delText` elements in the same part remain, and Word opens clean
  (`text_box_delText_changed_to_w_t` → `clean-open`).
- **Live Word this run**: not run — the file does not exist (above). The recorded assessment
  stands as the only Word evidence; it was made before the file disappeared.
- **This repo's check agrees**: `word-textbox-deleted-text`
  (`src/scripts/office/checks/word-deleted-text.ts:99-127`) — "Deleted text in a text box is
  wrapped in a deletion within that story" — fires one error per bare `w:delText`/`w:delInstrText`
  inside a `w:txbxContent` not wrapped in a `w:del *within that text box*`, and excludes
  `mc:Fallback` copies (`word-deleted-text.ts:62`), matching the analyst's 19 flagged sites
  (Choice boxes only) vs 38 total. Its evidence base is documented at
  `word-deleted-text.ts:6-19` (fixtures 0011/0044 VML+DrawingML shapes, 2054 Fallback tolerated,
  0948 in-box `w:del` tolerated). It is Word-blocking (`validators/word-blocking.ts:293`) and
  explained at `validators/word-error-explanations.ts:97-98`. Note: the analyst's citations
  `docx.ts:1082-1116` / `word-error-explanations.ts:66-67` are stale; at HEAD the check lives in
  `checks/word-deleted-text.ts`.

## Each violation

### 1. `word-textbox-deleted-text` — bare `w:delText` in text-box stories (`word/document.xml`)

**Offending XML** (analyst's excerpt; **could not be re-read — the part is gone**; consistent
with the proven assessment's 10 boxes / 38 elements):

```xml
<wps:txbx>
  <w:txbxContent>
    <w:p w14:paraId="4515BD6A" w14:textId="77777777">
      <w:r>
        <w:rPr>…</w:rPr>
        <w:delText xml:space="preserve">EMINENT
        </w:delText>
      </w:r>
      <!-- 8 more bare w:delText runs in this text box; no w:del anywhere inside the w:txbxContent -->
```

**Schema rule** (read this session; `src/scripts/office/schemas/ISO-IEC29500-4_2016/wml.xsd`):

```xml
<!-- wml.xsd:1745-1750 -->
<xsd:group name="EG_RunInnerContent">
  <xsd:choice>
    <xsd:element name="br" type="CT_Br"/>
    <xsd:element name="t" type="CT_Text"/>
    <xsd:element name="contentPart" type="CT_Rel"/>
    <xsd:element name="delText" type="CT_Text"/>   <!-- permitted in ANY w:r -->
    …
<!-- wml.xsd:3378-3381 -->
<xsd:complexType name="CT_TxbxContent">
  <xsd:group ref="EG_BlockLevelElts" minOccurs="1" maxOccurs="unbounded"/>
</xsd:complexType>
<xsd:element name="txbxContent" type="CT_TxbxContent"/>
```

The XSD places no co-occurrence constraint between `w:delText` and `w:del` — XSD 1.0 cannot
express it — so this construct is schema-valid and the defect is a **Word-runtime tracked-change
rule**: deleted text in a text-box story must be covered by a `w:del` inside that same text box.
A `w:del` outside the box does not cover it (`word-deleted-text.ts:14`, `:60-64`).

**Proposed fix**: docxodus must not emit bare `w:delText` inside `w:txbxContent`. Either wrap the
deleted runs in a `w:del` inside the text-box story, or emit plain `w:t` there (the proven
minimal repair: rename the 38 `w:delText` → `w:t` inside the 10 boxes only → clean open, with the
961 body `w:delText` untouched).

**Fix target**: `docxodus` (generator bug).

## Where the fixes belong and why

- **docxodus (generator)**: the redline's XML is genuinely Word-refusing — tracked deletions were
  serialized into text-box stories without the enclosing `w:del` the text-box story needs. This
  is the only code that must change.
- **docx-validate — no gap**: the check exists, is Word-blocking, and fires at the **default**
  CLI profile. Demonstrated this session on the sibling specimen
  `docxodus__01a11a49_1951.docx` (same defect class):
  `pnpm exec tsx src/scripts/office/validate.ts really_broken/tool_redlines/repair_prompt/docxodus__01a11a49_1951.docx --json`
  → `Counter({'word-textbox-deleted-text': 29, …}), valid: False` (3 in `word/document.xml`
  matching an independent Python count of 3 flagged in that part, 18 in footer1, 8 in footer2),
  all at the default profile (`word-valid` is the `.docx` default, `validate.ts:1600`). The
  analyst's claim that the defect is invisible "at the default lenient profile" does not hold at
  HEAD; under an explicit `--profile lenient` the Word rules intentionally do not run
  (`validate.ts:836-854`), which is documented behaviour, not a miss. No `repair` target: this is
  generator output, not a broken source document.

## Verification

1. **Re-materialize the specimen** (or re-point the lane): the 0070 bytes exist nowhere in the
   repo and were never committed; until they do, every site-level claim below rests on the
   analyst's run and the proven record, not on re-inspection. That is why this issue is disputed.
2. **Re-run on the recovered file**:
   `pnpm exec tsx src/scripts/office/validate.ts really_broken/tool_redlines/repair_prompt/docxodus__01a11a49_0070.docx --json`
   — expect ~19 `word-textbox-deleted-text` errors on `word/document.xml` (Choice boxes only;
   the Fallback twins are excluded by design) and `valid: false`.
3. **Minimal-change experiment (already proven, assessments.jsonl:31)**: inside the 10
   `w:txbxContent` only, rename `w:delText` → `w:t`; Word opens clean with the 961 body
   `w:delText` still present. Repaired twin fixture exists as
   `tests/fixtures/working/package.word-textbox-deleted-text-removed.docx`.
4. **Check regression on a living sibling** (run this session, reproduces the defect class):
   the 1951 command above; 29 errors, `valid: False`.

*Checks not run this session: live Word open (file missing), raw `xmllint` on the 0070 part
(file missing — but the cited XSD group was read directly and permits the construct, so no
schema error was expected).*
