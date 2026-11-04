# [jubarte 0792] repair-prompt: cell ends with nested `w:tbl`, no trailing `w:p`

## Summary

`tool_redlines/repair_prompt/jubarte__01a11a49_0792.docx` (sha256
`c63ee16523e32bfaa8d12d5abfc5a33fcfafa0667c20341909633a259261f21`, verified with
`shasum -a 256` this session) is a jubarte redline artifact — a Spanish-language schedule
document ("JUEVES 12 DE SEPTIEMBRE") carrying a tracked-changes redline
(`w:author="Redline"`, `w:date="1970-01-01T00:00:00Z"`): 149 table cells across 56
`w:tbl` (9 body-level), plus header/footer/notes parts.

Live Word outcome: **repair-prompt** — the recovery dialog "Word found unreadable
content in jubarte__01a11a49_0792.docx"; answering No was followed by "Word
experienced an error trying to open the file" and no document window. This agrees with
the manifest, which records the same file as a Word refusal at the coarser pass-1
category (`d1_failed_in = word:corrupt`, `local_word_error = "Word could not load it in
2 passes: document loaded empty (Word could not read it)"`). One defect explains the
refusal, and it is proven by an A/B Word experiment.

Note on the manifest's location: this repo checkout's `really_broken/` has no
`manifest.csv`; the row was read in the sibling corpus at
`/Users/arthrod/temp/T/neurotic_docx_bench/maybe_future_corpus/really_broken/manifest.csv`
(its sha256 column matches the repo file's hash above).

## Word behaviour

- **Live (this run):** repair-prompt — "Word found unreadable content in
  jubarte__01a11a49_0792.docx. Do you want to recover the contents of this document?";
  No → open error; no crash report; Word quit cleanly. `agrees: true`.
- **Manifest row:** `tool_redlines/repair_prompt/jubarte__01a11a49_0792.docx`,
  category `repair_prompt`, tool `jubarte`, `d1_failed_in: word:corrupt`.
- **Prior assessment (proven):** `docs/really-broken/assessments.jsonl:116` and
  `docs/really-broken/inspection-log.jsonl:116` (`source: word-ab`, status `proven`):
  - original → `unreadable-content-warning`;
  - `0792-add-p.docx` — insert an empty `<w:p/>` after that cell's last `</w:tbl>` and
    before its `</w:tc>` → **clean-open**;
  - `0792-cell0-no-p.docx` — remove the trailing `w:p` from a clean cell-0 isolate so
    that cell ends with `w:tbl` → `unreadable-content-warning` again.

## Violation 1 — `word-cell-ends-with-table` (1 site, confirmed)

**Offending XML** — `word/document.xml`, bytes 85141–156067 (70,926 bytes), lines
2303–4153. Of 149 `w:tc` in the package, exactly one has `w:tbl` as its last element
child; its children are `{tcPr: 1, p: 1, tbl: 39}`. It sits in the first body-level
table (document-order `w:tbl` #1 of 56), row 0, cell 1. The assessment records the same
cell as "body table 40, row 0, cell 1"; row and cell agree, the table ordinal reads
differently under this session's counting (first body-level table).

```xml
<w:tc>
  <w:tcPr>
    <w:tcW w:w="3589" w:type="dxa"/>
    <w:tcPrChange w:id="156" w:author="Redline" w:date="1970-01-01T00:00:00Z">…</w:tcPrChange>
  </w:tcPr>
  <w:p>…JUEVES 12 DE SEPTIEMBRE…</w:p>   <!-- "JUEVES…" at line 2327 -->
  <w:tbl>…</w:tbl>                        <!-- 39 nested tables -->
  <!-- no w:p after the last w:tbl -->
</w:tc>
```

The preceding sibling cell ends `…</w:del></w:p></w:tc>` and the next cell begins
`<w:tc><w:tcPr><w:tcW w:w="3589" w:type="dxa"/>…` (read at bytes 85080–85260).

**Schema rule.** The bundled Transitional XSD does *not* forbid this shape — the
trailing-paragraph requirement is Word's application-level rule. Quoted from
`src/scripts/office/schemas/ISO-IEC29500-4_2016/wml.xsd:2357-2363`:

```xsd
<xsd:complexType name="CT_Tc">
  <xsd:sequence>
    <xsd:element name="tcPr" type="CT_TcPr" minOccurs="0" maxOccurs="1"/>
    <xsd:group ref="EG_BlockLevelElts" minOccurs="1" maxOccurs="unbounded"/>
  </xsd:sequence>
  <xsd:attribute name="id" type="s:ST_String" use="optional"/>
</xsd:complexType>
```

and `EG_BlockLevelElts` (`wml.xsd:3393-3398`) is a plain `xsd:choice`:

```xsd
<xsd:group name="EG_BlockLevelElts">
  <xsd:choice>
    <xsd:group ref="EG_BlockLevelChunkElts" minOccurs="0" maxOccurs="unbounded"/>
    <xsd:element name="altChunk" type="CT_AltChunk" minOccurs="0" maxOccurs="unbounded"/>
  </xsd:choice>
</xsd:group>
```

`xmllint --noout --schema wml.xsd word/document.xml` (run this session) emitted 147
error lines, none naming `tc` — they are the tolerated shapes this validator already
carries as warnings (46 `bottom` cell-margin order, 91 `gridCol w:type`, 9
`shd w:fill=""`), whose counts exactly match the run's warning-severity findings. So
`CT_Tc`'s loose group model permits a cell ending with `w:tbl`; Word does not.

**Proposed fix.** Insert an empty `<w:p/>` after that cell's last `</w:tbl>` and before
its `</w:tc>`. Proven sufficient: the single insertion (`0792-add-p.docx`) opened clean
with all nested tables' widths, alignment, revision marks and text unchanged, and the
converse experiment (removing a trailing `w:p` from a clean cell) re-triggered the
warning — so the missing trailing paragraph, not the 39 nested tables, is the cause.

**Fix target: jubarte** (generator). The redline's XML is invalid-as-Word-reads-it and
the generator must change: whenever jubarte emits or wraps a nested `w:tbl` inside a
`w:tc` (here, redline markup around a schedule grid), it must close the cell with a
`w:p` after the last block-level table — e.g. always append an empty `<w:p/>` (or
`<w:p><w:pPr><w:rPr>…</w:rPr></w:pPr></w:p>` matching cell formatting) when the last
child written for a cell is a `w:tbl`.

## Where the fixes belong and why

- **jubarte** — the only code change needed. It produced a `w:tc` whose block-level
  content ends with `w:tbl`; every other of the 149 cells it emitted ends with `w:p`,
  so its writer already knows the rule and misses it on this nested-table path.
- **docx-validate — no change required.** The defect is already detected: the named
  check `word-cell-ends-with-table` (`src/scripts/office/checks/word-cell-ending.ts:22`,
  registered in `WORD_RULE_CHECKS`) reports exactly one error-severity issue for this
  file, and its header comment cites this 0792 fixture and the insert/remove
  experiments. The check exists *because* the XSD cannot catch this shape (see quoted
  `CT_Tc`); that is the correct division of labour, not a validator gap.
- **repair (source document)** — not applicable; the source is a generated redline, and
  the provenance is a generator defect, not a broken source document.

## Verification

Commands run this session (all from the repo root):

1. `shasum -a 256 really_broken/tool_redlines/repair_prompt/jubarte__01a11a49_0792.docx`
   → `c63ee165…261f21`, equal to the corpus manifest row's sha256.
2. Structural re-measure with `python3 -I` over the unpacked
   `really_broken/unpacked/tool_redlines__repair_prompt__jubarte__01a11a49_0792/word/document.xml`:
   149 `w:tc`; exactly 1 whose last element child is `w:tbl`; children
   `{tcPr:1, p:1, tbl:39}`; cell bytes 85141–156067, lines 2303–4153; first body-level
   table, row 0, cell 1; `tcPrChange w:id="156" w:author="Redline"` at byte 85227;
   exactly one whitespace-tolerant `</w:tbl>\s*</w:tc>` closing in the file.
3. `pnpm exec tsx src/scripts/office/validate.ts really_broken/tool_redlines/repair_prompt/jubarte__01a11a49_0792.docx`
   → `valid: false`, profile `word-valid`, exactly one ERROR:
   `word-cell-ends-with-table` on `word/document.xml` (via `--json`: 1 error, plus
   non-blocking warnings 91 `word-grid-column-attributes`, 46 `word-cell-margin-order`,
   9 `word-shading-color`, 4 `paraid-missing-element`, 3 info `rel-ids-sequential`).
4. `xmllint --noout --schema src/scripts/office/schemas/ISO-IEC29500-4_2016/wml.xsd <unpacked>/word/document.xml`
   → 147 error lines, 0 naming `tc`; error elements tally 46 `bottom` / 91 `gridCol` /
   9 `shd`, matching item 3's warning counts — confirming the tc shape is XSD-legal.
5. `pnpm exec vitest run tests/word-cell-ends-with-table.test.ts tests/checks-word-cell-ending.test.ts`
   → 2 files, 7 tests passed (broken fixture rejected, repaired control accepted,
   lenient silent, strict reports).

Minimal-change experiment for any jubarte-side fix: regenerate (or hand-edit, as in
`0792-add-p.docx`) so the cell ends `</w:tbl><w:p/></w:tc>`, and confirm two clean Word
opens of the same bytes per the repo's two-run rule; conversely, removing that `w:p`
must re-trigger the recovery dialog, as `0792-cell0-no-p.docx` showed.
