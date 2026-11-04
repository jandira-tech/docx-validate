# [docxodus 2257] repair-prompt: bare `w:delText` in the mc:Choice text box

## Summary

`tool_redlines/repair_prompt/docxodus__01a11a49_2257.docx` is a docxodus-generated redline artifact
(source per manifest: `https://jubarte.pro/live/artifacts/01a11a49-7e73-770c-997d-0013581d67e9/2257/docxodus-redline.docx`,
sha256 `a9cb35a2b556fdfc381b6c6dce118cf520a1cc91d558ce57be1c348d1c833842`, verified against the packed file this session).
Its `word/document.xml` marks a whole drawing deleted with a `w:del` around the drawing run, but inside that drawing's
`mc:Choice` (`wps`) text box the four runs carry `w:delText` — deleted text with **no `w:del` anywhere inside the text box**.

- **Live Word (this run):** recovery dialog *"Word found unreadable content in docxodus__01a11a49_2257.docx. Do you want to
  recover the contents of this document?"*; answering **No** was followed by *"Word experienced an error trying to open the
  file"* and no document opened. Outcome `repair-prompt`.
- **Manifest row:** `word:corrupt` — *"Word could not load it in 2 passes: document loaded empty (Word could not read it) |
  harvest 20261013_071629"*. The harvest pass recorded the refusal without naming the dialog; the live run names it. Both
  agree Word refuses the file. (Note: `really_broken/manifest.csv` is absent from this checkout; the row was read from the
  corpus copy at `/Users/arthrod/temp/T/neurotic_docx_bench/maybe_future_corpus/really_broken/manifest.csv`.)
- **docx-validate (this run):** `valid: false` with exactly **4 error-severity issues, all `word-textbox-deleted-text`**,
  the only errors; the rest are warnings/info (`style-default-missing` ×4, `paraid-missing-element` ×3,
  `xml-bom-leading` ×2, `rel-ids-sequential` ×2).

This is a **semantic / Word-behavioural break, not an XSD content-model break**: raw `xmllint --noout --schema
ISO-IEC29500-4_2016/wml.xsd` over `word/document.xml` emits 3087 lines (all Microsoft-extension noise) and **zero** lines
mentioning `delText` (verified this session, `grep -c delText` = 0).

## Word behaviour

- **Live result:** repair prompt, then open error on refusal; no crash report; Word quit cleanly (ask's live observation,
  `agrees: true` with the proven assessment).
- **Prior assessment — PROVEN** (`docs/really-broken/assessments.jsonl:143`, `docs/really-broken/inspection-log.jsonl:143`,
  source `word-ab`):

| Bytes | Change | Word | docx-validate |
|---|---|---|---|
| original | none | `unreadable-content-warning` (repair prompt) | invalid: `word-textbox-deleted-text` |
| `2257-choice-plain.docx` | in the mc:Choice text box only, rename all four `w:delText` → `w:t` | **clean open** | **valid**, no errors |
| `2257-leave-first.docx` | leave the first `w:delText`, rename the other three | repair prompt again | invalid: `word-textbox-deleted-text` |

  The mc:Fallback text box keeps its four `w:delText` in the clean-opening variant — Word reads the `mc:Choice` branch and
  tolerates the Fallback's. One bare `w:delText` in the Choice box is enough to trigger recovery. Header/footer deleted
  text (inside `w:del` in the body of those parts) stayed and is tolerated.

## Violation — `word-textbox-deleted-text`: four bare `w:delText` in the mc:Choice `w:txbxContent`

**Offending XML** (`word/document.xml`, mc:Choice text box at byte offset 102073; ancestry
`w:body/w:p/w:del/w:r/mc:AlternateContent/mc:Choice/w:drawing/…/wps:txbx/w:txbxContent/w:p/w:r`). The four runs spell the
deleted text-box title "Bri" + "" + "dge" + "Office":

```xml
<w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/>
  <w:sz w:val="96"/><w:szCs w:val="96"/></w:rPr>
  <w:delText>Bri</w:delText></w:r>
… three further runs in the same w:txbxContent carry
  <w:delText xml:space="preserve"></w:delText>   (empty; w:sz 16, w:noProof, w:lang w:eastAsia="nl-NL")
  <w:delText>dge</w:delText>                     (Times New Roman, w:sz 96)
  <w:delText xml:space="preserve">Office</w:delText>  (Times New Roman, w:sz 96)
```

Verified this session by tree parse: that text box contains 4 `w:delText`, **0 `w:del` and 0 `w:t` inside it**. The
`w:del` that deletes the drawing sits *outside* the box, wrapping the whole `mc:AlternateContent` run. The Fallback text
box (offset 108748) holds four identical `w:delText` — excluded by the check's design and tolerated by Word.

**Schema rule** (read this session in `src/scripts/office/schemas/ISO-IEC29500-4_2016/wml.xsd`):

- `wml.xsd:1745-1758` — `EG_RunInnerContent` declares `delText` as ordinary run content, so a bare `w:delText` in any
  `w:r` is schema-legal (hence xmllint's silence):

  ```xml
  <xsd:group name="EG_RunInnerContent">
    <xsd:choice>
      <xsd:element name="t" type="CT_Text"/>
      …
      <xsd:element name="delText" type="CT_Text"/>     <!-- wml.xsd:1750 -->
  ```

- `wml.xsd:1782-1788` — `CT_R` takes that group unbounded: `<xsd:group ref="EG_RunInnerContent" minOccurs="0"
  maxOccurs="unbounded"/>`.
- `wml.xsd:997-1005` — `CT_RunTrackChange`, the type of `w:del` (`wml.xsd:3406`), is deleted text's intended container;
  the deleted-text contract is that `w:delText` lives in a run inside a `w:del`. Inside a text box the story is its own
  content tree: the outer `w:del` does not cover it (`src/scripts/office/checks/word-deleted-text.ts:14`, and the walk
  resets deletion state at `txbxContent`, `word-deleted-text.ts:57-59`). Outside that contract, in a text box, Word
  refuses the file.

**Proposed fix** — target **docxodus** (generator): when docxodus emits a deleted drawing whose text box carries text, the
runs inside `w:txbxContent` must use `w:t`, not `w:delText` — the deletion of the whole drawing is already expressed by
the `w:del` wrapping the drawing run. (Alternative the check also tolerates, per its fixture 0948: wrap the text-box runs
in a `w:del` inside the `txbxContent`. The proven experiment exercised the rename, so it is the recommended change.)
Apply it wherever the generator writes text-box content, not just this file's Choice branch; the Fallback twin is never
read by Word here but should match for consistency.

**Fix target:** docxodus.

## Where the fixes belong and why

- **docx-validate — no change.** The named check `word-textbox-deleted-text`
  (`src/scripts/office/checks/word-deleted-text.ts:99-127`) exists, fires exactly 4× here (once per element), and its
  severity (`error`, Word-blocking) matches the proven and live Word refusal. No gap, no false positive.
- **docxodus — generator bug.** The defect is in the XML docxodus emits: deleted-text state claimed inside a text-box
  story that no `w:del` within that box covers. The generator must change as above.
- **repair — not this case.** This is a tool-produced redline whose generator is at fault, not a broken source document
  needing a repair prescription.

## Verification

Run this session (paths from the repo root):

1. `pnpm exec tsx src/scripts/office/validate.ts really_broken/tool_redlines/repair_prompt/docxodus__01a11a49_2257.docx`
   → 4 `ERROR [word/document.xml]: A text box contains deleted text…` lines; `--json` → `valid: false`,
   4 error-severity issues, all `word-textbox-deleted-text`.
2. Unpacked with `pnpm exec tsx src/scripts/office/unpack.ts … --merge-runs false --simplify-redlines false`
   (the ask's named pre-unpacked directory had been pruned by a concurrent process; recreated from the packed file, whose
   sha256 matches the manifest). Tree parse of `word/document.xml`: 4 `w:txbxContent`; Choice box @102073 = 4 `w:delText`
   / 0 `w:del` / 0 `w:t` inside; Fallback box @108748 = 4 `w:delText` (excluded by design); other two boxes clean.
3. `xmllint --noout --schema src/scripts/office/schemas/ISO-IEC29500-4_2016/wml.xsd word/document.xml`
   → 3087 lines, 0 mentioning `delText` — confirms the break is semantic, not content-model.

Minimal-change experiment (already carried out and recorded, `inspection-log.jsonl:143`): in the mc:Choice text box only,
rename all four `w:delText` → `w:t` → Word opens clean and `validate()` is valid (`2257-choice-plain.docx`); leaving just
the first one (`2257-leave-first.docx`) still draws the repair prompt. A regression fixture for this shape — Choice-branch
`wps:txbx` with bare `w:delText` beside a tolerated Fallback twin — would pin the check's branch exclusion on a real
docxodus specimen; the existing synthetic fixtures 0044/2054/0948 already cover the VML/Choice/Fallback/inner-`w:del`
variants.
