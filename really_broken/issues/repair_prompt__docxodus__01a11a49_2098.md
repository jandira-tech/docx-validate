# [docxodus 2098] repair-prompt: w:delText inside w:moveFrom makes Word offer recovery

## Summary

`tool_redlines/repair_prompt/docxodus__01a11a49_2098.docx` (19,107 bytes, sha256 `ee8d83d62800426463ff9d7a9939fb0af2f3613f0db948a8099e5525c4cff475`, re-verified this session) is a tracked-changes redline produced by **docxodus** on a Czech form document (move-source text: `Jméno a příjmení`, `Rodné číslo`). Every revision mark is authored `Open-Xml-PowerTools` with epoch date `2000-01-01T00:00:00Z` — the docxodus signature. One move (range `moveFromRangeStart w:id="49" w:name="move1"`) wraps its three source runs in `w:moveFrom` but fills them with `w:delText` instead of `w:t`.

- **Live Word outcome (this run):** `repair-prompt` — *"Word found unreadable content in docxodus__01a11a49_2098.docx. Do you want to recover the contents of this document?"*; **No** → *"Word experienced an error trying to open the file"*, no document window, no crash report, Word quit cleanly.
- **Manifest:** the live result's `manifestSays` is empty, and `really_broken/manifest.csv` is no longer present in the working tree (the shared `really_broken/` tree was wiped mid-analysis; the analyst restored these exact bytes from the manifest row's artifact URL and the hash above still matches). No recorded manifest expectation to disagree with; the live outcome agrees with the prior assessment's `unreadable-content-warning`.
- **docx-validate** (run for this issue, see Verification): `--profile word-valid` exit 1 with exactly **3 ERROR lines, one per `w:moveFrom`, code `word-move-from-deleted-text`**. The default profile for DOCX/DOCM *is* word-valid, so the bare command fails identically.

## Word behaviour

- **Live result:** recovery prompt → open error; no crash; clean quit (above).
- **Prior assessment: PROVEN** — `docs/really-broken/assessments.jsonl` and `inspection-log.jsonl` (source `word-ab`), check `word-move-from-deleted-text`. Experiment matrix, each row observed in Word:

  | variant | Word outcome |
  |---|---|
  | original bytes | `unreadable-content-warning` |
  | repair moveFrom 2+3, leave 1 (`2098-leave-move1.docx`) | still warns |
  | repair moveFrom 1+3, leave 2 (`2098-leave-move2.docx`) | still warns |
  | repair moveFrom 1+2, leave 3 (`2098-leave-move3.docx`) | still warns |
  | **rename `w:delText` → `w:t` inside all three (`2098-move-all.docx`)** | **`clean-open`, validator valid** |

  So each of the three sites is independently sufficient to trigger the prompt, and the one-shot rename is the complete fix. The 84 `w:delText` under ordinary `w:del` wrappers and the 58 `w:del` inside `w:rPr` stayed in every variant and are Word-tolerated.

## Each violation

One violation, three sites, one part: check `word-move-from-deleted-text`, defined in `src/scripts/office/checks/word-deleted-text.ts:129-157` (the analyst's citation of `validators/docx.ts:1123-1147` is stale — that region now holds `validateStyleDefaults`; the check file's header documents the Word rule with fixture 0726, `word-deleted-text.ts:16-18`). All three sites are the same shape, re-verified by parsing the raw unzipped part this session — 3 `w:moveFrom` (w:id 45/46/47), each holding exactly 1 run with 1 `w:delText` and no `w:t`:

```xml
<w:moveFromRangeStart w:id="49" w:name="move1" w:author="Open-Xml-PowerTools" w:date="2000-01-01T00:00:00Z"/>
<w:moveFrom w:author="Open-Xml-PowerTools" w:id="45" w:date="2000-01-01T00:00:00Z">
  <w:r>
    <w:rPr><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr>
    <w:delText xml:space="preserve">Jméno a příjmení&#160;&#160;…</w:delText>
  </w:r>
</w:moveFrom>
<!-- w:id="46": identical shape, whitespace-only w:delText, xml:space="preserve".
     w:id="47": identical shape, w:delText "Rodné číslo" (no xml:space). -->
```

For contrast, the same generator's ten `w:moveTo` elements (w:id 382–391) each correctly wrap a `w:r/w:t` — the destination side of the same move uses the right payload; only the source side is wrong.

**Schema rule.** This shape is *XSD-valid*; Word's refusal is a semantic constraint the XSD cannot express, which is why the named check exists. Chain, quoted from the bundled `ISO-IEC29500-4_2016/wml.xsd`:

- `wml.xsd:3407` (EG_RunLevelElts): `<xsd:element name="moveFrom" type="CT_RunTrackChange"/>`
- `wml.xsd:997-1003` (CT_RunTrackChange): `<xsd:extension base="CT_TrackChange"><xsd:choice minOccurs="0" maxOccurs="unbounded"><xsd:group ref="EG_ContentRunContent"/>…`
- `wml.xsd:2091` (EG_ContentRunContent): `<xsd:element name="r" type="CT_R"/>`
- `wml.xsd:1785` (CT_R) refs `EG_RunInnerContent`, whose choice at `wml.xsd:1748` / `:1750` offers `<xsd:element name="t" type="CT_Text"/>` and `<xsd:element name="delText" type="CT_Text"/>` as context-free alternatives — nothing in the XSD ties `w:delText` to a `w:del`/`w:moveFrom` context.

Confirmed empirically: `xmllint --noout --schema src/scripts/office/schemas/ISO-IEC29500-4_2016/wml.xsd` on the raw `word/document.xml` → exit 3 with 80 findings, **all** `w14:paraId`/`w14:textId` Microsoft-extension attributes (undeclared in the transitional ISO schema), **zero** structural findings.

**Proposed fix (generator):** when docxodus emits a move source, serialize the moved text inside `w:moveFrom` runs as `w:t` (keeping `xml:space="preserve"` where the payload is whitespace-padded), exactly as it already does inside its `w:moveTo` runs. Delete-nothing: the `w:del`-wrapped `w:delText` blocks and the `w:del` run-property marks elsewhere in the document are correct redline output and must stay.

**Fix target:** docxodus (generator bug). docx-validate already detects this exactly — not a validator gap.

## Where the fixes belong and why

- **docxodus** — the redline's XML is genuinely Word-invalid at exactly these three runs, and the generator demonstrably knows the correct shape (its own `w:moveTo` output). The source document is fine; the defect is introduced when the move is serialized. Fix in the move-source emission path.
- **Not docx-validate** — the named check exists (`checks/word-deleted-text.ts:129-157`), is registered (`checks/index.ts:133`), is Word-blocking (`validators/word-blocking.ts:294`), fires 3× here under word-valid and strict, and its non-firing under explicit lenient is documented, intended behaviour (Word rules don't run under lenient; pinned by `tests/word-move-from-deleted-text.test.ts:32-36`).
- **Not repair** — the source document is not the broken artefact.

One adjacent observation, out of scope for this issue: strict also reports `58 <w:p> element(s) lack a w14:paraId` in this part — a separate finding about the same file, not part of the Word-refusal cause proven above.

## Verification

Run this session on the restored bytes (hash re-verified before and after):

1. `pnpm exec tsx src/scripts/office/validate.ts --profile word-valid really_broken/tool_redlines/repair_prompt/docxodus__01a11a49_2098.docx` → exit 1, exactly 3 ERROR lines `A move source contains deleted text… Use w:t inside w:moveFrom, not w:delText.` The bare command (default word-valid for DOCX) fails identically; `--profile strict` → exit 1 (same 3 + the paraId finding); `--profile lenient` → `All validations PASSED!`, exit 0, by design.
2. Raw-byte census (plain `unzip`, not the repo's transform-applying unpack): 3 `w:moveFrom`, each 1 run / 1 `w:delText` / 0 `w:t`; 87 `w:delText` total = 3 in-moveFrom + 84 tolerated; 10 `w:moveTo` each 1 `w:t`; 58 `w:del` inside `w:rPr` — matches the assessment's counts.
3. `pnpm exec vitest run tests/word-move-from-deleted-text.test.ts` → **4 passed**, covering rejection of one `w:delText` in a `w:moveFrom` (`tests/fixtures/broken/package.word-move-from-deleted-text.docx`), the clean `w:t` twin (`tests/fixtures/working/package.word-move-from-text.docx`), and the lenient/strict profile split.
4. The proven minimal-change experiment to re-run against a fixed docxodus build: emit the same move with `w:t` inside `w:moveFrom`, open in Word twice with logs — expected `clean-open` (the assessment's `2098-move-all.docx` already showed it), and `validate()` expected valid.
