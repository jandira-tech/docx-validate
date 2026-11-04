# [superdoc-redlines 0900] repair-prompt: three uncovered `[trash]/*.dat` parts

## Summary

The document is a redline `.docx` produced by **superdoc-redlines** (jubarte run `01a11a49`, job 900,
artifact `superdoc-redlines-redline.docx`; sha256 `75057928152c552e9b9767eec4768a1af396f2df1557ba8aaa148b8afdf9c146`,
re-verified this session). Beside an otherwise ordinary WordprocessingML package it stores a
`[trash]/` folder with three parts — `0000.dat` (392 B), `0001.dat` (2496 B), `0002.dat` (652 B),
stored uncompressed; their bytes are U+FFFD mojibake of a UTF-16 BOM followed by NUL padding
(deleted-redline residue, confirmed by hex dump this session).

The defect: **none of the three parts resolves to a content type.** `[Content_Types].xml` has no
`<Override PartName="/[trash]/...">` and no `<Default Extension="dat">`, and no `.rels` part names
any of them. OPC requires a content type for every part; Word offers to recover the file.

Live Word outcome (this session): **repair-prompt** — "Word found unreadable content in
superdoc-redlines__01a11a49_0900.docx. Do you want to recover the contents of this document?";
answering No gave "Word experienced an error trying to open the file" and no document window.
This agrees with the proven assessment and the corpus manifest's earlier observation.

## Word behaviour

- **Live run** (ask-supplied result): recovery dialog → outcome `repair-prompt`, `agrees: true`.
- **Manifest row** — no `really_broken/manifest.csv` exists in this repo (verified: `grep` exit 1).
  The row lives in the sibling corpus
  `/Users/arthrod/temp/T/neurotic_docx_bench/maybe_future_corpus/really_broken/manifest.csv`:
  category `repair_prompt`, tool `superdoc-redlines`, same sha256, source
  `https://jubarte.pro/live/artifacts/01a11a49-7e73-770c-997d-0013581d67e9/900/superdoc-redlines-redline.docx`,
  prior soffice observation "Word could not load it in 2 passes: document loaded empty (Word could
  not read it)".
- **Proven assessment** (`docs/really-broken/assessments.jsonl` and `docs/really-broken/inspection-log.jsonl`,
  path `tool_redlines/repair_prompt/superdoc-redlines__01a11a49_0900.docx`, status `proven`,
  5 Word A/B runs):
  - original → `unreadable-content-warning`
  - `0900-leave-0000.docx` / `-0001` / `-0002` (exactly one `.dat` kept each) →
    `unreadable-content-warning` — **each part alone is sufficient for the repair prompt**
  - `0900-drop-all.docx` (all three `.dat` deleted, the empty `[trash]/` directory entry kept) →
    `clean-open`, and docx-validate returned `valid: true`
  - The empty `[trash]/` directory entry is **not** a cause.

## Violation — `ct-uncovered-part` × 3 (`[trash]/0000.dat`, `[trash]/0001.dat`, `[trash]/0002.dat`)

### Offending package state (read this session)

zip entries (`zipinfo -l` on the packed file, run this session):

```text
-rw----  2.0 fat  392 b-   392 stor  [trash]/0000.dat
-rw----  2.0 fat 2496 b-  2496 stor  [trash]/0001.dat
-rw----  2.0 fat  652 b-   652 stor  [trash]/0002.dat
drwx---  2.0 fat    0 b-     0 stor  [trash]/
```

`[Content_Types].xml:3-5` declares only these Defaults:

```xml
<Default Extension="png" ContentType="image/png"/>
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
```

followed by 26 `<Override>` entries (lines 6–31) covering only `word/…`, `customXml/…` and
`docProps/…` parts — no `Extension="dat"` (`grep -i dat` over the stream: no match) and no
`PartName="/[trash]/…"`. `grep -ril trash` across every file in the unpacked package returns
nothing: no relationships part or XML part references them.

### Rule (a package-level OPC invariant, not an XSD line)

- `src/lib/content-types.ts:17-22` — "OPC content-type resolution (ECMA-376 Part 2 §10.1.2.4).
  A part's content type is resolved by an exact `<Override PartName>` match first, then by a
  `<Default Extension>` matching the part name's extension."
- `src/scripts/office/validators/base.ts:1664-1669` — "Residual OPC-coverage pass. … This pass
  enforces the underlying spec invariant instead: every part must be matched by an
  `<Override PartName>` or a `<Default Extension>` in the content-types stream." The pass
  (`base.ts:1673-1692`) emits `severity: "error"`, code `ct-uncovered-part`, one issue per
  uncovered part.
- The bundled `src/scripts/office/schemas/ecma/fourth-edition/opc-contentTypes.xsd:7-16` declares
  `CT_Types` as a plain choice of `Default`/`Override` elements — it can express the vocabulary
  but not the coverage invariant. Run this session:
  `xmllint --noout --schema src/scripts/office/schemas/ecma/fourth-edition/opc-contentTypes.xsd
  'really_broken/unpacked/tool_redlines__repair_prompt__superdoc-redlines__01a11a49_0900/[Content_Types].xml'`
  → "validates" (exit 0). The stream is schema-clean; the **package** breaks the OPC rule.

### Re-verified this session (command + output)

```text
pnpm exec tsx src/scripts/office/validate.ts \
  really_broken/tool_redlines/repair_prompt/superdoc-redlines__01a11a49_0900.docx --json
→ valid: false; error-severity issues: 3
  ct-uncovered-part | [trash]/0000.dat
  ct-uncovered-part | [trash]/0001.dat
  ct-uncovered-part | [trash]/0002.dat
```

Remaining findings are warnings/info only (`rels-unreferenced` × 3 for the same parts,
`paraid-missing-element` in `word/glossary/document.xml`, `style-default-missing` × 2,
`xml-syntax` non-absolute xmlns warnings in `customXml/item5.xml`, `rel-ids-sequential` info) —
none is error-severity, and the proven assessment attributes the Word repair prompt solely to the
`.dat` parts.

### Proposed fix — fix target: **superdoc-redlines** (generator)

Stop emitting parts the package cannot describe:

1. **Primary (Word-proven):** do not write the `[trash]/` parts at all. They are unreferenced
   deleted-redline residue; deleting all three is the proven clean-open
   (`0900-drop-all.docx`, keeping the empty directory entry).
2. **Spec-compliant alternative (not Word-probed for this file):** if the residue parts must be
   kept, cover them in `[Content_Types].xml` — e.g.
   `<Default Extension="dat" ContentType="application/octet-stream"/>`, or one `<Override>` per
   part. OPC-legal, but only deletion is proven against Word for this document.

## Where the fixes belong and why

- **superdoc-redlines** — the generator wrote unreferenced `.dat` parts with no content-type
  coverage. No single XML part of the redline is malformed; the *package* it produced violates
  OPC. Its packing step must either drop residue parts or maintain `[Content_Types].xml` in
  lockstep with every zip entry it writes.
- **docx-validate** — no gap and no false positive: the `ct-uncovered-part` check exists, reports
  exactly one error per offending part, and its verdicts match the proven Word A/B (each part
  alone warns; all deleted opens clean). Nothing to change here.
- Not "repair" — the source document is not a broken input; this package is the generator's own
  output.

## Verification

- Re-run the validator: `pnpm exec tsx src/scripts/office/validate.ts
  really_broken/tool_redlines/repair_prompt/superdoc-redlines__01a11a49_0900.docx --json` →
  exactly the three `ct-uncovered-part` errors above.
- Minimal-change experiments already proven (inspection-log rows, two Word runs each): keep-one
  variants each still warn; `0900-drop-all.docx` opens clean and validates `true`. No new
  experiment is needed to establish the cause.
- Existing regression suites covering this check: `tests/word-undeclared-part.test.ts`,
  `tests/content-types.integration.test.ts`, `tests/validators-base.test.ts`,
  `tests/word-blocking.test.ts`.
- For the generator fix: rebuild the redline for the same input and assert the zip contains no
  `[trash]/` file entries (or that `[Content_Types].xml` covers every stored entry), then confirm
  a dialog-free Word open. Word verification was **not** run in this session; the live
  repair-prompt observation quoted above is the ask-supplied result.
