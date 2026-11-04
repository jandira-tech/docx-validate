# [source e98faffd] failed: five undeclared bytes follow the ZIP EOCD record

## Summary

`fixtures/repair_prompt/e98faffd73f88360e7dfd7560869929898f0785c44e28c6daf429b3c714d4ee6.docx`
is a **source document** (lane `repair_prompt`, producing tool `source-doc`) — not a tool redline.
Its XML parts are clean; the defect is in the **ZIP envelope**: five undeclared bytes
(`00 0D 0A 0D 0A`) follow the end-of-central-directory record, whose comment-length field is 0.
Word offers to recover the document; deleting exactly those five bytes is the recorded minimal
change that opens it clean.

**Provenance caveat, verified first-hand this session: the specimen is gone from the repo.**
`really_broken/fixtures/` does not exist, repo-wide `find … -name "*e98faffd*"` (excluding
`node_modules`, `.git`) returns nothing, the unpacked copy
`really_broken/unpacked/analyst-source__repair_prompt__…` does not exist, and
`really_broken/manifest.csv` does not exist. Only the proven-cause records
(`docs/really-broken/assessments.jsonl:13`, `docs/really-broken/inspection-log.jsonl:13`) and the
analyst's byte-level parse survive. The live Word run therefore never opened anything (outcome
`failed`) — that is an artifact of the missing file, not evidence against the defect.

## Word behaviour

- **Live result (this ask):** outcome `failed` — "the file does not exist … so Word was never
  opened". Absence re-verified myself: `ls really_broken/` shows only `issues/ tool_redlines/
  unpacked/`; the `find` above returns nothing.
- **Manifest row:** unreadable — `really_broken/manifest.csv` does not exist (checked with `ls`).
- **Prior proven assessment — the established Word cause** (`docs/really-broken/assessments.jsonl:13`,
  status `proven`; mirrored at `docs/really-broken/inspection-log.jsonl:13`): defect
  *"Five undeclared bytes follow the ZIP end-of-central-directory record. The record's comment
  length is 0. The bytes are 00 0D 0A 0D 0A."*; minimal change *"Delete those five bytes so the
  end-of-central-directory record is the last record in the file. Every preceding byte stays the
  same."*; Word: `original: unreadable-content-warning` → `after_this_change: clean-open`.
- **Recorded geometry** (analyst byte parse, ask-given, internally consistent — I hand-decoded the
  hexdump field by field): size 16,144 (0x3F10); EOCD signature `50 4B 05 06` at 16,117 (0x3EF5);
  disk 0; entries 11/11; cd_size 705 at 0x3F01; cd_offset 15,412 (0x3C34) at 0x3F05;
  comment_len 0 at 0x3F09; record ends 16,139 (0x3F0B); cd_offset+cd_size = 16,117 = EOCD offset,
  so the archive itself is consistent and only the 5-byte tail is foreign.

## Violation 1 — `undeclared-bytes-after-eocd` (ZIP envelope; 1 site)

**Offending bytes** (analyst hexdump, ask-given; the original file is gone, so this is the
surviving record of the site):

```text
00003ef0: 702e 786d 6c50 4b05 0600 0000 000b 000b  p.xmlPK.........
00003f00: 00c1 0200 0034 3c00 0000 0000 0d0a 0d0a  .....4<.........
```

EOCD signature at 0x3EF5; comment length `0000` at 0x3F09–0x3F0A; record ends at 0x3F0B; bytes
0x3F0B..0x3F0F = `00 0D 0A 0D 0A` follow it, belonging to no ZIP structure. The tail reads as
NUL + CR LF CR LF — the shape of a text-appended blank-line separator (e.g. a packaging script
writing `\r\n\r\n` after the ZIP stream), which is why the XML parts are clean.

**Rule.** No bundled XSD governs the ZIP envelope (`src/scripts/office/schemas/` covers XML parts
only), so the normative reference is the PKWARE `.ZIP File Format Specification` (APPNOTE.TXT),
re-fetched and quote-checked this session from
<https://pkware.cachefly.net/webdocs/casestudies/APPNOTE.TXT>:

- §4.3.6, overall .ZIP file format: the record sequence ends
  "… [zip64 end of central directory locator] **[end of central directory record]**" — the EOCD
  is the last element of the archive layout.
- §4.3.16, end-of-central-directory record field table: "**.ZIP file comment length  2 bytes**"
  and "**.ZIP file comment  (variable size)**" — the comment is the record's final field.
- §4.4.25: "**The length of the comment for this .ZIP file.**" — with comment length 0, zero
  bytes may follow the record.

**Re-verification this session** (original artifact absent — the defect was reproduced on a
byte-shape reconstruction, plus the rule and check verified in source):

1. Appended `00 0D 0A 0D 0A` to a known-clean fixture
   (`tests/fixtures/working/package.drawing-arc-radius-integer.docx` → `/tmp/e98f-verify/repro.docx`).
2. `python3 -I` struct parse of the repro: `comment_len=0`, record ends 11,904, trailing
   `5 bytes: 00 0d 0a 0d 0a`, `cd_offset+cd_size == EOCD offset` — the recorded shape exactly.
3. `pnpm exec tsx src/scripts/office/validate.ts /tmp/e98f-verify/repro.docx` (default profile,
   word-valid for DOCX) → `ERROR: ZIP end of central directory is followed by 5 undeclared
   byte(s). Word offers to recover the document. …`, exit 1.
4. `… --profile lenient` → `All validations PASSED!` (by design: envelope checks are excluded from
   lenient, `src/scripts/office/validate.ts:853-857`).
5. Controls: the unmodified fixture passes `--profile word-valid`; the repro with the 5 bytes
   deleted (`repro.docx[:-5]`) passes `--profile word-valid`.

The repo's check for this defect is **`zip-bytes-after-eocd`**
(`src/scripts/office/checks/package.ts:16-33`; logic `countZipBytesAfterEocd`,
`src/scripts/office/zip-bytes-after-eocd.ts:152-180`; registered
`src/scripts/office/checks/index.ts:146`; `archiveOnly: true`, so it runs under word-valid and
strict). It exists, fires, and says exactly what to do — **no docx-validate gap**.

**Proposed fix (target: repair).** Truncate the document to 16,139 bytes so the EOCD record is
the file's last structure — i.e. delete bytes 0x3F0B..0x3F0F (`00 0D 0A 0D 0A`), every preceding
byte unchanged. That is the assessment's proven minimal change (`unreadable-content-warning` →
`clean-open`). Equivalently, repacking the package with any correct ZIP writer produces a clean
envelope. A spec-conformant alternative — keep the bytes and declare them by writing comment
length 5 — is legal per APPNOTE §4.3.16/§4.4.25 but is **Word-unproven**; prefer the proven
deletion.

## Where the fixes belong and why

- **repair (the one fix this issue prescribes).** This is a broken source document: the producer
  appended five text-ish bytes after the ZIP stream. Nothing in the XML parts needs touching —
  the analyst's per-part XSD runs found only known non-defects (w14:paraId/textId and
  mc:Ignorable noise), and the repo's own schema dispatch passed every part.
- **Not docx-validate.** The check exists and reproduces the defect verbatim on the recorded
  shape (runs 3–5 above). Its lenient omission is documented, intended behaviour
  (`validate.ts:853-857`), not a miss.
- **Not a redline-generator bug.** Lane `analyst-source`/`source-doc` is a source document, not
  a docxodus/superdoc-redlines/jubarte redline. If the corpus is ever regenerated, the producer
  that writes a trailer after the ZIP stream deserves a look — but the actionable item here is
  the document repair.
- **Lane hygiene (out of band, same as sibling issue 0888):** the specimen, its unpacked copy
  and `really_broken/manifest.csv` are all missing, which cost this session its live Word
  observation. Restore the corpus artifacts so verification can be re-established.

## Verification

- **Established Word cause:** the proven assessment row (`assessments.jsonl:13`) — original bytes
  → `unreadable-content-warning`; original minus the five trailing bytes → `clean-open`; every
  preceding byte unchanged. Per this repo's probe protocol a `proven` row already rests on two
  agreeing Word runs with logs.
- **Portable check (run this session):** on the reconstruction, `validate.ts` default profile
  exits 1 with `zip-bytes-after-eocd` naming 5 undeclared bytes; the same bytes deleted →
  `All validations PASSED!`; `--profile lenient` passes in both cases by design.
- **On the restored specimen (blocked until the file is back):** re-run
  `pnpm exec tsx src/scripts/office/validate.ts <restored>.docx` — expect the `zip-bytes-after-eocd`
  error, count 5 — then the two-run Word A/B: original (expect the unreadable-content dialog)
  vs. the 16,139-byte truncation (expect a clean open). Any disagreement with the recorded
  outcomes reopens this issue.
