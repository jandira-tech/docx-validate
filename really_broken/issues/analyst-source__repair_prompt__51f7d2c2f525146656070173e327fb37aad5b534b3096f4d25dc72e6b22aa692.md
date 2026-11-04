# [source 51f7d2c2] failed: ZIP EOCD comment length 3328 overruns end of file

## Summary

`fixtures/repair_prompt/51f7d2c2f525146656070173e327fb37aad5b534b3096f4d25dc72e6b22aa692.docx`
(lane category `repair_prompt`, producing tool **source-doc** — a downloaded document, not a
redline generator) is a minimal 11-part WordprocessingML package: `[Content_Types].xml`,
`_rels/.rels`, `word/_rels/document.xml.rels`, `word/document.xml`, `word/theme/theme1.xml`,
`word/styles.xml`, `word/settings.xml`, `word/fontTable.xml`, `word/webSettings.xml`,
`docProps/app.xml`, `docProps/core.xml` (full listing via `unzip -l` on the byte-patched copy).

**Every XML part is clean; the only defect is in the ZIP envelope.** The end-of-central-directory
record declares a 3328-byte archive comment where 3 bytes remain, so the comment overruns
end-of-file by 3325 bytes. Word refuses the file.

**Live Word outcome vs manifest:** the live check this session returned `failed` without ever
opening Word — it looked for `really_broken/fixtures/…` inside this repo, which does not exist
(the repo's `really_broken/` holds only `issues/`, `tool_redlines/`, `unpacked/`, and no
`manifest.csv`). The specimen itself survives in the source corpus at
`/Users/arthrod/temp/T/neurotic_docx_bench/maybe_future_corpus/really_broken/fixtures/repair_prompt/51f7d2c2….docx`;
I re-verified its bytes there — 18562 bytes, `shasum -a 256` = `51f7d2c2f52514…aa692`, exactly the
ID, so the analysis below is pinned to those bytes. The repo-local copy and manifest row the ask
names are missing and should be restored or the ask's paths corrected.

## Word behaviour

- **Live result (this ask):** `failed` — Word never opened; `really_broken/fixtures/` does not
  exist in this repo and a repo-wide `find -name "*51f7d2c2*"` returns nothing. Not run again
  here; no macOS Word observation was made this session.
- **Manifest row** (source corpus `really_broken/manifest.csv:123`, corpus copy): category
  `repair_prompt`, sha256 `51f7d2c2…aa692`, origin
  `https://docxcorp.us/documents/51f7d2c2….docx`, `d1_failed_in: corrupt`, local_word_error
  `Word could not load it in 2 passes: document loaded empty (Word could not read it)`.
- **Prior assessment — PROVEN** (`docs/really-broken/assessments.jsonl:2`, mirrored in
  `inspection-log.jsonl:2`): original Word outcome `unreadable-content-warning`; after setting
  the comment-length field from 3328 to 3 — "No other byte changes" — Word outcome `clean-open`.
  That is the established Word cause for this file, and this session's validator experiment
  (below) reproduces the same minimal change at the package level.

## The violation — `zip-eocd-comment-overrun` (1 site, ZIP envelope)

**Offending bytes** (my own decode this session, Python `struct.unpack('<HHHHIIH', …)` over the
18562-byte file; `xxd` tail agrees):

```
00004870: 000b 000b 00c1 0200 00a8 4500 0000 0d0a
00004880: 0d0a
```

- EOCD signature `PK\x05\x06` at file offset **18537**; fixed record is 22 bytes.
- 11 entries (both count fields `000b`), central-directory size 705 (`00c1 0200`) at offset
  17832 (`a8 45 00 00`) — the directory reads fine.
- Comment-length field at offset **18557**, bytes `00 0d`, little-endian = **0x0D00 = 3328**.
- Bytes actually present after the record: **3** — `0a 0d 0a` (offsets 18559–18561), then EOF.
- Declared 3328 > remaining 3 → overrun of 3325 bytes. (Curiously `0x000D` = 13, so the field
  may be a byte-swapped intended length; either way nothing near 3328 bytes exists.)

**Rule.** No bundled XSD governs the ZIP container (OPC fallback). The rule is operational, in
`src/scripts/office/zip-bytes-after-eocd.ts`:

- Lines 69–71 of `zipEocdCommentOverrun` (function at 63–75):
  `const declared = readU16(buf, offset + 20); const remaining = buf.length - (offset + EOCD_FIXED_LENGTH); if (declared > remaining) return { declared, remaining };`
- Doc comment, lines 32–34: "`zipEocdCommentOverrun` reports a comment length that claims more
  bytes / than remain in the file. Setting the length to the number of bytes actually / present
  is not that defect." (The analyst's snippet cited these as lines 26–28; the quoted text is
  real, the line offset was slightly off.)
- Registered as check id `zip-eocd-comment-overrun`, severity error, formats docx/pptx/xlsx,
  `archiveOnly: true` — `src/scripts/office/checks/package.ts:54-71`, message at line 67.

**Reproduced this session** (`pnpm exec tsx src/scripts/office/validate.ts` on a copy):

- Default profile: `ERROR: ZIP end-of-central-directory comment length is 3328, but only 3
  byte(s) remain in the file. Word offers to recover the document. …` plus
  `ERROR: Cannot open … as an OPC package … End of data reached (data length = 18562, asked
  index = 21887). Corrupted zip ?`. Note 21887 = 18559 + 3328: the package-open failure is the
  same overrun read past EOF, not a second defect.
- `--profile word-valid`: the same two errors.

**Proposed fix (repair target).** Set the EOCD comment-length field — file offset 18557, 2 bytes
little-endian — from `00 0d` (3328) to `03 00` (3), leaving the 3 trailing bytes `0a 0d 0a` in
place as the archive comment. No other byte changes. (Equivalent alternative: truncate the 3
trailing bytes and set the field to `00 00`; the proven experiment used the 2-byte patch.)

**Fix target:** `repair` — a broken source document. docx-validate already detects and precisely
describes this defect under every profile; there is no validator gap and no generator to change.

## Where the fixes belong and why

- **repair (this document):** the corruption is in the downloaded file's ZIP envelope, upstream
  of any XML part. Whoever curates the corpus should apply the 2-byte patch above (or re-fetch
  and verify the fetch is not what truncated the file — a 3328-byte claim beside a 3-byte tail
  smells like a truncated download or a botched CRLF append by the serving host).
- **docx-validate:** nothing to fix — `zip-eocd-comment-overrun` (checks/package.ts:54) fired
  under both profiles with the exact declared/remaining counts, and its message already
  prescribes the repair.
- **Producing tool (source-doc):** n/a; the defect is in the source document's bytes, not in
  any redline generator's XML output.
- **Lane plumbing (out of code scope):** the repo working tree lacks `really_broken/fixtures/`
  and `really_broken/manifest.csv`, which made the live Word check fail on a missing path. The
  specimen exists only in the external corpus; the lane should either sync the fixture in or
  point the Word verifier at the corpus path.

## Verification

All run this session on `/tmp/fix51f7.TBBz/` copies (sha256 verified = `51f7d2c2…aa692`):

1. **EOCD decode** — Python `struct` unpack of the last EOCD: entries 11/11, CD size 705 @
   17832, comment length 3328, 3 bytes after the record (`0a 0d 0a`). Matches the analyst.
2. **Validator, unpatched** — `pnpm exec tsx src/scripts/office/validate.ts <file>` (default)
   and `… --profile word-valid <file>`: both emit the `zip-eocd-comment-overrun` error quoted
   above; the default profile also emits `package-open-failed` with `asked index = 21887`
   (= 18559 + 3328).
3. **Minimal-change experiment (the proven assessment's change, reproduced)** — copy, patch
   bytes 18557–18558 `00 0d` → `03 00` (verified the diff is exactly those 2 bytes), then
   `unzip -l` lists all 11 entries with no truncation caution, and
   `pnpm exec tsx src/scripts/office/validate.ts --profile word-valid <patched>` prints
   **`All validations PASSED!`** — confirming no XML-level defect exists in any of the 11 parts
   (each also individually well-formed under `xmllint --noout`, exit 0).
4. **Word (not run this session):** the live verifier could not find the repo-local path; the
   Word evidence for this file is the proven assessment's `unreadable-content-warning` →
   `clean-open` pair (assessments.jsonl:2). A fresh two-run Word observation of the patched
   bytes would re-confirm `clean-open`, and of the original would re-confirm the recovery
   offer, once the fixture is reachable again.
