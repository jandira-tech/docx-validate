# [source 8226d0ab] failed: 7 undeclared bytes after the ZIP EOCD record

## Summary

`fixtures/repair_prompt/8226d0ab…3929bd.docx` is a source document (repair_prompt lane; manifest row
`set=fixture`, `tool` empty — a downloaded artifact, `artifact=https://docxcorp.us/documents/8226d0ab….docx`,
not a redline product). The package contents are healthy: 12 central-directory entries / 11 parts
(`[Content_Types].xml`, both `.rels`, `docProps/{app,core}.xml`,
`word/{document,fontTable,numbering,settings,styles,webSettings}.xml`, `word/theme/theme1.xml`), and
`word/document.xml` validates against the bundled `wml.xsd`.

The only defect is in the ZIP envelope: **7 undeclared bytes follow the end-of-central-directory
record, whose comment-length field is 0**, so the archive's terminating record is not the last record
in the file. Word refuses the file over those bytes.

Path note: the ask's `really_broken/…` paths are not inside this repo. The corpus lives at
`/Users/arthrod/temp/T/neurotic_docx_bench/maybe_future_corpus/really_broken/` — packed fixture,
`unpacked/source__repair_prompt__8226d0ab…`, and `manifest.csv` were all verified there this session.
That is why the live Word probe failed: it looked for `really_broken/fixtures/` under the
docx-validate repo, where it does not exist, so **no live Word observation exists for this file this
session**; the Word evidence below is the lane's recorded, proven assessment.

## Word behaviour

- **Live result (this ask):** never opened — outcome `failed`, "the file does not exist". The file
  does exist at the corpus path above (15,603 bytes); only the probed path was wrong. I did not
  re-run Word against the correct path in this session.
- **Manifest row** (`really_broken/manifest.csv` at the corpus root, grepped for the fixture path):
  `category=repair_prompt`, `tool` empty, `d1_failed_in=corrupt`, `local_word_error="Word could not
  load it in 2 passes: document loaded empty (Word could not read it)"`.
- **Proven assessment** (`docs/really-broken/assessments.jsonl:4`, status `proven`; mirrored at
  `docs/really-broken/inspection-log.jsonl:4`): "Seven undeclared bytes follow the ZIP
  end-of-central-directory record. The record's comment length is 0. The bytes are 0A 0A 0A 0D 0A 0D
  0A." Minimal-change experiment: **deleting those seven bytes changed Word from
  `unreadable-content-warning` to `clean-open`.** Per `docs/really-broken/README.md`, all 144 recorded
  assessments are proven two-run Word investigations.

## Each violation

### 1. `zip-bytes-after-eocd` — 7 undeclared bytes after the EOCD record (1 site)

**Offending bytes** (re-measured this session by parsing the packed file; the EOCD at 15574 is the
only one in the file, so this is not the concatenated-archive case):

```text
file size 15603
EOCD offset 15574:  50 4B 05 06 | 00 00 | 00 00 | 0C 00 | 0C 00 | 01 03 00 00 | D5 39 00 00 | 00 00
                    signature    disk    cd-disk  12 entr  12 entr  cd size 769   cd off 14805   cmt len 0
record ends at 15574 + 22 + 0 = 15596   (cdOff 14805 + cdSize 769 = 15574 — the record abuts the
central directory, i.e. it is the archive's structural terminating record)
trailing bytes 15596..15602 (7 bytes):  0A 0A 0A 0D 0A 0D 0A   (LF LF LF CR LF CR LF)
```

**Rule.** No bundled XSD covers the ZIP container: `src/scripts/office/schemas/ecma/fourth-edition/`
holds exactly `opc-contentTypes.xsd`, `opc-coreProperties.xsd`, `opc-digSig.xsd`,
`opc-relationships.xsd` (content types, core properties, signatures, relationships), and no schema
under `src/scripts/office/schemas/` mentions EOCD/APPNOTE (grepped). The rule is the repo's own
byte-level check over ISO/IEC 29500-2's physical-package model — the package is a ZIP archive whose
end-of-central-directory record, declared comment included, terminates the archive. Canonical rule
text, `src/scripts/office/validators/word-error-explanations.ts:129-130`:

> "Bytes follow the ZIP end-of-central-directory record beyond the comment length that record
> declares. Word offers to recover the document. Delete those bytes. A ZIP comment whose length is
> declared in the record is not this defect. A second copy of the archive appended after the record
> is this defect even when that copy ends with its own end-of-central-directory record."

(The analyst note cited this text as `word-error-explanations.ts:96-97`; that range actually holds
`word-shared-story-part` / `word-textbox-deleted-text`. Citation slip only — the quoted text is
verbatim at lines 129-130, which I read.)

Check wiring: `src/scripts/office/checks/package.ts:16-33` (`zip-bytes-after-eocd`, group `package`,
error severity) calling `countZipBytesAfterEocd` (`src/scripts/office/zip-bytes-after-eocd.ts:151-180`).

**Proposed fix — repair the document:** truncate the file to 15,596 bytes, deleting the trailing
`0A 0A 0A 0D 0A 0D 0A` so the EOCD record is the last record in the file; every preceding byte stays
the same. (Equivalent alternative: set the record's comment-length field to 7 so the bytes become a
declared ZIP comment — structurally legal, but unprobed in Word; the proven experiment is deletion.)

**Fix target:** `repair` (broken source document; no producing tool to change).

## Where the fixes belong and why

- **repair** — the document's trailing bytes are the only thing wrong. The proven assessment's
  deletion experiment is the repair prescription.
- **Not docx-validate** — the validator already catches it, correctly, at the profile that carries
  envelope checks. Re-verified this session on the real file:
  - default profile (`word-valid`, the `.docx` default per `validate.ts --help`): `ERROR: ZIP end of
    central directory is followed by 7 undeclared byte(s). …` — exit 1; in `--json` it is the run's
    only error-severity issue (the rest are `paraid-missing-element` and `style-default-missing`
    warnings and `rel-ids-sequential` info).
  - `--profile lenient`: `All validations PASSED!` — exit 0. Documented policy, not a gap: lenient
    runs neither the Word rules nor the ZIP-envelope checks (AGENTS.md note 18).
  - `countZipBytesAfterEocd(fixtureBytes)` called directly returns `7`.
- No redline generator is involved (manifest `tool` empty), so there is no producer bug to file.

One correction to the analyst note for the record: it says the "default profile printed 'All
validations PASSED!'" and that word-valid "is the one that carries this container check". At this
revision the default for `.docx` **is** `word-valid` and it reports the error; the passing
observation reproduces only as `--profile lenient`. The violation itself is unaffected.

## Verification

- Byte measurement (this session, Node parse of the packed file): EOCD 15574, comment length 0,
  record end 15596, file size 15603, 7 trailing bytes `0A 0A 0A 0D 0A 0D 0A`, single EOCD in file.
- `pnpm exec tsx src/scripts/office/validate.ts <fixture>` → 1 error (`zip-bytes-after-eocd`),
  exit 1; `--json` shows `valid: false`, `profile: word-valid`, that one error issue.
- `pnpm exec tsx src/scripts/office/validate.ts <fixture> --profile lenient` → `All validations
  PASSED!`, exit 0.
- `xmllint --noout --schema src/scripts/office/schemas/ISO-IEC29500-4_2016/wml.xsd
  <unpacked>/word/document.xml` → `validates`; `[Content_Types].xml` and both `.rels` parts parse
  well-formed — no XML-level defect accompanies the envelope one.
- Unit suite for the check: `pnpm exec vitest run tests/zip-bytes-after-eocd.test.ts` → 13/13 pass.
- Word proof (recorded, not re-run this session): `docs/really-broken/assessments.jsonl:4` — delete
  the 7 bytes → `clean-open`; original → `unreadable-content-warning`. A repairer confirms the
  truncated copy with two Word opens per the repo's two-run rule.
