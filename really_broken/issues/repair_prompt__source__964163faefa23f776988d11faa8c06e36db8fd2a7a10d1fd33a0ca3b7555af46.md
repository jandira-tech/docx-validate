# [source 964163fa] failed: five undeclared bytes after the ZIP EOCD record

## Summary

`fixtures/repair_prompt/964163faefa23f776988d11faa8c06e36db8fd2a7a10d1fd33a0ca3b7555af46.docx`
is a **source document** (a Word file downloaded from `docxcorp.us`, per the corpus manifest row
129) in the `repair_prompt` lane: Microsoft Word offers to recover it. Its 11 XML parts are clean
— once the container defect below is removed, the file passes every docx-validate check under the
default (word-valid) profile. The defect is entirely in the **ZIP envelope**: five bytes
(`00 0D 0A 0D 0A`) follow the end-of-central-directory record, whose comment length is 0.

Provenance note: the live Word check for this ask reported `failed — file does not exist`, because
it looked under `really_broken/fixtures/` **inside the docx-validate repo**, where the corpus is not
checked out. I located the specimen this session at
`/Users/arthrod/temp/T/neurotic_docx_bench/maybe_future_corpus/really_broken/fixtures/repair_prompt/964163faefa23f776988d11faa8c06e36db8fd2a7a10d1fd33a0ca3b7555af46.docx`
(outside the repo) and verified `shasum -a 256` equals the file name, so all byte-level
re-verification below ran against the exact corpus bytes. No Word open happened this session; the
Word authority is the recorded two-run experiment (see Word behaviour).

## Word behaviour

- **Live result (this ask):** never opened — outcome `failed`, reason `really_broken/fixtures/`
  does not exist in the repo and a repo-wide `find` returned nothing (re-confirmed by me; the
  specimen lives in the external corpus above). No new Word observation was made.
- **Manifest row** (`…/neurotic_docx_bench/maybe_future_corpus/really_broken/manifest.csv:129`,
  read this session): category `repair_prompt`, source
  `https://docxcorp.us/documents/964163fa….docx`, `d1_failed_in=corrupt`, local_word_error
  `Word could not load it in 2 passes: document loaded empty (Word could not read it)`.
- **Prior proven assessment — the established Word cause**
  (`docs/really-broken/assessments.jsonl:8`, mirrored at `docs/really-broken/inspection-log.jsonl:8`,
  both read this session; `status: proven`):
  defect *"Five undeclared bytes follow the ZIP end-of-central-directory record. The record's
  comment length is 0. The bytes are 00 0D 0A 0D 0A."*; minimal change *"Delete those five bytes so
  the end-of-central-directory record is the last record in the file. Every preceding byte stays
  the same."*; Word verdicts: `original: unreadable-content-warning` (the repair prompt) →
  `after_this_change: clean-open`. The assessment names no check code; this repo's check for the
  shape is `zip-bytes-after-eocd`.

## Violation 1 — `zip-bytes-after-eocd`: 5 undeclared bytes after the EOCD record

**Offending bytes** (ZIP envelope — no package part holds this defect). Decoded this session with a
`struct` unpack over the fetched file:

```text
file size 14870 (0x3a16)
EOCD @ 0x39fb (14843):  50 4B 05 06 | disk 0000 | cd-disk 0000
  entries-this-disk 000B (11) | entries-total 000B (11)
  cd-size C1 02 00 00 (705) | cd-offset 3A 37 00 00 (14138)
  → 14138 + 705 = 14843, abuts the EOCD, so this is the live archive's record
  comment-length 00 00 (0) → record ends at 14843 + 22 + 0 = 14865 (0x3a11)
trailing bytes 0x3a11–0x3a15: 00 0D 0A 0D 0A   (NUL CR LF CR LF)
```

`xxd | tail -4` of the same file ends `…504b 0506 0000 0000 000b 000b 00c1 0200 003a 3700 0000
0000 0d0a 0d0a` — identical to the analyst's snippet. `unzip -l` lists 11 entries, matching
entries-total 11.

**Rule.** No bundled XSD governs the ZIP container (`src/scripts/office/schemas/` wml/opc XSDs
cover package parts, not the envelope). The rule is PKWARE APPNOTE.TXT, verified this session via
WebFetch of `https://pkware.cachefly.net/webdocs/casestudies/APPNOTE.TXT`:

- §4.3.16 "End of central directory record" field table, last two rows:
  `".ZIP file comment length        2 bytes"` / `".ZIP file comment       (variable size)"`
  — the record's declared extent is 22 fixed bytes plus the comment; here 22 + 0 = ends at 14865.
- §4.3.1: `"A ZIP file MUST contain an 'end of central directory record'."` and
  `"A ZIP file MUST have only one 'end of central directory record'."`

The APPNOTE copy contains no explicit "nothing may follow the EOCD" sentence (confirmed by my
fetch), so the citation rests on the record's declared extent — the file runs 5 bytes past the
archive's declared end — anchored by the proven Word experiment above (repair prompt → clean open
when exactly those 5 bytes are removed). That is the strongest available evidence chain, and it is
recorded, two-run Word evidence per `docs/really-broken/README.md`.

**This repo's check** (exists, accurate, fires — no validator gap):

- `src/scripts/office/checks/package.ts:16-33` — `zip-bytes-after-eocd` check, message at :29.
- `src/scripts/office/zip-bytes-after-eocd.ts:152-180` — `countZipBytesAfterEocd`.
- `src/scripts/office/validators/word-error-explanations.ts:129-132` — user explanation.
- Ran this session: `pnpm exec tsx src/scripts/office/validate.ts <specimen>` (default profile,
  which is word-valid for `.docx`) → exit 1, single ERROR `ZIP end of central directory is
  followed by 5 undeclared byte(s).`; `--profile word-valid` → identical; `--profile lenient` →
  exit 0 `All validations PASSED!` (envelope checks do not run under lenient — by design, AGENTS.md
  note 18). The analyst report's claim that the *default* profile passes is a mislabel: the CLI
  default for `.docx` **is** word-valid and does report the defect; only lenient misses it.

**Proposed fix** — delete the five bytes at file offsets 0x3a11–0x3a15 (14865–14869) so the EOCD
record is the last record in the file; every preceding byte stays the same. This is byte-identical
to the proven minimal change. **Fix target: `repair`** (broken source document; the producing
side is an external `docxcorp.us` download, not a redline generator in this repo).

## Where the fixes belong and why

- **`repair`** — the one and only defect. Truncate the file to its first 14865 bytes. No XML part
  changes; the 11 entries and their bytes are untouched.
- **Not `docx-validate`** — the check exists and fires under the profile Word-valid models, with a
  message that matches the recorded Word behaviour ("Word offers to recover the document").
- **Not a generator bug** — lane label is `source`; the document was fetched, not produced by
  docxodus / superdoc-redlines / jubarte. (Context, unproven: the bytes `00 0D 0A 0D 0A` and the
  validator's two info `rel-ids-sequential` findings — "likely regenerated by the repairer" —
  suggest the file already passed through some server-side repair/normalization stream that padded
  the tail; the origin is not established and needs no fixing to repair this file.)

## Verification

Minimal-change experiment, portably re-run this session (Word side already proven by the recorded
two-run assessment):

1. `python3` — wrote `/tmp/eocd-repair/fixed.docx` = first 14865 bytes of the specimen (dropped
   `00 0d 0a 0d 0a`), nothing else changed.
2. `pnpm exec tsx src/scripts/office/validate.ts /tmp/eocd-repair/fixed.docx` → **exit 0,
   `All validations PASSED!`** (was exit 1 with the single `zip-bytes-after-eocd` ERROR).
3. Recorded Word A/B (`assessments.jsonl:8`): original `unreadable-content-warning` → repaired
   `clean-open`. Not re-opened in Word this session (live check could not resolve the path); to
   re-confirm with real Word, point `--verify-word` at the corpus copy or restore
   `really_broken/fixtures/` in the repo.
