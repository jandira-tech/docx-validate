# [jubarte 1360] failed: no defect — crashes_word label is a harness artifact

## Summary

`tool_redlines/crashes_word/jubarte__01a11a49_1360.docx` is a jubarte-produced
redline (a 2.26 MB `word/document.xml`, 60,060 lines, carrying 77 `w:ins` and
3,309 `w:del` marks) over a contracts-style source document. **No violation
reproduces against the bundled XSDs, and the recorded proven Word evidence says
the file opens clean.** The `crashes_word` category rests on a lane-harness
misreading of an Apple Event transport error as a document crash.

The ask's in-repo paths were stale: `really_broken/tool_redlines/` contains only
`repair_prompt/`, `really_broken/manifest.csv` and any `*1360*` file are absent
repo-wide (verified with `find`), and the live Word probe therefore reported
`failed: The file does not exist`. The actual specimen lives in the sibling
corpus checkout at
`/Users/arthrod/temp/T/neurotic_docx_bench/maybe_future_corpus/really_broken/tool_redlines/crashes_word/jubarte__01a11a49_1360.docx`;
its sha256 `e2063d143f…b046` matches the manifest row exactly, so every check
below ran on the manifest's bytes.

## Word behaviour

- **Live result (this ask):** `outcome: failed` — "The file does not exist". I
  verified the observation myself: no `crashes_word/` directory, no `*1360*`
  file anywhere in the repo, no `really_broken/manifest.csv`. Word was never
  opened; this is not evidence about the document.
- **Manifest row** (`maybe_future_corpus/really_broken/manifest.csv:98`, header
  at line 1): `set=tool_redline, category=crashes_word, tool=jubarte,
  d1_failed_in=word:corrupt, local_word_error="Word crashed on it in 2 passes:
  Microsoft Word got an error: Connection is invalid."` — read directly, quoted
  verbatim.
- **Prior assessment (proven):** `docs/really-broken/assessments.jsonl:21`,
  mirrored in `docs/really-broken/inspection-log.jsonl:21` —
  `status: proven, defects: [], defect: "None. Microsoft Word opens the
  original file with no warning, error, or offer to repair.", change: none,
  word.original: clean-open`. Per `docs/really-broken/README.md:5`, a clean-open
  assessment deliberately carries an empty `defects` array: no demonstrated
  defect, not missing evidence.
- **Cause of the manifest's "crash":** the lane harness's own classifier,
  `neurotic_docx_bench/live-bench/harvest/harvest.py:191-192`:

  ```python
  if "Connection is invalid" in err or "Word crashed" in err or "error reporter appeared" in err:
      return "crashes_word"
  ```

  "Microsoft Word got an error: Connection is invalid." is the AppleScript/
  Apple Event error the driver sees when its event connection to Word drops —
  an automation-transport failure, not a document-crash symptom. docx-validate's
  own verifier already encodes the correct semantics: a vanished Word process is
  `crash` only when macOS wrote a Word crash report after the open, otherwise
  `unavailable` (`src/lib/word-verifier.ts:365-378`, #305), and a verifier that
  throws yields `unavailable`, never a rejection (#333).

## Violations

**None.** Every candidate finding re-verified this session reduces to one of
four known-benign classes. All xmllint commands ran against the bundled XSDs in
`src/scripts/office/schemas/` over the unpacked package at
`…/unpacked/tool_redlines__crashes_word__jubarte__01a11a49_1360/`.

1. **OPC parts are fully valid.** `xmllint --noout --schema
   ecma/fourth-edition/opc-contentTypes.xsd '[Content_Types].xml'` → *validates*;
   both `_rels/.rels` and `word/_rels/document.xml.rels` against
   `opc-relationships.xsd` → *validates*; `docProps/core.xml` against
   `opc-coreProperties.xsd` → *validates* (the bundled Dublin Core copies
   resolve; no remote-import issue in this run).
2. **`word/document.xml` (and the note/header/footer parts):** 4,210 xmllint
   findings, every one an undeclared-attribute notice for
   `w14:paraId`/`w14:textId` (2,099 each) or `wp14:anchorId`/`wp14:editId` (12)
   — e.g. `document.xml:9993: element p: … attribute
   '{…/word/2010/wordml}paraId': The attribute … is not allowed.` These are
   declared Microsoft extensions, not violations of the base ISO schema:
   `microsoft/wml-2010.xsd` declares `paraId`/`textId`, and the ISO
   `wml.xsd` leaves them to the extension schema set. This is exactly the
   Microsoft-extension noise class the bundled `microsoft/wml-*.xsd` schemas
   exist for.
3. **`word/settings.xml:94-96`:** `<w14:docId w14:val="605D9CEC"/>`,
   `<w14:defaultImageDpi>`, `<w15:docId …/>` — "element docId: This element is
   not expected" against the base schema. Same class: `docId` is declared in
   `microsoft/wml-2010.xsd`/`wml-2012.xsd`; Word itself emits this tail.
4. **`word/header1.xml:24` `mc:AlternateContent`:** the canonical Word shape
   wrapper — `<w:r><mc:AlternateContent><mc:Choice Requires="wps"><w:drawing>
   <wp:anchor …>` — flagged only because the ISO `wml.xsd` declares no
   `AlternateContent` reference (grep count: 0); `mce/mc.xsd:37-53` declares
   `AlternateContent`/`Choice`/`Fallback`, and Word processes the wrapper
   natively per MCE. Root `mc:Ignorable` attributes on every part are the same
   MCE class.

The repo's own validator — which applies the proper schema mapping and MCE
preprocessing — passes the file: `pnpm exec tsx src/scripts/office/validate.ts
<specimen>` printed exactly `All validations PASSED!` (zero issues, so no
check codes to collect). ZIP envelope: `zipinfo -1` lists 20 entries with no
duplicate names.

**Status: disputed.** No violation reproduces; the `crashes_word` premise
itself does not reproduce — the only recorded Word behaviour for these bytes is
a proven clean open.

## Where the fixes belong and why

- **Not docx-validate.** The validator passes the file, and that is the correct
  verdict: no defect exists to miss, and no false positive is reported. Its
  `#305` crash-vs-`unavailable` semantics (`src/lib/word-verifier.ts:365-378`)
  are the model the lane harness failed to follow.
- **Not jubarte.** The redline's XML is schema-legal (Microsoft extensions plus
  standard MCE), and Word opens it clean. There is no generator change to make.
- **Not a repair prescription.** The source document is not broken.
- **The fix belongs to the really-broken lane harness**
  (`neurotic_docx_bench/live-bench/harvest/harvest.py:186-193`): stop mapping
  the Apple Event "Connection is invalid" transport error to `crashes_word`,
  and re-run or re-classify specimen 1360. Concretely, treat that string as
  `unconfirmed`/`unavailable` unless a Microsoft Word crash report was written
  after the open (the `#305` rule).

## Verification

- Re-confirmed the manifest row and byte identity: `sed -n '98p' <manifest.csv>`
  plus `shasum -a 256` → matches `e2063d143f…b046`.
- Re-ran the full validator on the manifest's bytes → `All validations PASSED!`.
- Re-ran the xmllint sweep over every XML part (commands and tallies above);
  only the four benign classes appear, with zero residual findings.
- The recorded experiment stands as the established Word cause: the proven
  assessment (line 21 of both JSONL files) opened the original in Microsoft
  Word with no warning, error, or repair offer, and found no isolated edit that
  changes any verdict (`change: none`). No further minimal-change experiment is
  needed to prove a cause — there is no defect to prove.
- Follow-up for the lane: after `harvest.py`'s `category()` stops treating
  "Connection is invalid" as a crash, re-probe the file twice with the two-run
  protocol; the expected observation is `clean-open`, at which point it should
  leave the crashes_word lane.
