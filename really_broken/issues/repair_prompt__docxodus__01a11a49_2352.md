# [docxodus 2352] repair-prompt: undeclared `wps` prefix in `mc:Choice/@Requires`

## Summary

`tool_redlines/repair_prompt/docxodus__01a11a49_2352.docx` (sha256
`8a40bbafdfbadc7593974563c4b28ab7a39bfa415fe67628174b6150b0f6e565`) is a
docxodus/Open-Xml-PowerTools redline of a Czech document: a tracked-inserted run
holds an `mc:AlternateContent` whose single `mc:Choice` (prefix `ve`) names
`Requires="wps"` — a prefix declared nowhere in the part — plus a plain second
`w:drawing`, over regenerated styles/numbering. Live Word (this session) showed
the repair prompt ("Word found unreadable content … recover the contents?"); No
was followed by an open error and no document window; the manifest row agrees
(`word:corrupt`). The dominant, reproduced defect is the unbound `wps` prefix.
The second reported violation (whitespace inside drawing scalars) **does not
exist in the .docx** — it is manufactured by this repo's own `unpack`
pretty-printer — so this issue disputes it as a file defect and files the
round-trip corruption as the docx-validate finding instead.

## Word behaviour

- Live result (given): repair-prompt dialog, then "Word experienced an error
  trying to open the file"; no crash report; `outcome: "repair-prompt"`.
- Manifest row (corpus `manifest.csv`; `really_broken/manifest.csv` is absent
  from the working tree): `repair_prompt`, d1_failed_in `word:corrupt`, "Word
  could not load it in 2 passes: document loaded empty (Word could not read it)
  | harvest 20261013_131626" — same verdict as live Word.
- Prior assessment: none. `grep docxodus__01a11a49_2352` over
  `docs/really-broken/assessments.jsonl` and `inspection-log.jsonl`: 0 matches
  in both (run this session). The cause is **not yet proven** for this file;
  the evidence is the two reproduced ERROR codes plus hash-bound fixture
  observations below.

## Violation 1 — `Requires="wps"` names an undeclared prefix (reproduced)

Offending XML — `word/document.xml` root (unpacked copy line 2) and lines 25–26:

```xml
<w:document xmlns:ve="…markup-compatibility/2006" xmlns:o="…office:office"
  xmlns:r="…relationships" xmlns:m="…math" xmlns:v="…vml" xmlns:wp="…wordprocessingDrawing"
  xmlns:w10="…office:word" xmlns:w="…wordprocessingml/2006/main" xmlns:wne="…word/2006/wordml">
…
<ve:AlternateContent>
  <ve:Choice Requires="wps">
```

Verified in the packed bytes: `unzip -p … word/document.xml | grep -c
xmlns:wps` → `0`. The whole part declares only `ve, o, r, m, v, wp, w10, w, wne`
on the root and `a, a14, pic, p17` inline on descendants — never `wps`. The
Choice payload itself uses the intended namespace as a *default* namespace
(`<wsp xmlns="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">`,
line 47), so the shape was meant to carry a `wps` binding.

Schema rule — `src/scripts/office/schemas/mce/mc.xsd:34-53` (read this session):

```xsd
<!-- An AlternateContent element shall contain one or more Choice child elements, optionally followed by a
Fallback child element. … -->
<xsd:attribute name="Requires" type="xsd:string" use="required" />
```

Honest scope: `Requires` is `xsd:string`, so the undeclared token is **not** a
lexical XSD failure; the binding requirement is the MCE processing rule
(ISO/IEC 29500-3: each `Requires` token is a namespace prefix that must be
bound on the Choice or an ancestor), which this repo enforces —
`word-choice-requires.ts:21-38` walks ancestors for the `xmlns:` binding;
`markup-compatibility.ts:348` marks the branch malformed when
`namespaceUri(branch, prefix)` resolves nothing. One defect therefore draws two
ERROR codes: `word-choice-requires-undeclared` and `mce-malformed` ("…
requires Choices with declared Requires prefixes…"). Declaring the prefix
clears both.

Word evidence: hash-bound observations in `tests/fixtures-all.manifest.json` —
`broken/package.word-choice-requires-undeclared.docx` →
`unreadable-content-warning` (Word 16.115 recovery dialog) and
`working/package.word-choice-requires-declared.docx` → `clean-open` (both
2026-10-26). Same shape as this file's live repair prompt.

**Proposed fix (fix target: docxodus):** emit
`xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"`
on the `w:document` root (or the `ve:Choice`). The generator must preserve the
declarations its own Choice payloads require; as shipped, the Choice can never
be selected and consumers fall back to VML — or, as observed, refuse the file.

## Violation 2 — whitespace in drawing scalars (DISPUTED as a file defect)

The report cites six whitespace-bearing scalars (`wp:posOffset` at unpacked
lines 31/35/110/114, `pctWidth`/`pctHeight` at 81/85) and calls the CLI's
silence a docx-validate gap. Re-verification disproves the file defect:

- Packed bytes are **clean**: `unzip -p … word/document.xml | grep -o
  '<wp:posOffset>[^<]*</wp:posOffset>|<pctWidth>[^<]*</pctWidth>|<pctHeight>[^<]*</pctHeight>'`
  returns `1065530, -90170, 0, 0, 14606, -107718` — every value tight against
  its tags.
- The whitespace exists only in the unpacked copy, byte-identical (`diff -q`:
  IDENTICAL) to `unpack --merge-runs false --simplify-redlines false` run this
  session. `prettyXml` (`src/lib/xml-helpers.ts:331-371`) injects it: the text
  branch (362-368) appends the value to the opening tag's line, then the
  closing-tag branch (344-348) unconditionally starts a new line — turning clean
  input into `<wp:posOffset>14606\n…</wp:posOffset>` (reproduced with a direct
  `prettyXml` call on a minimal document).
- The "silenced check" claim is wrong in direction: nothing is silenced.
  `--checks word-drawing-scalar-whitespace` (run this session, 2 checks, exit 0)
  reports nothing because there is nothing to report. Introspection through
  `ValidationContext.fromArchive` + `effectiveXml` shows the effective view holds
  exactly the two Choice-external `posOffset` text nodes, `"14606"` and
  `"-107718"`, both trim-equal. The analyst's contrary reading came from the
  pretty-printed disk file, not the archive.

The real docx-validate defect (verified this session) is the opposite:
**unpack→pack manufactures the defect.** Packing the plain-unpacked tree with
`pack --validate false` yields a .docx whose scalars are `'1065530\n…'` etc.
(python `repr` of the round-tripped part); validating that file draws 2 ×
`word-drawing-scalar-whitespace` ERRORs (`'14606\n…'`, `'-107718\n…'`; the other
four sites sit inside the unselected Choice, which the effective view drops)
plus 4 ERRORs for `docProps/app.xml` booleans with value `"false\n  "` — the
same injection. `pack` cannot undo it: `condenseXml`
(`src/scripts/office/pack.ts:312-358`) removes only whitespace-**only** text
nodes (skipping localName `t`), and `14606\n…` is not whitespace-only —
contradicting AGENTS.md note 3's "contents round-trip identically" for
text-bearing elements.

**Proposed fix (fix target: docx-validate):** in `prettyXml`, keep an element
that holds only text (and optionally comments) on one line — emit
`<wp:posOffset>14606</wp:posOffset>` inline instead of breaking before the
closing tag; `condenseXml` cannot repair this class after the fact.

## Where the fixes belong and why

- **docxodus** — violation 1. The redline's XML is genuinely invalid: the
  generator wrote a Choice whose `Requires` token it never declared, so the MCE
  expression can never select and Word offers recovery. Only the generator can
  stop emitting this; the validator already fires correctly.
- **docx-validate** — the round-trip corruption behind violation 2. Not a
  missing check (the check fired on the round-tripped file) and not a repair
  prescription (the source .docx needs none for this), but `prettyXml`'s
  text-only-element formatting plus `condenseXml`'s inability to reverse it,
  which together turn a clean part into a Word-blocking one on
  unpack→edit→pack.

## Verification

1. Reproduced violation 1: `pnpm exec tsx src/scripts/office/validate.ts
   really_broken/tool_redlines/repair_prompt/docxodus__01a11a49_2352.docx` →
   exactly 2 ERROR lines on `word/document.xml` (`mce-malformed`,
   `word-choice-requires-undeclared`). `--strict` matched 0 `xsd-error` lines —
   consistent: `Requires` is `xsd:string`, and the strict walk validates the
   effective view.
2. Disproved violation 2 on the real file: scalar extraction from `unzip -p`
   (clean), isolated `--checks word-drawing-scalar-whitespace` (0 issues), and
   `fromArchive`/`effectiveXml` introspection (text nodes trim-equal).
3. Demonstrated the round-trip bug: `unpack --merge-runs false
   --simplify-redlines false` → `pack --validate false` → the round-tripped part
   carries `'14606\n…'` scalars and draws 2 ×
   `word-drawing-scalar-whitespace` plus 4 app.xml boolean-whitespace ERRORs. A
   `prettyXml` unit test pinning `<wp:posOffset>14606</wp:posOffset>` inline,
   plus an unpack→pack content-parity test over this file's `word/document.xml`
   and `docProps/app.xml`, would lock the fix.
4. Minimal-change experiment proving the Word cause (none is on record): add
   `xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"`
   to the `w:document` root of an otherwise untouched copy and open it in Word
   twice with logs (repo norm). The fixture pair above predicts a clean open;
   if the repair prompt persists, record the remaining cause as unknown.
