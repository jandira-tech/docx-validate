# docx-validate

[![CI](https://github.com/jandira-tech/docx-validate/actions/workflows/ci.yml/badge.svg?branch=main)](https://github.com/jandira-tech/docx-validate/actions/workflows/ci.yml)
[![CodeQL](https://github.com/jandira-tech/docx-validate/actions/workflows/codeql.yml/badge.svg?branch=main)](https://github.com/jandira-tech/docx-validate/actions/workflows/codeql.yml)
[![codecov](https://codecov.io/gh/jandira-tech/docx-validate/branch/main/graph/badge.svg)](https://codecov.io/gh/jandira-tech/docx-validate)
[![Known Vulnerabilities](https://snyk.io/test/npm/docx-validate/badge.svg)](https://snyk.io/test/npm/docx-validate)
[![npm](https://badgen.net/npm/v/docx-validate?icon=npm)](https://www.npmjs.com/package/docx-validate)
[![downloads](https://badgen.net/npm/dm/docx-validate)](https://www.npmjs.com/package/docx-validate)
[![types](https://badgen.net/npm/types/docx-validate)](https://www.npmjs.com/package/docx-validate)
[![bundle min](https://badgen.net/bundlephobia/min/docx-validate)](https://bundlephobia.com/package/docx-validate)
[![bundle gzip](https://badgen.net/bundlephobia/minzip/docx-validate)](https://bundlephobia.com/package/docx-validate)
[![tree-shakeable](https://badgen.net/bundlephobia/tree-shaking/docx-validate)](https://bundlephobia.com/package/docx-validate)
[![dependencies](https://badgen.net/bundlephobia/dependency-count/docx-validate)](https://bundlephobia.com/package/docx-validate)
[![license](https://badgen.net/badge/license/Apache-2.0/blue)](./LICENSE)
[![node](https://badgen.net/badge/node/%3E%3D20/green)](./package.json)
[![github](https://badgen.net/badge/icon/jandira-tech%2Fdocx-validate?icon=github&label)](https://github.com/jandira-tech/docx-validate)

OOXML validators and redline/comment helpers for `.docx` and `.pptx` files. Portable Word checks by default for DOCX, with optional XSD validation and real Microsoft Word verification. TypeScript, ESM, runs under Node (managed with pnpm) for the neurotic developer.

- Repo: [jandira-tech/docx-validate](https://github.com/jandira-tech/docx-validate)
- npm: `docx-validate`
- Maintainer: [jandira.tech](https://www.jandira.tech) — We are building legal tech. Jandira Technologies is the studio behind tools like [Cicero](https://www.cicero.im) (a legal workbench that turns messy inputs into redlines, issue lists, and memos), PII redaction models for Brazilian Portuguese, and AI/contract-drafting benchmarks. `docx-validate` falls out of that work — when redlines have to round-trip through Word, validating the OOXML directly beats trusting the renderer.
- docs: https://jandira-tech.github.io/docx-validate/

## Install

```bash
bun add docx-validate
# or
npm install docx-validate
```

## Layout

```
src/
  index.ts                       — barrel re-exports of the public surface
  lib/
    xml-helpers.ts               — DOM parse/serialize/pretty-print, NS helpers
    types.ts                     — shared types (validation results, OOXML namespaces)
    run-cli.ts                   — one-liner replacement for `if __name__ == "__main__"`
  scripts/
    accept-changes.ts            — accept tracked changes
    comment.ts                   — comment helpers
    office/
      pack.ts                    — repack an unpacked OOXML directory
      unpack.ts                  — unpack an OOXML file
      validate.ts                — top-level validate CLI
      soffice.ts                 — LibreOffice (`soffice`) wrapper
      schemas/                   — bundled OOXML XSD schemas (ISO/IEC 29500-4, ECMA OPC, MCE, Microsoft)
      helpers/
        merge-runs.ts            — merge adjacent w:r runs
        simplify-redlines.ts     — collapse/clean tracked changes
      validators/
        base.ts                  — XSD schema dispatch (libxml2-wasm via lib/xsd-validator.ts)
        docx.ts                  — DOCX validator
        pptx.ts                  — PPTX validator
        redlining.ts             — w:ins / w:del / w:moveTo etc. validator
scripts/                         — repo tooling (not shipped in the package)
  fixture-fingerprint.ts         — docx content fingerprint (JSZip + validate())
  derive-fixture-name.ts         — deterministic <subject>.<comment-or-error> naming
  apply-fixture-names.ts         — rename/sort fixtures by content (dedup, categories)
  update-manifest.ts             — regenerate tests/fixtures-all.manifest.json
tests/                           — vitest specs
  fixtures/
    broken/                      — synthetic single-defect specimens
    working/                     — valid / repaired specimens
    eigen/                       — real-world Plate/SuperDoc specimens (content-named)
    external/                    — vendored third-party specimens (see NOTICE/LICENSE)
```

## Conventions

- File names: kebab-case (`merge-runs.ts`, not `merge_runs.ts` or `mergeRuns.ts`).
- Modules: ES modules (`import`/`export`), no CommonJS.
- Imports: relative paths between sibling modules; no path aliases.
- XSD validation goes through `libxml2-wasm` behind the `XsdValidator` interface (`src/lib/xsd-validator.ts`).
- DOM work goes through `@xmldom/xmldom` + `xpath` via `lib/xml-helpers.ts` (`parseXml`, `serializeXml`, `prettyXml`, `getElementsByTagNameNSAll`) — do NOT call `DOMParser` directly so the implementation stays swappable.
- Temp dirs use the `tmp` package via `withTempDir(async (dir) => { ... })` (see `lib/run-cli.ts`).
- CLIs use `commander`. Wire each script with `runCli(...)` from `lib/run-cli.ts`.
- Validation results: return the `ValidationResult` shape from `lib/types.ts`; do not invent per-validator shapes.
- Tests: vitest, one test file per module, fixtures under `tests/fixtures/`.

## Dependencies

Runtime:

- `jszip` — Zip read/write
- `@xmldom/xmldom` — DOM API
- `xpath` — XPath queries against `@xmldom`
- `libxml2-wasm` — XSD validation (WASM libxml2; replaced the native `libxmljs2` addon, which cannot build on Node 26)
- `commander` — CLI argument parsing
- `tmp` — temp directory lifecycle

Dev:

- `@types/tmp`, `@types/node`, `typescript`, `vite-plus`, `bumpp`

## Running CLIs

The installed package ships a `docx-validate` bin that runs the validator CLI
(`dist/cli.mjs`):

```bash
npx docx-validate <path>          # or: node node_modules/docx-validate/dist/cli.mjs <path>
docx-validate --help              # prints only the validator's usage
```

`dist/index.mjs` is the library entry only — running or importing it never
starts a CLI.

From source, every script is still its own `tsx` entry point:

```bash
pnpm exec tsx src/scripts/office/validate.ts <path>
pnpm exec tsx src/scripts/office/unpack.ts <path>
pnpm exec tsx src/scripts/office/pack.ts <dir>
pnpm exec tsx src/scripts/accept-changes.ts <path>
pnpm exec tsx src/scripts/comment.ts <path>
```

## Programmatic use

The validator and packaged document helpers are available as plain function/class
imports. The package's barrel (`src/index.ts`) is hand-maintained (the
`export *` section and the four-class surface/`jubarte` helpers are curated).
Only the exports listed there are reachable from the package root; the registry,
scheduler, provenance, and verifier implementation modules remain source internals.

### Validate a `.docx` / `.pptx`

```ts
import { validate } from "docx-validate";

const result = await validate("./contract.docx");
//                ^? Promise<ValidateRunResult>
//                   { valid: boolean; issues: ValidationIssue[]; suffix: ".docx"; repairs: number }

if (!result.valid) {
    for (const issue of result.issues) {
        console.error(`${issue.severity}${issue.path ? ` [${issue.path}]` : ""}: ${issue.message}`);
    }
}
```

DOCX and DOCM default to `word-valid`: no known Word-opening defect detected by
the complete portable check set. This is a prediction, not an observed Word open.
PPTX keeps the `lenient` default and its existing XSD validator. The Word-valid
behavior below applies to DOCX/DOCM. Pass `{ profile: "lenient" }` (CLI:
`--profile lenient`) to retain the earlier DOCX policy.

The `strict` profile runs the Word rules and applicable XSD schemas. Schema
conformance is stricter than Word openability. A leading BOM is at most a warning
under every profile and never fails a document on its own:

```ts
const result = await validate("./contract.docx", { profile: "strict" });
```

Cross-check tracked changes against an original (the `--original` CLI flag):

```ts
const result = await validate("./redlined.docx", {
    original: "./baseline.docx",
    author: "Alice",
});
```

Word-openability checks run without XSD. The path API retains its XSD walk trace:

```ts
const result = await validate("./contract.docx", { profile: "word-valid" });
console.log(result.xsdWalk); // { started: false, parts: [], eligible: 0 }
console.log(result.metadata?.schemaCompleted); // false
```

For DOCX/DOCM, the path API and document validator do not compile or walk XSD
schemas under `word-valid`, so `schemaCompleted` is always `false`. Full `lenient` and `strict` runs walk eligible parts; a malformed part
can stop a lenient run before XSD. The XML part list uses OPC-equivalent
extensions: ASCII case is ignored and unreserved escapes are decoded, so `.XML`,
`.RELS` and `.%72els` parts are seen by the passes that iterate that list and by
the XSD walk. DOCX relationships validation passes use the same OPC-equivalent
`.rels` predicate. Relationship-id and hyperlink checks also resolve sidecars
by OPC part-name equivalence in packed and directory inputs, including full,
named and quick runs: a `.RELS` sidecar resolves like `.rels`, matching Word
16.115's clean open. An escaped stored name such as `.%72els` still fails full
Word-valid and strict validation, and quick Word-valid runs, through
`opc-part-name-escaped-unreserved` (OPC [M1.8]),
even though sidecar lookup resolves it ([#402](https://github.com/jandira-tech/docx-validate/issues/402)).
Named selections diagnose the escaped name when they include that check.
Empty-part auto-repair retains its exact lowercase `.rels` scope. The trace describes the schema walk, not a Microsoft Word invocation. Under `strict`, ISO Strict OOXML parts use the bundled
`schemas/iso-strict/` schemas. Lenient skips those parts with the info-level
`xsd-strict-skipped`; Word-valid emits no XSD skip issue because it runs no XSD.

Full DOCX/DOCM validation reports duplicate relationship `Id` values as
`rels-id-duplicate`, judging the relationships part of every owner part in
the package, whatever the owner's name suffix. Under `word-valid`, duplicates
block opening only in the
package relationships part (`_rels/.rels`) and the relationships part of a
story root: the main document, resolved through the package's Transitional or
ISO Strict `officeDocument` relationship wherever it is stored, and every
glossary document a main-document `glossaryDocument` relationship reaches.
Word 16.115 offers recovery even for identical or unreferenced duplicates
there; a header part's own duplicate unreferenced hyperlinks opened clean, and
so did the same pair in an orphan `word/document.xml` sidecar no main document
owns, while a reached glossary's own sidecar was refused, as was the same pair
in the sidecar of a main document or reached glossary stored as
`document.bin` ([#615](https://github.com/jandira-tech/docx-validate/issues/615), session 29,
[#617](https://github.com/jandira-tech/docx-validate/issues/617), session
30, and [#618](https://github.com/jandira-tech/docx-validate/issues/618), session
31, in [the Word probe
evidence](docs/plans/evidence/word-probes-2026-10-24.md)). Other sidecars
retain their strict error and Word-valid warning, and a `.rels` whose owner
part is not in the package is not judged by this walk. Lenient keeps its
existing sidecar errors and XSD-only package diagnosis. This legacy ID walk is
outside the named-check registry, so named and quick runs omit it.

Named checks and quick checks use the same options in the path and byte APIs:

```ts
import { readFile } from "node:fs/promises";
import { Validate, jubarte, validate } from "docx-validate";

const quick = await validate("./contract.docx", { quick: true });
console.log(quick.metadata?.omittedChecks);
const bytes = new Uint8Array(await readFile("./contract.docx"));
const selected = await new Validate({ checks: ["tables"] }).run(bytes);
const equivalent = await jubarte.validate(bytes, { checks: ["tables"] });
```

`Validate` and `jubarte.validate` also forward `profile`, `strict`, `schemasDir`,
`xsdValidator`, `verifyWord`, and `wordVerifier`. Set `suffix` on byte APIs for
formats other than their default `.docx` (for example, `.pptx` or `.docm`).

- `--quick` / `{ quick: true }` selects `QUICK_CHECK_GROUPS`: `package`, `xml`,
  `content-types`, `relationships`, `tables`, and `drawings`.
- `--checks tables,word-run-misplaced` / `{ checks: ["tables", "word-run-misplaced"] }`
  selects check ids or group names, plus prerequisites. Unknown ids and explicitly
  named checks for another format are configuration errors.
- A partial run applies its profile as a full run does, so it never fails a package
  the full run passes. `lenient` runs neither the Word rules nor the ZIP-envelope
  checks: only archive-only checks and Word rules are refused — naming one, or
  only groups made of them, is a configuration error. Under `word-valid`, an
  error Word tolerates is reported as a warning.
- `--list-checks` lists every registered id, group, format, and description without
  reading a file or starting XML/XSD engines. `--help` also needs no engine startup.
- `--strict` / `{ strict: true }` is shorthand for `profile: "strict"`. Strict
  rejects partial selection; `--quick` and `--checks` are mutually exclusive.
  `--strict` also rejects an explicit profile other than `strict`. API conflicts
  throw `ModeConflictError`; CLI configuration errors exit 1.
- Partial runs perform no XSD validation and report `omittedChecks`. A successful
  partial run prints “Selected checks passed”, followed by counts, never “All
  validations PASSED!”, even when it omitted no registered check (a PPTX or XLSX
  registers only the envelope and XML syntax checks). An `Also not run:` line names
  what has no check id: the format validator's own rules, where a full run walks
  schemas the schema walk, and for a DOCX given `--original` the redlining
  validation against it. It does not establish full-document validity. Selections run the registered
  checks only, without the full validator's other checks or severity remapping.
- `--json` emits `{ valid, issues, metadata }` for a completed validation, including
  warnings and available locations. Exit status is 0 for valid and 1 for invalid;
  configuration/startup errors go to stderr. The path API also returns `suffix`,
  `repairs`, and `xsdWalk`, which are not included in CLI JSON.

```bash
docx-validate --list-checks
docx-validate --quick contract.docx
docx-validate --checks tables,word-run-misplaced --json contract.docx
docx-validate --strict contract.docx
```

Full lenient and strict CLI runs first compile a small startup probe schema to
check the XSD engine. This is separate from document validation and is not
recorded in `xsdWalk`. Word-valid runs compile no schema, so they skip that probe,
as do help, discovery, and partial runs (#318).

The registry in `src/scripts/office/checks/index.ts` includes `WORD_RULE_CHECKS`,
which full DOCX validation runs under `word-valid` and `strict`, alongside its
other validators. Full lenient validation does not run those Word rules. A
`ValidationContext` caches part bytes and DOMs per run. Checks declare `requires`;
the scheduler runs prerequisites once, in registration order. `xml-syntax` gates
by part: dependent rules skip malformed parts and keep judging well-formed parts.
A failed prerequisite without part-scoped findings skips its dependents. A
throwing check yields `check-crashed`. For unpacked directories, `runChecks()`
marks archive-only checks `unavailable`; use packed input to check ZIP integrity.

The complete named check catalogue is below. Each Word check except
`word-markup-compatibility` uses its id as its finding code. Package checks are
`zip-bytes-after-eocd`, `zip-bytes-before-archive`, `zip-eocd-comment-overrun`,
`opc-equivalent-part-names`, `opc-part-name-escaped-unreserved`,
`opc-part-name-escaped-slash`, and `opc-part-name-word-refused`; `xml-syntax` is the shared
XML prerequisite. All `word-*` checks apply to DOCX/DOCM only.
`ct-duplicate-declaration`, `ct-override-partname-invalid`,
`ct-override-partname-word-refused` and `ct-markup-compatibility` apply to
every package format; `word-content-type-invalid`,
`word-part-content-type-generic`, `word-part-content-type-inexact` and
`word-part-content-type-mismatch` apply to DOCX/DOCM only.

`zip-bytes-before-archive` applies to DOCX/DOCM, PPTX, and XLSX package
selections, including the `package` quick group, and reads the archive bytes
before extraction. It reports the bytes in front of the archive JSZip loads —
the rebase offset plus the lowest local-header offset as JSZip resolves it,
so a prepended ZIP, junk with relative offsets, junk with the absolute
offsets a self-extractor writes, and the ZIP64 spelling (a saturated offset
whose 0x0001 extra field carries the value) all measure the same prefix — as
one blocking error under strict and Word-valid. The measure is undefined
exactly where JSZip's reader throws — the local header's extent, its name and
extra fields (each bounded on its own read, so a small negative resolved size
cannot mask a name/extra overrun), the central-declared compressed stream (bounded on both sides:
stream past the buffer, or a compressed size that reads negative — the 32-bit
field 0x80000000 or a ZIP64 0x0001 value below zero — sending the final read's
index below zero), and a known compression
method (store or deflate) are modelled too — so a package whose local part
JSZip cannot read reports no prefix
([#524](https://github.com/jandira-tech/docx-validate/issues/524),
[#555](https://github.com/jandira-tech/docx-validate/issues/555),
[#566](https://github.com/jandira-tech/docx-validate/issues/566)). Word 16.115 offers to recover
every prefixed package probed
([#464](https://github.com/jandira-tech/docx-validate/issues/464);
probe session 14 in
[docs/plans/evidence/word-probes-2026-10-24.md](docs/plans/evidence/word-probes-2026-10-24.md)).
Lenient runs no ZIP-envelope check. A concatenated `[zip][docx]` is reported
once, by this check: the earlier archive lies wholly in front of the archive
JSZip loads, so `zip-bytes-after-eocd` no longer counts the live archive as
undeclared trailing bytes — the earlier record's fixed extent decides, even
when its declared comment reaches into the live archive
([#491](https://github.com/jandira-tech/docx-validate/issues/491)). Trailing
data behind a prefixed archive is still `zip-bytes-after-eocd`'s finding,
measured from the loaded archive's own record.

`opc-equivalent-part-names` applies to DOCX/DOCM, PPTX, and XLSX package selections,
including the `package` quick group. It reads stored ZIP item names before full-run
extraction, or walked file names for directory inputs, and reports one finding per
equivalent group under OPC [M1.12], listing every stored spelling. Strict and
word-valid (DOCX/DOCM) report a blocking error: Word 16.115 offers to recover every
duplicate probed, even when both spellings hold identical bytes ([#351](https://github.com/jandira-tech/docx-validate/issues/351)).
Lenient, and word-valid for PPTX/XLSX, report a warning. The existing entry choice
is unchanged: two entries stored under the identical spelling report the same
finding, naming the spelling once per entry (Word offers to recover these too),
and the package view reads the last stored entry (#411). The stored names come
from the central directory of the archive JSZip loads: the last end-of-central-directory
signature, with its offsets rebased past any prepended bytes, so the names always
describe the package view, even behind trailing data or a prepended archive (#447, #462).
This check, the item-name refusal and the package view the checks judge all use
those central-directory names, with one shared CP437 decoder for names without
general-purpose bit 11 and UTF-8 for flagged names (#480). Re-keying uses the
central spelling, as Word 16.115 does, and ignores the local file header's name
and the Info-ZIP unicode-path extra field (0x7075) (#445). Re-keying covers
ZIP64 packages: the ZIP64 end-of-central-directory record's directory size and
offset are rebased with the directory, while the 32-bit record keeps its
saturated fields and the locator's record offset keeps naming the record's
absolute position — JSZip reads that field before its rebase, so a prefixed
self-extractor package whose locator names the record absolutely, even one
whose EOCD comment carries a fake ZIP64-record signature, still loads after
re-keying (#479, #521). If re-keying is
unavailable (for example, an entry whose local-header offset or compressed
size lives in a ZIP64 extra field), JSZip still prefers a valid central Unicode
Path field. An entry's kind follows the central name as well: JSZip decides
`dir` from the name it reads, so a central name ending in `/` is a folder in
the package view and its part leaves the view ([#482](https://github.com/jandira-tech/docx-validate/issues/482)). `unpack` is the exception: it extracts under raw JSZip's own names,
the local header's or the unicode-path field's (#469), with UTF-8 as its default
decoder. A flag-less CP437 `caf\x82.xml` therefore extracts as `caf�.xml`, while
the validator's archive view uses `café.xml`; directory validation of that
extracted tree sees `caf�.xml` (#480).

`opc-part-name-escaped-unreserved` uses the same formats, package-group selection,
pre-extraction timing, directory support, and severity policy. It reports one
finding per stored name containing a percent-encoded unreserved character
(`A–Z`, `a–z`, `0–9`, `-`, `.`, `_`, `~`), as OPC [M1.8] forbids, in code-unit
order with the stored name as `path`. Word 16.115 offers recovery for every
escape probed: letters in either case, a digit, `-` in either hex case, `.`, `_`
and `~`. It does so even on an unreferenced extra item (`word/%78tra.xml`), and
opens the `%20` control clean. Other escapes such as `%20`, `%25`, and
`%5B`, malformed escapes, and plain names do not trigger this check; an escaped
separator has its own check below. An escaped spelling beside its decoded twin reports both M1.8 and M1.12;
validation does not rewrite either name ([#408](https://github.com/jandira-tech/docx-validate/issues/408)).

`opc-part-name-escaped-slash` uses the same formats, package-group selection,
pre-extraction timing, directory support, and severity policy. It reports one
finding per stored name with a segment holding a percent-encoded `/` or `\`
(`%2F`, `%2f`, `%5C`, `%5c`), as OPC [M1.7] forbids, in code-unit order with the
stored name as `path`; the escape must be whole, so `%252F` (an encoded `%`
before `2F`) stays legal. Word 16.115 offers recovery for an unreferenced
`word%2Fxtra.xml`, refuses `word%5Cxtra.xml`, and opens the well-formed
`word/xtra.xml` control clean ([#436](https://github.com/jandira-tech/docx-validate/issues/436)).

`opc-part-name-word-refused` reports one finding per stored file-entry name
containing a percent-escaped RFC 2396 mark (`!`, `'`, `(`, `)`, `*`), a
percent-escaped octet at or above `0x80`, or a literal non-ASCII character.
It reads stored ZIP file-entry names before extraction, including
unreferenced entries and refused package views, and walked file names for
directories. Folder entries — a trailing `/`, the MS-DOS directory attribute,
or both — are not judged, and Word was not probed for a folder entry whose
only defect is its name. Hex is case-insensitive and whole
triplets are consumed: `%2521` stays clean, as do `%20`, `%26`, `%3A`, `%5B`,
`%7F`, and a literal `!`. Unreserved escapes keep their existing M1.8 finding.

`ct-override-partname-word-refused` applies the same escaped-mark, high-octet
and literal non-ASCII rules to every Override PartName, even when no part is
stored under that name. It reports one finding per declaration at the stored
content-types stream, quoting the PartName. A literal non-ASCII PartName is
reported here exactly when the §9.1.1.1 grammar accepts it — a literal
`ucschar` Part IRI such as `/word/café.xml` passes `ct-override-partname-invalid`
(#455), so only this check names Word's refusal — and never repeats the grammar
finding a name outside both forms already has. A spelling containing
both `%41` and `%21` reports both applicable codes once.

Both checks are errors for DOCX/DOCM under word-valid and strict (strict includes
every Word-blocking defect), warnings under lenient, and warnings for PPTX/XLSX
under every profile: these spellings are spec-valid and PowerPoint/Excel refusal
has not been observed. Full, quick and named selections share this policy;
strict still rejects partial selection, and XLSX supports partial runs only.
Word 16.115 offers recovery for these spellings in DOCX packages, as recorded
with SHA-256 hashes in [probe sessions 6–12](docs/plans/evidence/word-probes-2026-10-24.md)
([#451](https://github.com/jandira-tech/docx-validate/issues/451),
[#408](https://github.com/jandira-tech/docx-validate/issues/408),
[#443](https://github.com/jandira-tech/docx-validate/issues/443)). Validation leaves
names and input bytes unchanged.

| Group                     | Named checks                                                                                                                                                                                                                                                                                                                                        |
| ------------------------- | --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `package`                 | `zip-bytes-after-eocd`, `zip-bytes-before-archive`, `zip-eocd-comment-overrun`, `opc-equivalent-part-names`, `opc-part-name-escaped-unreserved`, `opc-part-name-escaped-slash`, `opc-part-name-word-refused`                                                                                                                                        |
| `math`, `tracked-changes` | `word-math-run-insertion`, `word-math-run-deletion`                                                                                                                                                                                                                                                                                                 |
| `math`                    | `word-math-parse`, `word-math-spre-body`                                                                                                                                                                                                                                                                                                            |
| `story`                   | `word-shading-color`, `word-page-margin-attributes`, `word-run-misplaced`, `word-run-properties-misplaced`, `word-underline-misplaced`, `word-paragraph-properties-misplaced`, `word-section-properties-misplaced`, `word-footnote-reference-missing`, `word-numbering-picture-bullet`, `word-sdt-id-empty`, `word-text-contains-paragraph`         |
| `tables`                  | `word-grid-column-attributes`, `word-cell-margin-order`, `word-table-properties-required`, `word-table-properties-misplaced`, `word-row-without-cell`, `word-cell-ends-with-table`                                                                                                                                                                  |
| `drawings`                | `word-textbox-content-namespace`, `word-drawing-extent-missing`, `word-drawing-extent-misplaced`, `word-picture-nvpicpr-missing`, `word-picture-sppr-missing`, `word-empty-drawing`, `word-anchor-missing-attribute`, `word-position-offset-and-align`, `word-drawing-scalar-whitespace`, `word-drawing-arc-radius`, `word-drawing-extent-negative` |
| `content-types`           | `word-content-type-invalid`, `ct-duplicate-declaration`, `ct-override-partname-invalid`, `ct-override-partname-word-refused`, `ct-markup-compatibility`, `word-part-content-type-generic`, `word-part-content-type-inexact`, `word-part-content-type-mismatch`                                                                                      |
| `relationships`           | `word-duplicate-theme-relationship`, `word-duplicate-font-table-relationship`, `word-duplicate-web-settings-relationship`, `word-duplicate-settings-relationship`, `word-hyperlink-relationship`, `word-header-footer-relationship`, `word-main-document-part`, `opc-relationships-structure`                                                       |
| `tracked-changes`         | `word-textbox-deleted-text`, `word-move-from-deleted-text`                                                                                                                                                                                                                                                                                          |
| `settings`                | `word-custom-property-filetime-empty`, `word-attached-toolbar-utf8`                                                                                                                                                                                                                                                                                 |
| `xml`                     | `word-markup-compatibility`, `word-document-root`, `word-choice-requires-undeclared`                                                                                                                                                                                                                                                                |
| `comments`                | `comment-markers`, `comment-threading`                                                                                                                                                                                                                                                                                                              |
| `schema`                  | `word-unknown-element`, `word-link-without-namespace`                                                                                                                                                                                                                                                                                               |

`word-content-type-invalid` also reports an explicitly empty `ContentType=""` on the main document's Override, resolved through the package `officeDocument` relationship and the shared content-type table after Markup Compatibility processing ([#581](https://github.com/jandira-tech/docx-validate/issues/581)). Word 16.115.26100424 offers recovery for this specimen (session 22 in [docs/plans/evidence/word-probes-2026-10-24.md](docs/plans/evidence/word-probes-2026-10-24.md)). It is a blocking error under Word-valid and an error under strict. Lenient runs no Word rule, so it reports nothing here beyond the XSD rejection. Full, quick and named Word-valid runs agree. Empty values on other Overrides or on Defaults stay silent under Word-valid; strict continues to reject them through XSD.

`word-content-type-invalid` reports `invalid-*` content types only when they apply to a stored part: an Override must name an existing part under OPC equivalence, and a Default must cover a stored part without its own Override. Word 16.115 opens an Override for a missing part and an unused Default cleanly; an invalid type on an existing part is refused or recovered.

`ct-duplicate-declaration` reports each repeated `Default@Extension` (ASCII case-insensitive)
or `Override@PartName` (OPC part-name equivalence), even when the declarations have identical
content types. Each finding names both spellings at the stored content-types stream path.
It is an error under strict and under Word-valid for DOCX/DOCM, and a warning under lenient
and Word-valid PPTX/XLSX. Word 16.115 offered recovery for all five observed duplicate shapes;
the unchanged minimal control opened clean ([#431](https://github.com/jandira-tech/docx-validate/issues/431)).
Full runs, quick runs and named selections share the rule. An unparseable stream reports XML
diagnostics without duplicate findings. Strict continues to reject partial selectors;
full XLSX validation remains unsupported.

`ct-override-partname-invalid` reports every `Override` whose `PartName` is not
a valid OPC part name (ECMA-376 Part 2, 4th edition (2012), §9.1.1.1, the
edition of the bundled `schemas/ecma/fourth-edition/` OPC schemas, which strict
follows): a part name shall be an IRI and shall be encoded as either a Part URI
(§9.1.1.1.2: `segment = 1*( pchar )`, RFC 3986) or a Part IRI (§9.1.1.1.1:
`isegment = 1*( ipchar )`, RFC 3987, which allows a literal `ucschar`), so a name
that does not start
with `/`, holds an empty, a `.`/`..` or a dot-only (`...`) segment, ends with
`/` or `.`, has a folder segment ending with `.`, carries a `?`, a `#` or a `\`, percent-encodes a
separator (OPC [M1.7]), an unreserved character (OPC [M1.8]) or — once a literal
`ucschar` leaves only the Part IRI form — an `iunreserved` character (OPC
[M1.8]), or holds a character in neither the RFC 3986 `pchar` nor the RFC 3987
`ipchar` set, is reported. `/word/café.xml` (a Part IRI) and
`/word/caf%C3%A9.xml` (a Part URI, ECMA-376 Example 9–1) are therefore valid,
while `/word/caf%C3%A9-é.xml` is neither, and a literal non-ASCII character that
is not `ucschar` (a C1 control, private use, a noncharacter) is in neither
grammar ([#455](https://github.com/jandira-tech/docx-validate/issues/455)). The
`ucschar` set is RFC 3987 §2.2's: `%xA0-D7FF`, `%xF900-FDCF`, `%xFDF0-FFEF`,
each astral block `%x10000-1FFFD` … `%xD0000-DFFFD`, and the last one,
`%xE1000-EFFFD` — so `U+E0000..U+E0FFF` is in neither grammar and is reported,
while `U+E1000` and `U+EFFFD` are valid
([#522](https://github.com/jandira-tech/docx-validate/issues/522)). A
percent-escape run that does not decode as UTF-8 (a lone `%E9`, an overlong
`%C0%AF`) is pct-encoded octets rather than an encoded character, and the scan
resynchronizes at the next octet, so the `%C3%A9` behind one still encodes an
`iunreserved` character and is reported ([#523](https://github.com/jandira-tech/docx-validate/issues/523)). A
literal `ucschar` the grammar accepts stays refused by Word, which
`ct-override-partname-word-refused` reports (#451). Such a PartName names no
part, so the declaration is malformed whatever its content type. Resolution is
unchanged: content types are still resolved through the canonical spelling, so
a malformed PartName can still resolve a stored part, as `word/document.xml`
without its leading slash resolves the main document, while a PartName with a
dot segment resolves nothing
([#457](https://github.com/jandira-tech/docx-validate/issues/457)). One finding is
reported per offending Override, at the stored content-types stream, quoting
the PartName and naming the rule it breaks; a folder segment ending with `.` is
named only when no other rule is broken. An `Override` without a `PartName` is
reported as having none, and Word offers to recover it too
([#458](https://github.com/jandira-tech/docx-validate/issues/458), session 13). A content-types stream that does not
parse is reported by the XML syntax gate instead, and nothing is reported from
it here. Word 16.115 offers to recover a package for every shape probed, even
when nothing is stored under that name and the content type is well formed,
except a folder segment ending with `.` that holds another character
(`/word./x.xml`), which it opens clean, as it opens a well-formed PartName for a
missing part. A dot-only folder segment (`/word/.../x.xml`) is refused
([#414](https://github.com/jandira-tech/docx-validate/issues/414),
[#450](https://github.com/jandira-tech/docx-validate/issues/450),
[#453](https://github.com/jandira-tech/docx-validate/issues/453); probe
sessions 3c, 4, 5, 7 and 12 in `docs/plans/evidence/word-probes-2026-10-24.md`).
The finding is an error under strict, and an error and Word-blocking under
word-valid for DOCX/DOCM, where a folder segment ending with `.` is a warning;
lenient and the unobserved PPTX/XLSX report a warning. Full, quick and named
runs agree.

`ct-markup-compatibility` reports one finding for a content-types stream that
uses Markup Compatibility: an attribute in the MC namespace (`mc:Ignorable`,
`mc:ProcessContent`, `mc:MustUnderstand`), an element in it
(`mc:AlternateContent` and its branches), or an element or attribute in a
namespace an `mc:Ignorable` lists. A bare `xmlns:mc` declaration that nothing
uses is not reported. The message lists every such element and attribute once
per occurrence, in document order — an element before its own attributes
([#495](https://github.com/jandira-tech/docx-validate/issues/495),
[#526](https://github.com/jandira-tech/docx-validate/issues/526)). Microsoft
Word 16.115 does not process Markup
Compatibility in this stream: it offers recovery for every shape probed — an
`mc:Ignorable` attribute, an ignored foreign element, and an AlternateContent
in either orientation — while the declared-only control opens clean
([#459](https://github.com/jandira-tech/docx-validate/issues/459); probe
session 14, `docs/plans/evidence/word-probes-2026-10-24.md`). The check reads
the raw stored stream, while `ct-duplicate-declaration`,
`ct-override-partname-invalid`, `ct-override-partname-word-refused`,
`word-content-type-invalid`, `word-part-content-type-generic`,
`word-part-content-type-inexact`, `word-part-content-type-mismatch` and content-type
resolution all judge the stream after Markup Compatibility processing, as OPC
[M1.19] requires — so an `Override` in an ignored `mc:Choice` no longer draws
a `ct-override-partname-invalid` finding, and a bad `mc:Fallback` still does.
That processor understands the stream's own vocabulary: an `mc:Ignorable`
naming the content-types namespace itself blinds no coverage or resolution
pass, and `.rels` parts read through the same view add the
package-relationships namespace
([#493](https://github.com/jandira-tech/docx-validate/issues/493)).
The finding is an error, and Word-blocking, under word-valid for DOCX/DOCM and
under strict (whose schema walk already refuses the same stream); lenient and
the unobserved PPTX/XLSX warn. Full, quick and named runs agree.

`word-part-content-type-generic` reports a part Word reads whose content type
resolves to a generic XML type — media type `application/xml` or `text/xml`,
compared ASCII case-insensitively with any MIME parameter stripped, so
`application/xml; charset=utf-8` counts as generic — from either source: an
`Override`, or the `<Default>` for the part's own extension, which the finding
names ([#504](https://github.com/jandira-tech/docx-validate/issues/504),
[#517](https://github.com/jandira-tech/docx-validate/issues/517),
[#540](https://github.com/jandira-tech/docx-validate/issues/540)). The reached parts
are the main document (the package `officeDocument` relationship's target) and
the parts the main document's relationships of type styles, settings,
fontTable, webSettings, numbering or theme name, wherever they are stored.
Resolution follows
ECMA-376 Part 2 §10.1.2.4 through the shared table: a part with a specific
non-generic Override — right or wrong — is out of scope, and under
word-valid and lenient customXml stays outside this check. Under strict the
check also judges every other Part 1
WordprocessingML part type a main-document relationship names —
header, footer, endnotes, footnotes, comments and glossaryDocument — whose
findings cite the content type ECMA-376 Part 1 (4th edition), clause 11.3,
assigns rather than a Word observation
([#506](https://github.com/jandira-tech/docx-validate/issues/506)). Under strict, a Default-typed main or theme finding is suppressed only when
`ct-undeclared-part` actually reports an error on that same path; renamed
main parts, including non-XML names and parts under `docProps/`, retain their
generic-type error ([#535](https://github.com/jandira-tech/docx-validate/issues/535)). A part reached by two
relationship types is judged against every type that reaches it — a
Default-typed theme also named by an unreferenced footer relationship still
counts as the theme for that suppression — and the finding names the
segment the check's fixed reporting order ranks first, an observed
relationship before a spec-only one, so the order of the relationships in
the relationships part decides neither the verdict nor the message
([#593](https://github.com/jandira-tech/docx-validate/issues/593)). One finding
is reported per part, at the part's path, naming the relationship type and
the content type the part needs — for the main document, all four Word
main-document content types, since the package does not say whether it is
macro-enabled
([#505](https://github.com/jandira-tech/docx-validate/issues/505)). Word 16.115 offers to recover the package
for all five parts probed typed only by the generic Default — the main
document and styles with their Override removed, and settings, fontTable and
webSettings added with a relationship but no Override — and opens every
proper-Override control clean ([#432](https://github.com/jandira-tech/docx-validate/issues/432),
session 8a in `docs/plans/evidence/word-probes-2026-10-24.md`). Session 20
adds related numbering (`word/numbering.xml`) and theme
(`word/theme/theme1.xml`) typed only by the `application/xml` Default:
Word 16.115.26100424 raised a repair warning for both, and opened their exact
Override controls clean (session 20, chunks s20c and s20d, in
`docs/plans/evidence/word-probes-2026-10-24.md`; [#571](https://github.com/jandira-tech/docx-validate/issues/571)). Session 23
also observed recovery for a theme only a reached glossary relates, typed
by the `application/xml` Default; its exact-Override control opened clean
([#595](https://github.com/jandira-tech/docx-validate/issues/595)). The
finding names the glossary's theme relationship. Session 24 also observed
recovery for its styles part typed by the generic Default, while the exact
Override control opened clean ([#601](https://github.com/jandira-tech/docx-validate/issues/601)).
The finding names the glossary's styles relationship: error and Word-blocking
under word-valid, error under strict, warning under lenient. Strict retains
this finding because no `ct-undeclared-part` error replaces it at that path.
Only theme and styles join the glossary's generic-type scope; settings,
fontTable, webSettings, numbering and stories remain unprobed. Other generic
spellings are inference from resolution.
Other generic spellings for numbering and theme are inference from the same
resolution. Strict keeps only `ct-undeclared-part` for the observed
Default-typed theme, while numbering keeps `word-part-content-type-generic`;
full, quick and named word-valid runs agree, and lenient warns. This does
not widen the separate inexact-type check's scope. Session 15
probed the main document and styles with `text/xml` by Default or Override
and with `application/xml` by Override, and Word offers to recover every
spelling; the Override and `text/xml` spellings for settings, fontTable and
webSettings are inference from the same resolution ([#517](https://github.com/jandira-tech/docx-validate/issues/517)).
Session 16 probed styles typed `application/xml; charset=utf-8` by an
Override — the media-type comparison above comes from that observation —
and the parameterized spellings on settings, fontTable and webSettings
are inference from the same resolution
([#540](https://github.com/jandira-tech/docx-validate/issues/540)).
The finding
is an error, and Word-blocking, under word-valid for DOCX/DOCM, and an error
under strict, where a Default-typed main is not reported twice when
`ct-undeclared-part` rejects that same path; lenient
reports a warning. Full,
quick and named runs agree.

`word-part-content-type-mismatch` reports a related part whose declared media
type matches none of the judged relationship segments, but is another
part's expected type in the shared Word type map
([#600](https://github.com/jandira-tech/docx-validate/issues/600)). Media types
are trimmed, compared ASCII case-insensitively, and stripped of parameters.
Generic XML, an inexact spelling of the part's own media type, and unknown
types such as `image/png` stay outside this check, so the generic, inexact
and mismatch checks are disjoint. A part read through story references is judged against the segments
those references use. An unused matching relationship does
not exempt it; a match through either of two used references still does.
Unread parts are judged against all reaching segments under strict (#607).

Word-valid reports only read header/footer swaps: a referenced footer typed
`…header+xml`, or a referenced header typed `…footer+xml`, from the main
document or a reached glossary. Session 24 observed Word 16.115 offer recovery
for the main document's header and footer swaps and the glossary footer swap,
including `…header+xml; charset=utf-8`; their exact-type controls and the
minimal package opened clean. Session 26 observed recovery for the glossary
header typed `…footer+xml` and for a read footer typed `…header+xml` despite
an unused header relationship to it; the latter's `…footer+xml` control
opened clean. The recovery sentence cites session 24 for the main header,
main footer and glossary footer swaps, and session 26 for the glossary
header swap (#610). The finding names the relationship, declared
type and expected type, errors and blocks Word-valid; lenient warns. Unread
stories, a footer typed `…styles+xml`, and comments/footnotes/endnotes typed
with each other's types stay silent under Word-valid until probed.

Strict judges every related part in the shared map, read or unread, citing
ECMA-376 Part 1 (4th edition), clause 11.3's content-type assignments (theme
is DrawingML and cites clause 14.2.7). The main document is named the Main
Document part and cites clause 11.3.10 (#609). Candidate relationships are
ranked in the same fixed order as the generic check, so relationship entry
order does not change the reported segment or clause (#608). Only an
observed read header/footer swap carries the Word recovery sentence;
unobserved shapes keep spec wording. Full, quick and named selections agree
(#381).

`word-part-content-type-inexact` — renamed from
`word-part-content-type-parameter` for
[#550](https://github.com/jandira-tech/docx-validate/issues/550) — reports a
part Word reads whose declared content type is not exactly the part's own
Word type ignoring ASCII case, while its media type — trimmed, parameters
stripped — is: any `; …` parameter section, with or without a space and
whether `charset` or not, and padding whitespace. Word 16.115 refuses
`; charset=utf-8` even where RFC 7303 §9.1 registers the parameter, refuses
a trailing space, and opens an upper-case spelling clean — it folds ASCII
case and nothing else
([#540](https://github.com/jandira-tech/docx-validate/issues/540),
[#550](https://github.com/jandira-tech/docx-validate/issues/550); probe
sessions 16–17, 2026-10-24, the table living with issue #550). Scope is the
parts sessions 17 and 20 observed: the main document — all four Word
main-document content types
([#505](https://github.com/jandira-tech/docx-validate/issues/505)) —
styles, settings, fontTable, webSettings, a header the main document
references through `w:headerReference`, a related numbering or theme part
(both eagerly parsed), and a footer, footnotes, endnotes or comments part
the document uses; a related-but-unused lazy part opens clean
([#516](https://github.com/jandira-tech/docx-validate/issues/516)) and
stays silent. Session 20 probed the types #550 left open: an inexact
Override gets the repair warning on each of them, while the same spelling
on a glossaryDocument opens clean and stays silent under word-valid.
Session 21 observed a trailing space refused on the same parts, and a
related but unused footer, footnotes, endnotes, comments or header part
with a parameter opening clean
([#570](https://github.com/jandira-tech/docx-validate/issues/570); sessions
20 and 21 in `docs/plans/evidence/word-probes-2026-10-24.md`). Sessions 21
and 22 probed the parts a reached glossary's own relationships part names:
an inexact Override on its styles part — eagerly parsed, so related is
enough — and on a header its `w:headerReference` uses gets the repair
warning, and the finding names the glossary's relationship, not the main
document's
([#579](https://github.com/jandira-tech/docx-validate/issues/579)). Session 23
observed the same recovery offer for a `; charset=utf-8` Override on a footer,
footnotes, endnotes or comments the glossary references, with every exact
type control opening clean
([#596](https://github.com/jandira-tech/docx-validate/issues/596)). These
stories are gated by `wordReadParts` just like the glossary's header:
unreferenced parts stay silent under word-valid, while strict uses the
spec text without a Word recovery claim (#589). The
glossary's settings, fontTable, webSettings, numbering and theme were not
probed for an inexact type and stay silent under word-valid; strict judges
every Part 1 WordprocessingML relationship — and theme — on the glossary
too, citing clause 11.3 for the unprobed WordprocessingML segments; the
theme, a DrawingML part, cites no clause. A part the main document
and the glossary both relate by the same type is reported once, naming the
main document's relationship unless only the glossary reads the part through
it. A part several relationship types reach is judged
against each: the finding names the type whose expected content type's
media type matches the declared value — the types are pairwise distinct,
so at most one can — and an unreferenced lazy relationship neither reports
nor hides a read one, because usage is decided per (part, type) pair
([#593](https://github.com/jandira-tech/docx-validate/issues/593)). Strict
judges the same probed set — theme included, though it is a DrawingML part,
not a clause-11.3 one — plus the unprobed Part 1 WordprocessingML
glossaryDocument type, citing the content type ECMA-376 Part 1
(4th edition), clause 11.3, assigns
([#506](https://github.com/jandira-tech/docx-validate/issues/506)). A
strict finding for a lazy part Word does not read — a related-but-unused
header, footer, comments, footnotes or endnotes part, the main document's
or a reached glossary's header — cites clause 11.3 the same way and claims
no Word recovery: session 21 (s21b) observed those specimens open clean
([#589](https://github.com/jandira-tech/docx-validate/issues/589); session
21 in `docs/plans/evidence/word-probes-2026-10-24.md`), so the recovery
sentence stays on the parts Word reads — eager parts and used lazy parts
alike, whose findings keep the observed wording under strict exactly as
word-valid reports it. Lenient
reports a warning. One finding is reported per part, naming the declared
value — quoted, so padding is visible — and whether a Default or an
Override produced it; the generic check judges the media type with
parameters stripped, so a part is never reported by both checks. Under
strict, a Default-typed main finding is suppressed only when
`ct-undeclared-part` actually reports an error on that same path, exactly
as the generic check's
([#535](https://github.com/jandira-tech/docx-validate/issues/535)). Full,
quick and named runs agree.

The `schema` group contains structural checks, not an XSD walk. The `comments`
group runs `comment-markers` (range pairing and references to existing comments)
and `comment-threading` (marker counts, paragraph ids, parent ids, and durable ids
across comment extension parts). Both apply to DOCX/DOCM under every profile,
including lenient; they are separate from `WORD_RULE_CHECKS`. They retain their
own `comment-marker-parse` / `comment-thread-parse` diagnostics rather than
requiring `xml-syntax`. A named run that selects them also runs `xml-syntax`
and reports a malformed part once, through `xml-syntax`, as the full run does
([#385](https://github.com/jandira-tech/docx-validate/issues/385)). Under lenient that run stops where `xml-syntax` rejects a part, as the
full lenient run stops at a malformed part, so the comment checks then do not run
([#391](https://github.com/jandira-tech/docx-validate/issues/391)). Full validation runs them once and records their ids in
`executedChecks`; `--quick` does not select them. Like every other check, they
find a comment part under an OPC-equivalent name (`word/comments.XML`).

Named `word-math-parse` runs and full runs that cannot stage the archive retain
the positioned XML diagnostics without repeating `word-math-parse` for the same
malformed part ([#500](https://github.com/jandira-tech/docx-validate/issues/500)).

The math-run insertion/deletion, shading-color, grid-column-attributes,
cell-margin-order, table-properties-required, page-margin-attributes, and
textbox-content-namespace checks emit warnings. They remain advisories in
Word-valid; strict additionally enforces the applicable schemas. Shading checks
cover stories, `styles.xml`, and `numbering.xml`; math revision checks support
both Transitional and ISO Strict namespaces. Other registered Word findings are
errors before full Word-valid policy is applied. Compatibility finding codes
and profile-specific severities are described below.

XML syntax is checked from original part bytes with libxml2, without parser
recovery. UTF-8 and UTF-16 parts are decoded by BOM/declaration for the DOM layer.
Diagnostics use `xml-syntax`, `xml-undeclared-prefix`, and `xml-encoding`, with the
part name and available 1-based line/column. Every part whose extension is `.xml` or
`.rels` as OPC compares part names (any case, `word/COMMENTS.XML`; escaped,
`word/comments.%78ml`) is checked, by full and partial runs alike
([#390](https://github.com/jandira-tech/docx-validate/issues/390)). Word-valid blocks on a malformed part by
reachability, not by folder ([#437](https://github.com/jandira-tech/docx-validate/issues/437)):
the main document wherever it is stored, its own relationships part — Word 16.115
offers to recover a `word/_rels/document.xml.rels` that is zero bytes or missing
its root end tag, wherever the main document is stored
([#514](https://github.com/jandira-tech/docx-validate/issues/514); probe session 15) — and
the parts a main-document relationship names: eagerly parsed types always
(styles, settings, fontTable, webSettings, numbering, customXml, theme,
glossaryDocument), lazy types only when the document uses them (footer and
header through their `w:*Reference r:id`, comments through
`w:commentReference`, footnotes and endnotes through their references). Word
16.115 reads a used header or endnotes part and a theme or glossaryDocument
part wherever they are stored — offering repair or refusing the file when one
is malformed — and opens a related-but-unused header or endnotes part clean,
even under `word/` ([#516](https://github.com/jandira-tech/docx-validate/issues/516); probe session 15).
A glossary document reached through a `glossaryDocument` relationship is a
second story root: Word 16.115 reads its own relationships part — offering
repair for one that is unclosed or zero bytes, even when the glossary
references nothing — and reads the header, footer, comments, footnotes and
endnotes it references (`w:headerReference` / `w:footerReference` `r:id` in
its `w:docPartBody` `w:sectPr`; `w:commentReference`, `w:footnoteReference`
or `w:endnoteReference` in its runs), while a header it relates but does not
reference opens clean
([#539](https://github.com/jandira-tech/docx-validate/issues/539),
[#549](https://github.com/jandira-tech/docx-validate/issues/549); probe
sessions 16 and 17). References from inside a footnote are not followed — an
endnotes part referenced only from there opens clean. A reached glossary's
styles, settings, fontTable, webSettings, numbering and theme relationships
are eagerly parsed like the main document's — a malformed one gets the
repair offer even when the glossary references no story, while a `customXml`
or nested `glossaryDocument` it relates is observed, not read — malformed,
both open clean — and stays unfollowed
([#562](https://github.com/jandira-tech/docx-validate/issues/562),
[#569](https://github.com/jandira-tech/docx-validate/issues/569); probe
sessions 19 and 20).
When the main document or a glossary fails byte-level XML checks, including an
unsupported encoding declaration, lazy usage is unknown and those related
parts stay reached in full, partial and staging-fallback runs
([#536](https://github.com/jandira-tech/docx-validate/issues/536)).
Any other part under `word/` that a known Office relationship of an unobserved
type names also blocks — a type has an Office segment only when it is exactly
an Office relationship-type namespace plus one segment, so
`…/relationships/custom/styles` is an unobserved type, not `styles`
([#519](https://github.com/jandira-tech/docx-validate/issues/519)); on the content-types
stream and the package relationships part `_rels/.rels` it always blocks.
`word/` is matched as OPC compares part names, so `WORD/document.xml` counts, and
the stream is matched in any case (`[CONTENT_TYPES].XML`), as System.IO.Packaging
finds it ([#392](https://github.com/jandira-tech/docx-validate/issues/392)). Word 16.115 offers to recover a package whose `_rels/.rels`
is zero bytes, truncated, or uses an undeclared prefix ([#388](https://github.com/jandira-tech/docx-validate/issues/388)). An undecodable part (`xml-encoding`)
follows the same policy: Word refuses a `word/document.xml` that declares an
unknown encoding ([#397](https://github.com/jandira-tech/docx-validate/issues/397)), and one that declares a real encoding
libxml2 cannot decode (`macintosh`, `x-mac-cyrillic`, `IBM866`, `UTF-7`), so the
blocking covers those labels too ([#410](https://github.com/jandira-tech/docx-validate/issues/410)).
Word-valid also blocks `xml-syntax`, `xml-encoding`, and `xml-undeclared-prefix`
on a properties part named by a core, extended, or custom properties relationship
in the package relationships part. Word reads properties through those relationships:
Word 16.115 offers recovery for all three defects in the referenced `docProps/core.xml`,
`docProps/app.xml`, and `docProps/custom.xml` specimens
([#409](https://github.com/jandira-tech/docx-validate/issues/409)). The relationship types
include the legacy core spelling and Strict extended/custom spellings. Targets resolve
from the package root (a leading `/` is allowed; external targets are ignored), and
part names compare by OPC equivalence, including the stored `_rels/.rels` spelling.
Missing or unreadable package relationships yield no properties targets; an unreadable
`_rels/.rels` already blocks Word-valid itself.

Without a properties relationship, these findings stay warnings on `docProps/*.xml`.
Word opens unreferenced malformed `docProps/core.xml`, `docProps/app.xml` and
`docProps/custom.xml` clean. The SuperDoc `annotations_import_2.docx`, whose
unreferenced `docProps/custom.xml` has an undeclared prefix, opens clean too and stays
Word-valid. A truncated core part stored as `props/meta.xml` and named by the
core-properties relationship makes Word offer recovery, and its well-formed control
opens clean, so the rule follows the relationship, not the name. Full, quick and named
runs share this policy; strict verdicts are unchanged.
External entities are not loaded or expanded; XSD imports resolve through allowed local schema roots.
Validation leaves source bytes unchanged unless you request `autoRepair`.

Content-type checks, XSD dispatch, and semantic diagnostics find the ZIP stream
ASCII case-insensitively and retain its stored name in findings. Stream lookup does
not decode percent escapes: `%5BContent_Types%5D.xml` is an ordinary part, not the
content-types stream ([#396](https://github.com/jandira-tech/docx-validate/issues/396)).
Of two stored spellings, `[Content_Types].xml` is read, otherwise the first in
code-unit order, where System.IO.Packaging reads the last in ZIP order; the second
stream is reported by `opc-equivalent-part-names` ([#351](https://github.com/jandira-tech/docx-validate/issues/351)).
Content-type coverage resolves an exact `Override` before an extension `Default`.
An image covered only by `Override` needs no matching `Default`. Part-name
comparison is ASCII case-insensitive and decodes unreserved percent escapes.
The comparison rule does not make those stored spellings legal: the
`opc-part-name-escaped-unreserved` check diagnoses them separately ([#408](https://github.com/jandira-tech/docx-validate/issues/408)).
Reserved escapes stay distinct (`%40` differs from `@`), while spaces and other
characters outside the URI path character set are encoded. Truly uncovered parts
still produce `ct-uncovered-part`.

Compatibility processing uses a validation-only clone. It selects the first
`mc:Choice` whose non-empty `Requires` names only understood namespaces,
otherwise the fallback, including nested alternatives; a `mc:Choice` without
a usable `Requires` is malformed (`mce-malformed`) and never selected
([#494](https://github.com/jandira-tech/docx-validate/issues/494)). Inherited
`Ignorable`, `ProcessContent`, and `MustUnderstand` directives apply.
Word rules read `effectiveXml()` with `WORD_UNDERSTOOD_NAMESPACES`; the
content-types stream and `.rels` parts add their own vocabulary, which the
processor applying those streams' rules understands
([#493](https://github.com/jandira-tech/docx-validate/issues/493)); checks of the
compatibility markup itself can read the raw DOM. Unselected branches do not
produce structural Word findings. The schema view uses its own understood
namespaces. Extension payloads in DrawingML/chart/presentation/spreadsheet `ext`
slots survive for lax schema validation; genuinely empty required extensions
still fail their schemas.

`word-document-root` resolves the main part through the package's `officeDocument`
relationship and requires an effective `document` root in the Transitional or
Strict WordprocessingML namespace. A foreign root makes Word refuse to open the
file; an ignored root makes Word silently drop all content. Both are blocking
errors in Word-valid and strict, and the check is included in the quick preset.

`word-main-document-part` reports a package that names no main part: `_rels/.rels`
is missing, or it parses and none of its direct internal `officeDocument` relationships
targets a part in the package (an empty ZIP, a package of only `[Content_Types].xml`
or only core properties). ECMA-376 requires the main document part. Word 16.115
offers to recover an empty archive and a package without `word/document.xml`
(docs/plans/evidence/word-probes-2026-10-24.md, session 3c), and a package whose
`officeDocument` relationship is removed (session 4), so the finding is a
blocking error in Word-valid and strict alike, and the check is included in the
quick preset; lenient runs no Word rule (#380). An `_rels/.rels`
that does not parse is left to `xml-syntax`, which Word-valid reports as a blocking
error, because Word offers to recover such a package (#388).

The same check also requires exactly one direct internal `officeDocument`
relationship. More than one is a blocking error in Word-valid and strict for
DOCX/DOCM, reported on `_rels/.rels` with every relationship `Id`. Word 16.115
offers recovery both for distinct well-formed main targets and for a repeated
target ([#515](https://github.com/jandira-tech/docx-validate/issues/515), Word probes
session 15, s15b/s15d). A main-typed part with no relationship still passes
Word-valid; strict retains its `rels-unreferenced` error. Transitional and Strict
relationship types count together; the mixed pair is inferred, not observed.
`TargetMode="External"` relationships do not count. Full, quick and named runs
use the same finding and severity; lenient runs no Word rule.

A package with no `[Content_Types].xml` stream at all is `ct-missing` — Word
cannot type a single part and 16.115 offers to recover it (same evidence file),
so that finding is Word-blocking too (#435).

`opc-relationships-structure` requires every relationships part to have a
`Relationships` root in the package-relationships namespace and only `Relationship`
elements in that namespace as direct children. It reports one finding for a wrong
root, or one per offending child, at the stored part name. Malformed XML is left
to `xml-syntax`. The check runs under Word-valid and strict, including quick and
named relationships selections; explicit lenient keeps its existing checks.
Word-valid blocks malformed package relationships (`_rels/.rels`, compared by OPC
part-name equivalence) and malformed relationships parts under `word/`, and warns on
other sidecars, which have no Word evidence yet. Strict reports errors on all of them. Main-part discovery also ignores nested
relationships ([#433](https://github.com/jandira-tech/docx-validate/issues/433),
[#434](https://github.com/jandira-tech/docx-validate/issues/434)).
Properties discovery, hyperlink and toolbar resolution, duplicate counts, and
the validators' semantic relationship checks likewise use direct children under
Word-valid and strict. Legacy lenient traversal and raw repair-inventory collection
remain unchanged.

Word 16.115 (build 16.115.26100424, macOS, 2026-10-24) offered recovery for all
three minimal package mutations: a foreign sibling with every relationship
unchanged (`43205d2359aaac36`), a foreign wrapper around the main relationship
(`176e6d8f133812b3`), and a root renamed to `Foo` (`1a5c659713cf7f9b`). These are
SHA-256 prefixes from the [recorded Word probes](docs/plans/evidence/word-probes-2026-10-24.md);
the unchanged minimal control (`0160dfc1db4c5ce1`) opened cleanly. A foreign element
appended to `word/_rels/document.xml.rels` got recovery too (`e7dae629d3e792a1`,
session 4), beside a clean control (`f67e18fa7a50dae1`).

`word-hyperlink-relationship` requires every `w:hyperlink` with an `r:id` in
a story part to name a relationship whose Type is exactly the Transitional or
the ISO Strict hyperlink relationship URI, with TargetMode `External`. Word
16.115 accepts exactly those two Types, byte for byte, and offers recovery for
`http://example.com/hyperlink`, `…/relationships/custom/hyperlink` and
`…/relationships/HYPERLINK` alike; a hyperlink relationship without
TargetMode still warns
([#614](https://github.com/jandira-tech/docx-validate/issues/614); session 28,
chunk s28a, in
[docs/plans/evidence/word-probes-2026-10-24.md](docs/plans/evidence/word-probes-2026-10-24.md)).
The finding is an error, and Word-blocking, under word-valid and strict;
lenient runs no Word rule.

`word-header-footer-relationship` requires every `w:headerReference` and
`w:footerReference` in a story root's `w:sectPr` to name a relationship whose
Type is exactly the Transitional or the ISO Strict header (respectively
footer) relationship URI. The story roots are
every main document and every glossary a main-document `glossaryDocument`
relationship reaches, and the reference resolves in that part's own
relationships part, in either relationship namespace. Word 16.115 offers
recovery for every refused shape probed: header→footer and footer→header
relationships in the main document's body `w:sectPr` (whether the target's
root was `w:hdr` or `w:ftr`, typed with the other story's content type, a
parameterized footer type, or styles), both mirrors from a reached glossary's
`w:docPartBody` `w:sectPr`, a `w:headerReference` naming the main
document's own styles relationship, and every other Type tried in place of
the two accepted URIs — `http://example.com/header`,
`…/relationships/custom/header`, `…/relationships/HEADER`, the URI padded
with spaces, wrapped in tab and newline character references, and the bare
`header`, the footer pair likewise (#613). A matching reference and
relationship opens clean even when the target's root element is the other
story's, so the check judges the relationship Type alone
([#606](https://github.com/jandira-tech/docx-validate/issues/606),
[#613](https://github.com/jandira-tech/docx-validate/issues/613); sessions 26
and 27, chunks s26a/s26b/s26c and s27a/s27b/s27c, in
`docs/plans/evidence/word-probes-2026-10-24.md`).
A reference whose `r:id` names no Relationship in the part's own relationships
part is this check's finding too — when that relationships part exists and
parses, as Word 16.115 offered recovery for a body `w:footerReference`
`r:id="rId9"` beside a kept `word/_rels/document.xml.rels` holding no `rId9`
([#619](https://github.com/jandira-tech/docx-validate/issues/619); session 32
in the same evidence file). In a full run, a part with no relationships part
keeps `rels-missing-sidecar` as its only finding for a non-empty `r:id`; the
legacy relationship walk that reports it runs only in full runs, so quick and
named runs report nothing for that shape.
Under word-valid for DOCX/DOCM, `rels-missing-sidecar` blocks when the part
is a story root — the main document or a reached glossary — under any `.xml`
name, even `word/header1.xml`: Word judges a story root by the
relationship that reaches it, not by its file name, and session 33 observed
recovery for both a renamed main document and a renamed reached glossary
whose `.rels` was gone beside their story reference, while the renamed
controls opened clean
([#625](https://github.com/jandira-tech/docx-validate/issues/625); session 33,
chunk s33b, in the same evidence file). Every other part keeps the basename
rule: a genuine `word/header1.xml` or `word/footer1.xml` bearing relationship
references with no sidecar keeps its word-valid warning — Word's own
regenerated output opens clean carrying that finding — and any other name
blocks as before. A story root stored without a `.xml` suffix
(`word/document.bin`) is not checked as XML, so it gets no
`rels-missing-sidecar` at all
([#620](https://github.com/jandira-tech/docx-validate/issues/620)).
An empty or missing id in both relationships namespaces is this check's
finding regardless of whether the sidecar exists or parses
([#623](https://github.com/jandira-tech/docx-validate/issues/623)). Session 33
observed Word 16.115 offer recovery for `r:id=""` on a header or footer
reference in the main document's body `w:sectPr` (s33a/s33c), and report an
error opening the file when either reference has no id (s33c/s33e). An id
in a third namespace counts as missing, even when it names a real footer
relationship (s33d). An ISO Strict relationships id naming a real footer
relationship on a Transitional reference opens clean and renders the footer
(s33c, #631 refuted). The check applies to the same placements and story
roots as before; session 33 observed these empty/missing shapes only in the
main document's body `w:sectPr`, while sessions 26/27 established the footing
for the other placements and reached glossaries. Set a non-empty id in
either relationships namespace naming a Relationship of the required Type,
or add one with that Id and Type. Strict also keeps the XSD error for a
missing `r:id`. The finding is an error under
strict and Word-blocking under word-valid for DOCX/DOCM; lenient runs no Word
rule. In a full word-valid or strict run neither `rels-id-mismatch` nor
`rels-id-missing` repeats an element this check reported — the run drops that
duplicate for a non-empty id, pairing the two findings by the element identity both checks
register on the issue (part path, namespace URI, local name, raw `r:id` and
the namespace the `r:id` was read in), never by message text — while `validateAllRelationshipIds()` called on its own
still reports every `rels-id-mismatch` and `rels-id-missing` (#612, #616,
#619). Empty/missing-id findings register no such identity, because the
legacy walk reads only truthy ids and reports neither shape.
A glossary no main-document relationship reaches keeps its
`rels-id-mismatch` or `rels-id-missing` under strict (Word does not read it,
so word-valid stays silent), and lenient keeps every one of them.
Full, quick and named Word-valid runs judge the findings they report alike
(#381).

`rels-missing-sidecar` reports a part that uses a non-empty `id`, `embed` or
`link` attribute in either the Transitional or ISO Strict relationships
namespace but has no `.rels` sidecar. Session 33 observed Word 16.115 offer
recovery for a Strict document's body `w:footerReference` with a Strict
`r:id="rId9"` and no sidecar (`s33a`), while the same Strict package without
the reference opened clean (`s33c`)
([#624](https://github.com/jandira-tech/docx-validate/issues/624);
`docs/plans/evidence/word-probes-2026-10-24.md`). A missing sidecar alone
does not trigger this finding. Counting Strict `embed` and `link` attributes
extends the existing Transitional rule on the same footing; those shapes
were not observed in Word in this session.

`word-markup-compatibility` reports processor errors for XML parts under `word/`.
Root validation, the processor, and the schema walk use these stable codes:

| Code                          | Meaning and profile behavior                                                                                                                                                              |
| ----------------------------- | ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `word-document-root`          | The main part has a foreign or missing effective root: Word refuses to open it or silently drops all content. Blocking in Word-valid and strict.                                          |
| `mce-malformed`               | Invalid compatibility structure, such as AlternateContent without a Choice. Blocking in Word-valid and strict.                                                                            |
| `mce-must-understand`         | A declared required namespace is not understood by the selected view. Blocking in Word-valid and strict; Word and schema capabilities can differ.                                         |
| `mce-unbound-prefix`          | A directive names an unbound prefix. Error in lenient/strict schema processing; warning in full Word-valid. A repeated root Ignorable finding is retained once as `ignorable-undeclared`. |
| `mce-invalid-token`           | An invalid ProcessContent QName. Error in lenient/strict schema processing; warning in full Word-valid.                                                                                   |
| `mce-root-ignored`            | An ignorable root disappears from the schema view. Informational for non-main parts; not emitted by the full Word-valid path.                                                             |
| `mce-main-root-ignored`       | The required main document root disappears from the schema view. Error in lenient/strict; the Word view reports `word-document-root` for a missing main root.                             |
| `mce-extension-not-ignorable` | Office extension attributes lack an Ignorable declaration and are removed from the schema view. Warning from schema preprocessing; not emitted by the full Word-valid path.               |

Under Word-valid and strict, full DOCX/DOCM runs retain root Ignorable findings
once as `ignorable-undeclared` even when staging fails. Named Word-valid runs
share that identity, severity and count. Descendant-only findings keep
`mce-unbound-prefix` ([#499](https://github.com/jandira-tech/docx-validate/issues/499)).

Strict bundles both Transitional and ISO Strict schema trees, with local Dublin
Core imports. Strict does not subtract errors shared with `original` or suppress
messages in `IGNORED_VALIDATION_ERRORS`; only lenient retains those XSD behaviors.
Strict preserves each schema error occurrence and available engine line. If
preprocessing changes the view, the message marks the line as referring to the
preprocessed view. Schema-load failures remain blocking `xsd-schema-load-failed`
issues; arbitrary unmapped custom XML is outside schema eligibility.

Optional real-Word verification is independent of portable checks:

```ts
const observed = await validate("./contract.docx", { verifyWord: true });
console.log(observed.metadata?.wordVerification);
```

```bash
docx-validate --verify-word --json contract.docx
```

The built-in verifier requires macOS, Microsoft Word, and Accessibility automation
permission for the host terminal. It opens a uniquely named temporary copy of the
exact input bytes, without repair. The matching document must remain open and
dialog-free for the default 3-second observation period. The default launch,
observation, and cleanup budget is 45 seconds, with up to 2 seconds reserved for
closing the document. Cleanup closes that copy without saving and never quits or
force-kills Word. Portable validation and Word verification both contribute to
`valid`; requesting verification does not make a portable failure pass.

`wordVerification` records `requested`, `satisfied`, `outcome`, `fileSha256`,
`platform`, `observedAt`, `observationMs`, `details`, and available `wordVersion`
and `wordBuild`. Outcomes are `clean-open`, `repair-warning`, `open-error`,
`crash`, `password-prompt`, `unknown-dialog`, `timeout`, or `unavailable`.
Only `clean-open` satisfies verification. Unavailable or inconclusive results
make an explicitly requested run invalid even when `issues` contains no errors.
An unpacked directory has no packed bytes to verify and returns `unavailable`;
other platforms or missing Word also return `unavailable`. `crash` requires a
Microsoft Word crash report written after this run's open; a Word process that
vanishes without one (an external kill or quit) is `unavailable` (#305). The Word
version and availability calls run inside the same budget as the rest of the run,
so a hanging call ends as `timeout` instead of outliving it (#320). A verifier
that throws is reported as `unavailable` with its message; `validate` does not
reject because of it (#333).

Before opening, the verifier observes Word once. If a dialog is already showing,
it returns `unavailable` without opening or closing anything: that dialog cannot
be attributed to this file. Verification needs exclusive use of Word: runs must
not overlap and Word must not be in use, because a dialog that appears after the
pre-open observation is attributed to this file (#331). macOS cleanup dismisses
the Word dialog showing at close, preferring Cancel, No, then OK; under exclusive
use that is the dialog this run's open raised. When the open itself overran its
budget, cleanup first lets it settle within the rest of the budget, so a late open
never lands after the close (#328). It does not relaunch Word after a crash.
The source implementation (`src/lib/word-verifier.ts`) supports injected transport
and clock dependencies, plus a `wallNow` source: `observedAt` and the crash-report
open instant are wall time, independent of the deadline clock (#329); public validation options accept a `wordVerifier` callback
for deterministic tests or a custom verifier.

### Compose individual checks with the validator classes

```ts
import { DOCXSchemaValidator, defaultSchemasDir } from "docx-validate";

const v = new DOCXSchemaValidator({
    unpackedDir: "./unpacked",
    schemasDir: defaultSchemasDir(), // override if you bundle your own XSDs
    profile: "lenient",
});

const xsdResult = await v.validateAgainstXsd();
const idResult = await v.validateUniqueIds();
const relResult = await v.validateAllRelationshipIds();
```

`PPTXSchemaValidator`, `BaseSchemaValidator`, and `RedliningValidator` follow
the same shape — see `src/scripts/office/validators/` for the full method
list.

### Side helpers

```ts
import {
    pack, // repack an unpacked dir into .docx/.pptx/.xlsx
    unpack, // unzip + pretty-print + optional run-merging
    addComment, // append a w:comment to an unpacked DOCX
    mergeRuns, // collapse adjacent w:r runs with identical formatting
    simplifyRedlines, // collapse adjacent same-author tracked changes
} from "docx-validate";

// e.g. unpack → mutate → repack:
await unpack("./contract.docx", "./unpacked");
// (your edits go here)
await pack("./unpacked", "./contract.modified.docx");
```

### Drive the CLIs programmatically

If you want the CLI behaviour (commander parsing, exit codes, the same status
messages) without spawning a subprocess:

```ts
import { runValidateFromArgv, buildValidateCommand } from "docx-validate";

const exit = await runValidateFromArgv(["./contract.docx", "--profile", "strict"]);
process.exit(exit);
```

Each script ships its own `build*Command` / `run*FromArgv` pair:
`runValidateFromArgv`, `runPackFromArgv`, `runUnpackFromArgv`,
`runCommentFromArgv` (and corresponding `build*Command` factories that
return the underlying commander `Command`).

### Not in the package surface: LibreOffice helpers

`acceptChanges()` (LibreOffice macro for accepting tracked changes) and the
underlying `runSoffice()` / `getSofficeEnv()` / `ensureShim()` helpers live
in `src/scripts/accept-changes.ts` and `src/scripts/office/soffice.ts` but
are intentionally **not** re-exported from the package barrel. They use an
`LD_PRELOAD` shim to make `soffice` boot in sandboxed Linux VMs, and that
pattern triggers supply-chain malware heuristics in automated scanners (it
genuinely matches the shape, even though the use is benign — see
[SECURITY.md](./SECURITY.md)).

To use those helpers, run the scripts directly from a checkout
(`pnpm exec tsx src/scripts/accept-changes.ts <input> <output>`) rather than
importing from the published package.

### Result shape — `ValidationResult`

```ts
import type { ValidationMetadata } from "docx-validate";

interface ValidationIssue {
    severity: "error" | "warning" | "info";
    message: string;
    path?: string; // file path inside the unpacked dir, when applicable
    code?: string; // stable string ID; safe to switch on
    line?: number; // 1-based, when available
    column?: number; // 1-based, when available
}

interface ValidationResult {
    valid: boolean; // requested checks and any requested Word verification passed
    issues: ValidationIssue[];
    metadata?: ValidationMetadata;
}
```

`ValidationMetadata` has optional `profile`, `executedChecks`, `omittedChecks`,
`schemaCompleted`, and `wordVerification` fields. `executedChecks` reports the
registry ids a complete run executed (#319): lenient lists `opc-equivalent-part-names`,
`opc-part-name-escaped-unreserved`, `opc-part-name-escaped-slash`, and
`opc-part-name-word-refused` first, then `xml-syntax` and the rest; Word-valid
and strict add the ZIP-envelope checks for a packed file and every Word rule whose
prerequisites passed. A partial run lists the selected checks that ran, so an
archive-only check is not listed for an unpacked directory. A missing target (`ENOENT` or `ENOTDIR`) throws
`Error: <target> does not exist`; other access failures throw `Error: Cannot access <target>: <system error>`.
The CLI prints that error on one line and exits 1 ([#574](https://github.com/jandira-tech/docx-validate/issues/574)). An archive that does
not open fails with `package-open-failed` in every mode, and only its envelope
checks are listed (#332). A packed file that cannot be read at all — a mode-000
file `fs.stat` still answers for — fails the same way, with no check listed
([#422](https://github.com/jandira-tech/docx-validate/issues/422)). Its message says `Cannot read` followed by the system error;
the encrypted, corrupted or truncated archive explanation is reserved for JSZip refusals
([#547](https://github.com/jandira-tech/docx-validate/issues/547)). One that opens but whose package view cannot be read —
a stored item name that places no part — judged on the central-directory name,
never the local header's or the unicode-path field's (#445) — or an unreadable
entry — fails with
`package-extract-failed`, as the full run's extraction reports it (#413). Directory inputs apply the same stored-name refusal before either full or partial validation: literal backslash names with empty or dot-only segments are refused with the same code and message (#424).
Direct `DOCXSchemaValidator.validate()` and `PPTXSchemaValidator.validate()` calls also return an error-severity
`package-extract-failed` for such refused directory names, with the same message as `validate()`; programming errors still propagate
([#548](https://github.com/jandira-tech/docx-validate/issues/548)). They also apply the same directory walk refusal
for broken symbolic links, non-regular files, unlistable directories and symbolic-link loops under every profile,
before running any part-content check ([#557](https://github.com/jandira-tech/docx-validate/issues/557)). For either refusal,
direct validators include the same stored-name findings, in the same order and at the same profile-dependent severities,
as `validate()` for the same path string, and list the stored-name checks they ran in `executedChecks`, as `validate()`'s
`metadata.executedChecks` does ([#572](https://github.com/jandira-tech/docx-validate/issues/572)). Its stored
entry names and archive bytes are still read. Selected `opc-equivalent-part-names`,
`opc-part-name-escaped-unreserved`, `opc-part-name-escaped-slash`, and
`opc-part-name-word-refused` checks and
selected envelope checks (such as `zip-bytes-after-eocd`) run and report what the
full run reports. Every other selected check needs the package's parts and is
listed in `omittedChecks`
([#420](https://github.com/jandira-tech/docx-validate/issues/420)). It is not a
per-method execution trace. In the source scheduler, use `runChecks().outcomes`
for individual
`passed` / `failed` / `skipped` / `unavailable` statuses. `omittedChecks` names ids
left out by selection and, when the package view cannot be read, selected ids that
could not run. `schemaCompleted` records completion of the eligible schema
walk without schema-load failures; it is separate from whether schema findings
made the document invalid. `wordVerification` is present only when requested.

Full validation accepts unpacked directories, including POSIX files whose stored
names contain literal backslashes (for example `word\document.xml`). Such inputs
are validated in a temporary copy under package part names (`word/document.xml`),
using the same duplicate selection and extraction guards as packed inputs.
Directories whose tree holds a symbolic link — a stored entry or a directory
segment — are staged the same way: the stored bytes are read through the link
into the package view, while the write-back never follows it
([#425](https://github.com/jandira-tech/docx-validate/issues/425)). A link
whose target resolves outside the input root is not followed at all: the walk
refuses it at the link's stored name (`a symbolic link outside the input
directory`), so foreign files never become package parts, content outside the
root never decides the verdict, and a link to an ancestor of the root is named
at the link itself rather than under a path inside it
([#576](https://github.com/jandira-tech/docx-validate/issues/576)). A link that
resolves inside the root keeps the behaviour above: the loop detection of
[#546](https://github.com/jandira-tech/docx-validate/issues/546) and the
write-back refusal of
[#544](https://github.com/jandira-tech/docx-validate/issues/544). The bound
fails closed: an input root `realpath(3)` cannot resolve is refused as an
unreadable directory, and a link it cannot resolve is refused (`a symbolic
link the walk cannot resolve`) rather than followed. A stored
destination is refused when any path segment between the input root and that
destination is a symbolic link. When several stored paths name one physical
file and any of them reaches it through a link, the whole group is refused, so
the link's target stays untouched and one physical file is never written twice
([#544](https://github.com/jandira-tech/docx-validate/issues/544)). A stored
regular file whose link count is greater than one is also refused, because a
hard-link alias may name a file outside the input directory
([#545](https://github.com/jandira-tech/docx-validate/issues/545)).
Directories with ordinary forward-slash names and no links are validated
directly. Validation
leaves source files unchanged unless `autoRepair: true` / `--auto-repair` is used.
For a staged directory, repairs write changed bytes back to each original stored
file name and delete originals of parts removed by a repair; untouched files and
dropped duplicate entries remain byte-identical. No forward-slash replacement
files are added to the input directory. Quick and named-check runs continue to
read the directory's package view directly ([#374](https://github.com/jandira-tech/docx-validate/issues/374)).

A stored file that cannot be read fails the run with `package-extract-failed`,
naming the stored file — the package view cannot be staged. This includes a
stored entry that can never yield part bytes at all: a broken symbolic link,
or a non-regular file such as a FIFO, which is reported by its stored name
instead of silently vanishing from the package
([#426](https://github.com/jandira-tech/docx-validate/issues/426)); a FIFO is
never opened, so it cannot hang the run. A subdirectory the walk cannot list
fails the same way, naming the directory — every stored file under it would
otherwise vanish with it, leaving the run to judge the parts it can still see
([#543](https://github.com/jandira-tech/docx-validate/issues/543); the input
directory itself is named `.`). A directory is no stored file entry, so it
contributes no name to the package view: it is named only in the
`package-extract-failed` detail, and the stored-name checks never judge a
name no package holds
([#573](https://github.com/jandira-tech/docx-validate/issues/573)). A symlink loop — a directory link whose target
is already on the path into it, or a link chain that never resolves — is
reported as `a symbolic link loop` at the link's stored name (`word/self`),
not as a broken link under a synthetic multi-level path
([#546](https://github.com/jandira-tech/docx-validate/issues/546)). A link out
of the input root fails the same way, naming the link's own stored name
([#576](https://github.com/jandira-tech/docx-validate/issues/576)). When several
entries are unreadable, the message names them one clause each, in code-unit
order — the order the file list uses — so the same input yields the same
message on every host, whatever the host's locale collation is
([#575](https://github.com/jandira-tech/docx-validate/issues/575)). Full,
quick and named-check runs
report the same failure. (Under ordinary
forward-slash names the direct path reads parts lazily, so the same file yields
`check-crashed` warnings instead.) When a write-back cannot reach its stored
file — a read-only file, one removed mid-run, or a stored file that is a
symbolic link, whose root-relative path contains a symbolic-link segment, has
more than one hard link, or resolves outside the input directory, whose target
the write-back never touches — the run reports a
`repair-write-back-failed` error naming the part and keeps writing back the
remaining repairs; the stored file is left untouched. That error is
profile-independent: a repair the caller asked for and did not get is an
operational failure, not a property of the document, so no profile downgrades
it ([#428](https://github.com/jandira-tech/docx-validate/issues/428)). `repairs` counts repairs
applied to the package view, so a repair whose write-back failed still counts
([#412](https://github.com/jandira-tech/docx-validate/issues/412)).

These write-back checks are not race-resistant. They inspect the stored path
and inode, then later write by path; a symbolic or hard link swapped in
between the check and the write can still be followed. The same check-then-use
limit is documented by `containRelationshipTarget`; descriptor-based
no-follow writes would be required to close that race
([#545](https://github.com/jandira-tech/docx-validate/issues/545)).

Distinct normalized part names that collide under ASCII case folding and NFC
normalization now refuse the package view **on every host**, before extraction
(#441). The same guard includes folder prefixes (`word/a.xml` beside
`WORD/b.xml`), so a staging filesystem cannot silently rename their parts.
Full, quick and named runs report `package-extract-failed` with the same detail,
run only applicable stored-name and envelope checks, and omit part checks. This
is a portable staging policy, independent of whether the host filesystem is
case-sensitive. It does not turn every OPC-equivalent spelling into a staging
refusal: identical stored names and slash/backslash aliases that normalize to
one part retain their existing entry selection; escaped spellings retain their
existing OPC diagnosis.

A later extraction failure on an otherwise readable view still merges the
archive-backed findings under the selected profile (#454). For example, the
host may refuse an overlong filename. Already-run stored-name and envelope
checks are not repeated; merged checks appear in `executedChecks`. This keeps
partial errors a subset of full errors (#381). Regression coverage pins case,
NFC/NFD and NUL-bearing names for quick, XML-only and package-only selections
under Word-valid and lenient (#423). A partial run can still pass when full
extraction fails on a NUL-bearing name: omitted work can miss a defect.

`ValidateRunResult` extends this with `suffix` (e.g. `".docx"`), `repairs`
(number of issues auto-repaired when `autoRepair: true`), and
`xsdWalk: { started, parts, eligible }`. `Validate.run()` and `jubarte.validate()`
return the shared `ValidationResult`, including metadata, without those extra fields.

## Development

```bash
pnpm install
pnpm test --coverage          # vitest with coverage
pnpm run test:coverage-gate   # full coverage suite + per-module thresholds
pnpm run check     # vite-plus check
pnpm run build     # vp pack: library, CLI, declarations, schemas
```

Coverage-enabled Vitest runs print `Coverage: N% lines, M% branches`, including
failed runs, and write text, JSON-summary, HTML, and LCOV reports. The summary is
printed only when coverage is enabled. `scripts/coverage-gate.ts` reads
`coverage/coverage-summary.json` and enforces 80% lines / 70% branches for the
listed new modules, and 90% / 85% for mode policy, scheduling, registry, and result
aggregation. A gate with no matching coverage fails; run the full suite before
using it. CI runs the same gate after its coverage test step.

`tests/fixtures-all-{strict,lenient}.test.ts` pin validator outcomes in
`tests/fixtures-all.manifest.json`. `tests/fixtures-all-word-valid.test.ts` instead
compares each attributable DOCX fixture against Word's hash-bound observation:
`clean-open` means valid, observed rejection means invalid. Unknown, timed-out,
crashed, and stale observations, and any outcome spelling the accounting does not
know, are excluded (#316); floors of 358 clean opens and 67 rejections
prevent silent loss of evidence. The 24 jubarte 0.13.0 specimens live separately
under `tests/lane/` with their own integration suite.

Manifest `word` observations use `{ outcome, sha256, observedAt?, wordVersion? }`.
The rejections are `open-error`, `unreadable-content-warning`, `password-required`,
and `unknown-dialog`. In source, `toManifestOutcome` translates a verifier outcome
to these spellings (`repair-warning` → `unreadable-content-warning`,
`password-prompt` → `password-required`, `unavailable` → `unknown`).
A mismatching hash makes the observation `unknown`; legacy strings remain readable
but cannot be checked for staleness. `mergeWordObservation` preserves superseded
observations in `history` and does not let a timeout/unknown re-probe displace
attributable evidence for unchanged bytes. `pnpm run test:fixtures:refresh` uses
`scripts/update-manifest.ts`, which preserves this provenance. Review expectation
changes individually; current validator output is not independent Word evidence.

Run the installed-package smoke separately from the regular suite. It builds,
packs, installs in a fresh consumer, validates with bundled Transitional/ISO Strict
schemas, checks packaged CLI help, and confirms the library starts no CLI.
The real-Word integration lane is also separate and opens the actual application:

```bash
PACKAGE_SMOKE=1 pnpm exec vitest run tests/package-artifact.integration.test.ts --coverage
WORD_VERIFY=1 pnpm exec vitest run tests/word-verifier.word.test.ts --coverage
```

The Word lane requires macOS with Word installed and host-terminal Accessibility
permission. Regular portable runs do not perform fresh Word opens.

## Acknowledgments

Many of the test fixtures under `tests/fixtures/external/` are borrowed from
other open-source OOXML projects' own test corpora. Thanks to the maintainers
of these projects for keeping their fixtures public — they made it possible
to validate against real-world malformed and edge-case documents instead of
only synthetic ones:

- [apache/poi](https://github.com/apache/poi) — Apache-2.0
- [dotnet/Open-XML-SDK](https://github.com/dotnet/Open-XML-SDK) — MIT
- [plutext/docx4j](https://github.com/plutext/docx4j) — Apache-2.0
- [guigrpa/docx-templates](https://github.com/guigrpa/docx-templates) — MIT
- [mwilliamson/mammoth.js](https://github.com/mwilliamson/mammoth.js) — BSD-2-Clause

Per-vendor `NOTICE` / `LICENSE` files live alongside the fixtures in
`tests/fixtures/external/<vendor>/`, and `tests/fixtures/external/README.md`
documents file-level provenance.

## Validators (etc.) for the neurotic developer

A place for "why isn't this more complete?" questions. If you came here
because something looked suspiciously trim, this is where it gets
explained.

### Why doesn't `NS` in `lib/types.ts` cover every namespace declared in the bundled XSDs?

`NS` is a **runtime lookup table for code that needs namespace URIs** —
`getElementsByTagNameNSAll(root, NS.W, "p")`,
`el.getAttributeNS(NS.W, "author")`, the `XPATH_NS` prefix map. It is not
a registry of every namespace that exists in OOXML.

The bundled XSDs declare more namespace URIs than `NS` exposes. Schema-dispatch
tables and compatibility-capability sets also describe namespaces used by the
validators; `NS` remains a focused helper lookup. XSD validation is independent:
the wasm
engine reads the `.xsd` files directly, so `NS` membership has no effect on
what gets validated.

When a future validator needs `c:chart` (DrawingML chart) or `v:shape`
(VML), add the constant to `NS` next to the validator that uses it —
don't pre-populate speculatively. The schemas under
`src/scripts/office/schemas/` remain the source of truth for _element
definitions_; `NS` is the source of truth for _namespace URIs the TS
code references at runtime_.

## Releases and trusted publishing

`pnpm run release:check` runs the prerelease preflight: type-check, type-aware lint, formatting, spelling, coverage tests, build, bundle inspection, and an npm pack dry run. It verifies that the package contains its ESM entry, declarations, and bundled Word schema. It publishes nothing.

Update `package.json` and `CHANGELOG.md`, run the preflight, commit, and tag the tested commit. Create a GitHub release for that tag; mark versions such as `0.53.0-rc.1` as prereleases. Publishing uses [npm trusted publishing](https://docs.npmjs.com/trusted-publishers/) from `publish.yml`, with the existing `npm-publish` environment, `id-token: write`, a GitHub-hosted Node 24 runner, and no dependency caches or npm publish token.

The npm trusted publisher must authorize organization `jandira-tech`, repository `docx-validate`, workflow filename `publish.yml`, environment `npm-publish`, and direct `npm publish`. Package ownership and the trusted-publisher configuration are separate settings.

The workflow checks out the exact release tag, verifies its package version and GitHub prerelease status, runs the preflight again, and publishes with provenance. Prerelease versions go to `next`; stable versions go to `latest`.

To retry publishing an existing tag through the same GitHub workflow:

```sh
pnpm run release:publish v0.53.0-rc.1
```

This dispatches Actions; it does not use local npm credentials. Monitor the [publish workflow](https://github.com/jandira-tech/docx-validate/actions/workflows/publish.yml) for the registry result. An npm version already published cannot be overwritten.

## License

Apache-2.0 — see [LICENSE](./LICENSE).

## Desired Conversion Methodology

**Desired Conversion Methodology** means fully parsing the original document into a semantic, format-independent representation — extracting its content and the meaning of that content (including its structure, formatting, and relationships, and every other discrete constituent component of the document (each, an "**Element**")), and not its byte form, such that the document can be faithfully reconstructed into any target format.

Where an Element cannot be parsed semantically by available means — including, by default, DrawingML (`<w:drawing>`), VML pictures (`<w:pict>`), and embedded OLE objects (`<w:object>`) (each, a "**Non-Parseable Element**") — the Desired Conversion Methodology nonetheless applies in full to the entire document and to every Element enclosing that Non-Parseable Element; only the single, irreducible Element that resists semantic parsing may be preserved verbatim, and only after custom parsing techniques have been exhausted to reduce that residue to the smallest possible extent.

For the avoidance of doubt, the Elements enumerated above are not the sole Non-Parseable Elements; every time an Element is classified as a Non-Parseable Element, such classification must (i) emit a warning during code execution (at runtime), and (ii) be documented in `AGENTS.md`, `README.md` and `CLAUDE.md`.
