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
every package format; `word-content-type-invalid` and
`word-part-content-type-generic` apply to DOCX/DOCM only.

`zip-bytes-before-archive` applies to DOCX/DOCM, PPTX, and XLSX package
selections, including the `package` quick group, and reads the archive bytes
before extraction. It reports the bytes in front of the archive JSZip loads —
the rebase offset plus the lowest stored local-header offset, so a prepended
ZIP, junk with relative offsets, and junk with the absolute offsets a
self-extractor writes all measure the same prefix — as one blocking error
under strict and Word-valid. Word 16.115 offers to recover every prefixed
package probed ([#464](https://github.com/jandira-tech/docx-validate/issues/464);
probe session 14 in
[docs/plans/evidence/word-probes-2026-10-24.md](docs/plans/evidence/word-probes-2026-10-24.md)).
Lenient runs no ZIP-envelope check. A concatenated `[zip][docx]` is reported
once, by this check: the earlier archive lies wholly in front of the archive
JSZip loads, so `zip-bytes-after-eocd` no longer counts the live archive as
undeclared trailing bytes. Trailing data behind a prefixed archive is still
`zip-bytes-after-eocd`'s finding, measured from the loaded archive's own
record.

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
and the Info-ZIP unicode-path extra field (0x7075) (#445). If re-keying is
unavailable (for example, ZIP64), JSZip still prefers a valid central Unicode
Path field. `unpack` is the exception: it extracts under raw JSZip's own names,
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

`ct-override-partname-word-refused` applies the same escaped-mark and high-octet
rules to every Override PartName, even when no part is stored under that name.
It reports one finding per declaration at the stored content-types stream,
quoting the PartName. Literal non-ASCII PartNames already fail
`ct-override-partname-invalid` and are not reported twice. A spelling containing
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
| `content-types`           | `word-content-type-invalid`, `ct-duplicate-declaration`, `ct-override-partname-invalid`, `ct-override-partname-word-refused`, `ct-markup-compatibility`, `word-part-content-type-generic`                                                                                                                                                           |
| `relationships`           | `word-duplicate-theme-relationship`, `word-duplicate-font-table-relationship`, `word-duplicate-web-settings-relationship`, `word-duplicate-settings-relationship`, `word-hyperlink-relationship`, `word-main-document-part`, `opc-relationships-structure`                                                                                          |
| `tracked-changes`         | `word-textbox-deleted-text`, `word-move-from-deleted-text`                                                                                                                                                                                                                                                                                          |
| `settings`                | `word-custom-property-filetime-empty`, `word-attached-toolbar-utf8`                                                                                                                                                                                                                                                                                 |
| `xml`                     | `word-markup-compatibility`, `word-document-root`, `word-choice-requires-undeclared`                                                                                                                                                                                                                                                                |
| `comments`                | `comment-markers`, `comment-threading`                                                                                                                                                                                                                                                                                                              |
| `schema`                  | `word-unknown-element`, `word-link-without-namespace`                                                                                                                                                                                                                                                                                               |

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
a valid OPC part name (ECMA-376 Part 2 §6.2.2.2): a name that does not start
with `/`, holds an empty, a `.`/`..` or a dot-only (`...`) segment, ends with
`/` or `.`, has a folder segment ending with `.`, carries a `?`, a `#` or a `\`, percent-encodes a
separator (OPC [M1.7]) or an unreserved character (OPC [M1.8]), or holds a
character outside the RFC 3986 `pchar` set. Such a PartName names no part, so
the declaration is malformed whatever its content type. Resolution is
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
uses is not reported. Microsoft Word 16.115 does not process Markup
Compatibility in this stream: it offers recovery for every shape probed — an
`mc:Ignorable` attribute, an ignored foreign element, and an AlternateContent
in either orientation — while the declared-only control opens clean
([#459](https://github.com/jandira-tech/docx-validate/issues/459); probe
session 14, `docs/plans/evidence/word-probes-2026-10-24.md`). The check reads
the raw stored stream, while `ct-duplicate-declaration`,
`ct-override-partname-invalid`, `word-content-type-invalid` and content-type
resolution all judge the stream after Markup Compatibility processing, as OPC
[M1.19] requires — so an `Override` in an ignored `mc:Choice` no longer draws
a `ct-override-partname-invalid` finding, and a bad `mc:Fallback` still does.
The finding is an error, and Word-blocking, under word-valid for DOCX/DOCM and
under strict (whose schema walk already refuses the same stream); lenient and
the unobserved PPTX/XLSX warn. Full, quick and named runs agree.

`word-part-content-type-generic` reports a part Word reads whose content type
resolves only through the generic `<Default Extension="xml"
ContentType="application/xml"/>` — no Override of its own. The reached parts
are the main document (the package `officeDocument` relationship's target) and
the parts the main document's relationships of type styles, settings,
fontTable or webSettings name, wherever they are stored. Resolution follows
ECMA-376 Part 2 §10.1.2.4 through the shared table: an exact `Override` wins,
so a part with a specific Override — right or wrong — is out of scope, and
numbering and customXml stay silent until a probe observes them. One finding
is reported per part, at the part's path, naming the relationship type and
the content type the part needs. Word 16.115 offers to recover the package
for all five parts probed typed only by the generic Default — the main
document and styles with their Override removed, and settings, fontTable and
webSettings added with a relationship but no Override — and opens every
proper-Override control clean ([#432](https://github.com/jandira-tech/docx-validate/issues/432),
session 8a in `docs/plans/evidence/word-probes-2026-10-24.md`). The finding
is an error, and Word-blocking, under word-valid for DOCX/DOCM, and an error
under strict, where the main document is not reported twice because
`ct-undeclared-part` already rejects it; lenient reports a warning. Full,
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
([#390](https://github.com/jandira-tech/docx-validate/issues/390)). Word-valid blocks on a malformed part under `word/`, on
the content-types stream, or on the package relationships part `_rels/.rels`;
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
understood `mc:Choice`, otherwise the fallback, including nested alternatives;
inherited `Ignorable`, `ProcessContent`, and `MustUnderstand` directives apply.
Word rules read `effectiveXml()` with `WORD_UNDERSTOOD_NAMESPACES`; checks of the
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
archive-only check is not listed for an unpacked directory. An archive that does
not open fails with `package-open-failed` in every mode, and only its envelope
checks are listed (#332). One that opens but whose package view cannot be read —
a stored item name that places no part — judged on the central-directory name,
never the local header's or the unicode-path field's (#445) — or an unreadable
entry — fails with
`package-extract-failed`, as the full run's extraction reports it (#413). Its stored
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
Directories with ordinary forward-slash names are validated directly. Validation
leaves source files unchanged unless `autoRepair: true` / `--auto-repair` is used.
For a staged directory, repairs write changed bytes back to each original stored
file name and delete originals of parts removed by a repair; untouched files and
dropped duplicate entries remain byte-identical. No forward-slash replacement
files are added to the input directory. Quick and named-check runs continue to
read the directory's package view directly ([#374](https://github.com/jandira-tech/docx-validate/issues/374)).

A stored file that cannot be read fails the run with `package-extract-failed`,
naming the stored file — the package view cannot be staged. (Under ordinary
forward-slash names the direct path reads parts lazily, so the same file yields
`check-crashed` warnings instead.) When a write-back cannot reach its stored
file — a read-only file, or one removed mid-run — the run reports a
`repair-write-back-failed` error naming the part and keeps writing back the
remaining repairs; the stored file is left untouched. `repairs` counts repairs
applied to the package view, so a repair whose write-back failed still counts
([#412](https://github.com/jandira-tech/docx-validate/issues/412)).

A packed file whose extraction fails — two stored names that collide on the host
filesystem, such as `[Content_Types].xml` beside `[CONTENT_TYPES].XML` — still
reports what the archive holds. The full run then runs the checks a partial run
of that package runs on the in-memory archive context, under the profile, and
merges their findings into the `package-extract-failed` failure; the checks it
had already run — the ZIP-envelope checks and the stored-name checks — are not
run twice, and the merged checks are listed in `executedChecks`. The verdict
does not change: the package cannot be staged, so it stays invalid. A quick,
group or named run of the same package therefore no longer reports an error the
full run drops
([#454](https://github.com/jandira-tech/docx-validate/issues/454)).

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
