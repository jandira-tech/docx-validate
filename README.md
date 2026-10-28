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
the XSD walk. It also holds every other part `[Content_Types].xml` declares
with an XML type (`application/xml`, `text/xml`, or an `application/…+xml`
type), whatever its name: Word reads a part by its relationship and content
type, so a main document stored as `word/document.bin` under the main
document's Override is checked as `word/document.xml` would be, and Word 16.115
refuses it truncated (#620, session 32). Image types such as `image/svg+xml`
are not XML parts here. DOCX relationships validation passes use the same OPC-equivalent
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
The walk sends a part to the schema its identity maps to. The glossary
document's parts (`word/glossary/*.xml`) use `wml.xsd`, and the DrawingML
diagram parts (`diagrams/dataN.xml`, `layoutN.xml`, `quickStyleN.xml`,
`colorsN.xml` under `word/`, `ppt/` or `xl/`) use `dml-diagram.xsd` (#628). A
part with no mapped schema, such as `customXml/item1.xml` or a diagram's
`drawingN.xml` cache, is not sent: `xsdWalk.parts` lists exactly the parts
that were.

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

Under Word-valid and strict, the legacy relationship walks read direct
`Relationship` children after markup compatibility, using the same relationships
namespace capabilities as the named checks. Duplicate IDs split between a direct
child and a selected Fallback are counted together; unselected Choice relationships
do not supply IDs, types or referenced targets. Word 16.115 offered recovery for
the package duplicate in session 34, `s34a`; the analogous split in a story-root
sidecar is unobserved in Word. The selected footer and main relationships that
opened clean in session 33, `s33d`/`s33e`, no longer draw false `rels-id-missing`
or `rels-unreferenced` findings ([#635](https://github.com/jandira-tech/docx-validate/issues/635),
[Word probe evidence](docs/plans/evidence/word-probes-2026-10-24.md)). Strict still
reports stored relationships markup compatibility and schema defects. Lenient
keeps its raw traversal and existing findings.

Under Word-valid and strict, the file-reference walk behind `rels-broken` and
`rels-unreferenced` also skips `TargetMode="External"` relationships: an External
target names a resource outside the package, never a part (ECMA-376 Part 2 §9.3),
so a `file:///…`, UNC or `ftp://` target — such as the `attachedTemplate`
relationship Word writes to a local template — is no longer resolved as a package
path and reported as a broken reference. Word 16.115 opened External hyperlinks
to `file:///missing.rels`, `file:///C:/docs/_rels/missing.xml` and
`../customXml/missing.xml` clean, as direct children and inside a selected `w14`
Choice, and one a `w:hyperlink` names; these are the targets Word-valid's
`rels-broken` still blocked, so they no longer block (Word probes session 40,
`s40a`/`s40b`). Lenient keeps its historical scan, which skips only `http…` and
`mailto:` targets ([#665](https://github.com/jandira-tech/docx-validate/issues/665)).
An internal broken target blocks Word-valid when the part it resolves to is a
relationships part or a `customXml` part outside `word/`, whatever the Target's
spelling: Word 16.115 refused an absent package-root `customXml/item1.xml` spelled
`../customXml/item1.xml`, `/customXml/item1.xml`, `..\customXml\item1.xml`,
`../CUSTOMXML/%69tem1.xml` and `../../customXml/item1.xml` (an extra `..` stays at
the package root), and opened clean an absent image spelled `media/_rels/x.png`,
`notes.rels` or `word/../customXml/x.png` and a reached glossary's absent
`../customXml/item1.xml` (Word probes session 54, `s54a`/`s54b`;
[#734](https://github.com/jandira-tech/docx-validate/issues/734)).

The processed view is DOCX/DOCM-only: every observation behind it is Word
opening `.docx` files, and PowerPoint was never shown markup compatibility in a
relationships part, so PPTX keeps the stored view under every profile
([#643](https://github.com/jandira-tech/docx-validate/issues/643)). A DOCX run
builds each relationships part's view once, through the run context's cached
`effectiveXml` ([#642](https://github.com/jandira-tech/docx-validate/issues/642)).

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
`xsdValidator`, `verifyWord`, `wordVerifier`, and `deadlineMs`. Set `suffix` on byte APIs for
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
  `--strict` also rejects an explicit profile other than `strict`. A partial run
  never repairs, so `--auto-repair` / `autoRepair: true` with `--quick` or
  `--checks` is refused rather than ignored (#652). API conflicts throw
  `ModeConflictError`; CLI configuration errors exit 1.
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
- `--deadline-ms <ms>` / `{ deadlineMs }` bounds the whole run, counted from the
  call; the byte APIs count their temporary write too (#630). The bound is
  cooperative: each step asks the budget before it starts — the package stages,
  every named check, every rule of the DOCX or PPTX validator, every part of the
  schema walk, the redlining validation, and Word verification — and a step
  already running finishes, so a run overruns by at most its longest single step.
  With a deadline the DOCX validator's rules run one after another rather than
  all at once, so the budget can stop one between two others. When the budget
  stopped any step, the run is invalid and reports one error-severity
  `validation-deadline` finding naming what did not run: a run cut short is never
  a pass, whatever the steps that ran found (timeout > engine error > invalid >
  valid). The profile's severity policy still applies to what ran.
  `metadata.deadline` reports `{ budgetMs, notRun }`. A schema walk the budget
  stopped leaves `schemaCompleted: false` and counts the parts it never reached
  in `xsdWalk.notExamined`. Word verification gets what is left of the budget, at
  most its own 45 seconds; once the budget is spent it does not run and is
  `unavailable`. The budget is a finite number of milliseconds, 0 or more
  (`RangeError` otherwise; the CLI exits 1). It cannot be combined with
  `--auto-repair`: a repair writes, and a deadline could stop it part-way.

```bash
docx-validate --list-checks
docx-validate --quick contract.docx
docx-validate --checks tables,word-run-misplaced --json contract.docx
docx-validate --strict contract.docx
docx-validate --deadline-ms 30000 --json contract.docx
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
`opc-part-name-escaped-slash`, `opc-part-name-word-refused`, and
`opc-item-name-refused`; `xml-syntax` is the shared XML prerequisite. All `word-*` checks apply to DOCX/DOCM only.
`ct-duplicate-declaration`, `ct-override-partname-invalid`,
`ct-override-partname-word-refused` and `ct-markup-compatibility` apply to
every package format; `word-content-type-invalid`,
`word-part-content-type-generic`, `word-part-content-type-inexact`,
`word-part-content-type-mismatch` and `word-part-content-type-unrecognized`
apply to DOCX/DOCM only.

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

`opc-item-name-refused` reports one finding per distinct stored file-entry name
that places no package part: one that starts with `/` or `\`, names a drive
(`c:`), or has an empty, `.`, `..` or other dot-only segment, read with `\` as
`/`. The finding's `path` is the stored name and its message is the package
view's refusal. It reads the same stored names as the checks above, and folder
entries are not judged. Every such name still refuses the staging, so the
finding is an error under every profile, and it is Word-blocking: Word 16.115
rejects every kind, even for an unreferenced entry. It offers recovery for a
leading `/` or `\`, a drive, and an empty, `.`, `..` or `...` segment (#411;
sessions 3, 42 and 43). The
`package-extract-failed` the refusal raises names only the first such name;
this check names each one
([#629](https://github.com/jandira-tech/docx-validate/issues/629)).

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

| Group                     | Named checks                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                       |
| ------------------------- | ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `package`                 | `zip-bytes-after-eocd`, `zip-bytes-before-archive`, `zip-eocd-comment-overrun`, `opc-equivalent-part-names`, `opc-part-name-escaped-unreserved`, `opc-part-name-escaped-slash`, `opc-part-name-word-refused`, `opc-item-name-refused`                                                                                                                                                                                                                                                                                              |
| `math`, `tracked-changes` | `word-math-run-insertion`, `word-math-run-deletion`                                                                                                                                                                                                                                                                                                                                                                                                                                                                                |
| `math`                    | `word-math-parse`, `word-math-spre-body`                                                                                                                                                                                                                                                                                                                                                                                                                                                                                           |
| `story`                   | `word-shading-color`, `word-page-margin-attributes`, `word-run-misplaced`, `word-run-properties-misplaced`, `word-underline-misplaced`, `word-paragraph-properties-misplaced`, `word-section-properties-misplaced`, `word-footnote-reference-missing`, `word-endnote-reference-missing`, `word-shared-story-part`, `word-numbering-picture-bullet`, `word-sdt-id-empty`, `word-text-contains-paragraph`                                                                                                                            |
| `tables`                  | `word-grid-column-attributes`, `word-cell-margin-order`, `word-table-properties-required`, `word-table-properties-misplaced`, `word-row-without-cell`, `word-cell-ends-with-table`                                                                                                                                                                                                                                                                                                                                                 |
| `drawings`                | `word-textbox-content-namespace`, `word-drawing-extent-missing`, `word-drawing-extent-misplaced`, `word-picture-nvpicpr-missing`, `word-picture-sppr-missing`, `word-empty-drawing`, `word-anchor-missing-attribute`, `word-position-offset-and-align`, `word-drawing-scalar-whitespace`, `word-drawing-arc-radius`, `word-drawing-extent-negative`, `word-part-reference-relationship`                                                                                                                                            |
| `content-types`           | `word-content-type-invalid`, `ct-duplicate-declaration`, `ct-override-partname-invalid`, `ct-override-partname-word-refused`, `ct-markup-compatibility`, `word-part-content-type-generic`, `word-part-content-type-inexact`, `word-part-content-type-mismatch`, `word-part-content-type-unrecognized`                                                                                                                                                                                                                              |
| `relationships`           | `word-duplicate-theme-relationship`, `word-duplicate-font-table-relationship`, `word-duplicate-web-settings-relationship`, `word-duplicate-settings-relationship`, `word-duplicate-styles-relationship`, `word-duplicate-numbering-relationship`, `word-duplicate-endnotes-relationship`, `word-duplicate-footnotes-relationship`, `word-hyperlink-relationship`, `word-header-footer-relationship`, `word-main-document-part`, `opc-relationships-structure`, `opc-relationships-target-mode`, `word-part-reference-relationship` |
| `tracked-changes`         | `word-textbox-deleted-text`, `word-move-from-deleted-text`                                                                                                                                                                                                                                                                                                                                                                                                                                                                         |
| `settings`                | `word-custom-property-filetime-empty`, `word-attached-toolbar-relationship`, `word-attached-toolbar-utf8`                                                                                                                                                                                                                                                                                                                                                                                                                          |
| `xml`                     | `word-markup-compatibility`, `word-document-root`, `word-choice-requires-undeclared`                                                                                                                                                                                                                                                                                                                                                                                                                                               |
| `comments`                | `comment-markers`, `comment-threading`                                                                                                                                                                                                                                                                                                                                                                                                                                                                                             |
| `schema`                  | `word-unknown-element`, `word-link-without-namespace`                                                                                                                                                                                                                                                                                                                                                                                                                                                                              |

`word-duplicate-endnotes-relationship` reports the main document's relationships part when it holds two or more exact-Type `endnotes` relationships and that main document's own content holds a `w:endnoteReference`: Word 16.115 offered recovery for every such shape probed — whichever part held the note, and whether the second part was typed `endnotes+xml` or `application/octet-stream` — and opened the same pair with no endnote reference clean ([#807](https://github.com/jandira-tech/docx-validate/issues/807); session 62, `s62c` and `s62e`, in [docs/plans/evidence/word-probes-2026-10-24.md](docs/plans/evidence/word-probes-2026-10-24.md)), while a referenced pair in a reached glossary's own sidecar is unprobed and stays unjudged. The count reads exact Types whatever their target or `TargetMode`, in either relationship namespace, so two relationships to the same part, an internal plus an External one and a Strict-Type pair are judged too; session 62's rows are two internal Transitional relationships to distinct parts, so those shapes are unprobed.

`word-duplicate-footnotes-relationship` reports the main document's relationships part when it holds two or more exact-Type `footnotes` relationships and that main document's own content holds a `w:footnoteReference`: Word 16.115 offered recovery for every such shape probed — whether the second part held another footnote id or the same footnote — and opened the unreferenced pair clean ([#827](https://github.com/jandira-tech/docx-validate/issues/827); session 63, `s63d`, and session 45, `s45g`, in [docs/plans/evidence/word-probes-2026-10-24.md](docs/plans/evidence/word-probes-2026-10-24.md)), while a referenced pair in a reached glossary's own sidecar and a reference only in a header, footer, endnote or comment are unprobed and stay unjudged. The count reads exact Types whatever their target or `TargetMode`, in either relationship namespace, so two relationships to the same part, an internal plus an External one and a Strict-Type pair are judged too; session 63's `s63d` rows are two internal Transitional relationships to distinct parts, so those shapes are unprobed.

`word-content-type-invalid` also reports an explicitly empty `ContentType=""` on the main document's Override, resolved through the package `officeDocument` relationship and the shared content-type table after Markup Compatibility processing ([#581](https://github.com/jandira-tech/docx-validate/issues/581)). Word 16.115.26100424 offers recovery for this specimen (session 22 in [docs/plans/evidence/word-probes-2026-10-24.md](docs/plans/evidence/word-probes-2026-10-24.md)). Sessions 49 and 50 (2026-10-26) observed the same recovery offer for an empty `ContentType=""` on the Override of the styles part the main document relates (s49b) and of a related settings or fontTable part (s50c), and session 53 (s53d) for a related numbering part, so the finding now covers those Overrides too, each citing its session ([#720](https://github.com/jandira-tech/docx-validate/issues/720), [#740](https://github.com/jandira-tech/docx-validate/issues/740)); one finding per part, named by the first relationship in the shared reporting order. It is a blocking error under Word-valid and an error under strict. Lenient runs no Word rule, so it reports nothing here beyond the XSD rejection. Full, quick and named Word-valid runs agree. Empty values on every other Override — a reached glossary's parts included — and on Defaults stay silent under Word-valid until probed; strict continues to reject them through XSD.

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
`word-part-content-type-inexact`, `word-part-content-type-mismatch`,
`word-part-content-type-unrecognized` and content-type
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
not exempt it; a match through either of two used references still exempts
the story reaches. Since [#739](https://github.com/jandira-tech/docx-validate/issues/739)
a type that matches one reach no longer exempts a mismatch on another
observed reach: every observed reach is judged over all of the part's
reaching entries, so a read footer cannot hide an eagerly parsed reach's
defect behind it. Unread parts are judged against all reaching segments
under strict (#607).

Word-valid reports the observed shapes only: read header/footer swaps — a
referenced footer typed `…header+xml`, or a referenced header typed
`…footer+xml`, from the main document or a reached glossary — the eager
reaches sessions 49 and 50 (2026-10-26) probed ([#722](https://github.com/jandira-tech/docx-validate/issues/722)):
the main document typed with another mapped Word type (s50d), and a part
the main document's own relationships part relates by `styles`, `settings`
or `numbering` typed with another mapped Word type — the main styles part
with the settings type (s49b), a related settings part with the styles type
and a related numbering part with the settings type (s50d); those parts are
eagerly parsed, so a relationship is enough and no usage lookup gates
them — and the read story co-reach: session 53 (s53a/s53b,
2026-10-26, [#739](https://github.com/jandira-tech/docx-validate/issues/739))
an eagerly read part — the main document, a related numbering part or the
styles part — that a footer reference the body makes also reaches, typed
with the main document, numbering or styles type alike, and the styles part
also targeted by the main document's `settings` relationship, typed styles
or settings alike; session 57 (s57c–s57e,
[#778](https://github.com/jandira-tech/docx-validate/issues/778)) the
reaches with no eager reach: a footer the body references typed with the
main document, numbering, styles or settings type, a header the body
references typed with the main document or styles type — the exact-typed
styles part a read header relationship targets included (s57e) — and a
reached glossary's own styles part a main footer relationship also reaches
(s57d); and session 58 (s58b,
[#785](https://github.com/jandira-tech/docx-validate/issues/785)) the
other read stories: a used comments, endnotes or footnotes part the body
references holding its own Transitional root
([#803](https://github.com/jandira-tech/docx-validate/issues/803)) and
typed with the styles type, a spelling with a parameter included (session
62 s62c, [#802](https://github.com/jandira-tech/docx-validate/issues/802)).
Session 63 (s63b) showed a comments part whose stored root is an
`mc:AlternateContent` holding one Choice requiring the w14 namespace with
only its own Transitional root, then one empty Fallback, typed styles exactly
([#818](https://github.com/jandira-tech/docx-validate/issues/818)); XML
whitespace, comments and processing instructions preserve that citation.
Session 65 (s65b–s65d, two runs, all readings agree) showed five more shapes:
text before the root inside that Choice, text in the Fallback, a Choice
requiring the Transitional Word namespace instead of w14, a Strict comments
root inside the w14 Choice, and a Strict root stored directly in a
Transitional package. Each was refused typed styles exactly and opened clean
typed comments, so these findings cite session 65
([#847](https://github.com/jandira-tech/docx-validate/issues/847)). Text
directly between the branches was refused typed comments too, so it supports
no content-type citation; #851 owns its `mce-malformed` finding. Other
Requires namespaces, a Strict root in the Fallback, text elsewhere inside the Choice or Fallback, or combinations of these edits, the endnotes and footnotes twins, and other spellings of
the new shapes are unprobed and keep their previous verdicts without
borrowing a citation, as do CDATA anywhere in the AlternateContent, its
Choice or its Fallback, between the branches included, and a Strict comments
root in a package whose main document is not Transitional
([#857](https://github.com/jandira-tech/docx-validate/issues/857)). Other stored roots that the effective view unwraps to
the story's root, including a selected Fallback and a Choice with no Fallback
or beside another Choice, remain unprobed for this type mismatch
([#837](https://github.com/jandira-tech/docx-validate/issues/837),
[#843](https://github.com/jandira-tech/docx-validate/issues/843)). The
exact-typed styles part that a comments, endnotes or footnotes relationship
nothing references also
targets opens clean (s58a/s58b) — the observed shape is the read story
typed styles, not the styles part co-reached. Word refused every co-reach
shape whatever the Override
says, so a matching reach exempts none of them; typings no session showed
on a story part stay silent. Session 24 observed Word 16.115 offer recovery
for the main document's header and footer swaps and the glossary footer swap,
including `…header+xml; charset=utf-8`; their exact-type controls and the
minimal package opened clean. Session 26 observed recovery for the glossary
header typed `…footer+xml` and for a read footer typed `…header+xml` despite
an unused header relationship to it; the latter's `…footer+xml` control
opened clean. The recovery sentence cites session 24 for the main header,
main footer and glossary footer swaps, and session 26 for the glossary
header swap (#610); every eager pair names the session that showed it —
session 49 the settings type on the styles reach (s49b), session 50 the
styles type on the main document and the settings reach and the settings
type on the numbering reach (s50d), session 53 the other mapped types it
ran on the main document (s53d–s53f) and on the settings, numbering and
styles reaches (s53f–s53h), session 56 the footer type on the styles reach
(s56a), session 57 the thirteen pairs no earlier session had shown:
the main document typed webSettings, the settings reach typed webSettings,
glossary, footer, footnotes or endnotes, the numbering reach typed
webSettings, glossary, theme or footer, and the styles reach typed
webSettings, glossary or footnotes
([#779](https://github.com/jandira-tech/docx-validate/issues/779)), and
session 59 the remaining three pairs: the styles reach typed endnotes and
the numbering reach typed footnotes or endnotes (s59a)
([#789](https://github.com/jandira-tech/docx-validate/issues/789)) — all
48 eager pairs are observed, and every recovery sentence names the session
that showed its pair; the
footer co-reach pairs cite sessions 53 and 57, the header co-reach session
57 (#739, #778), and the comments, endnotes and footnotes pairs typed
styles session 58 for a directly stored Transitional root, or session 62
when that type carries a parameter (#785, #802); the observed comments
AlternateContent and Strict-root shapes cite sessions 63 and 65 as described
above. The finding names
the relationship, declared
type and expected type, errors and blocks Word-valid; lenient warns. Unread
stories, a related fontTable or webSettings part, the glossary's reaches,
and comments/footnotes/endnotes typed with each other's types stay silent
under Word-valid until probed.

Strict judges every related part in the shared map, read or unread, citing
ECMA-376 Part 1 (4th edition), clause 11.3's content-type assignments (theme
is DrawingML and cites clause 14.2.7). The main document is named the Main
Document part and cites clause 11.3.10 (#609). Candidate relationships are
ranked in the same fixed order as the generic check, so relationship entry
order does not change the reported segment or clause (#608). Only an
observed finding — a read header/footer swap or read story co-reach,
the main document, or a related styles, settings or numbering part whose
pair a session showed — carries the Word recovery sentence; unobserved
shapes and a story part typed with
a type no session showed keep spec wording. Full, quick and named
selections agree
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
([#593](https://github.com/jandira-tech/docx-validate/issues/593)); a read
story relationship hides no eager one either — session 56 (s56a, 2026-10-26)
observed Word recover the styles part typed `…styles+xml; charset=utf-8`
with a read footer relationship also targeting it, so the spelling is
judged over every reach the story roots read
([#760](https://github.com/jandira-tech/docx-validate/issues/760)). Strict
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

`word-part-content-type-unrecognized` reports a part Word reads whose
resolved media type — trimmed, parameters stripped, compared ASCII
case-insensitively — is none of the types the package's own relationships
expect: not its own Word content type, not a generic XML type
(`word-part-content-type-generic`'s class), and not another mapped Word
part's type (`word-part-content-type-mismatch`'s class), so the four part
content-type checks stay disjoint
([#707](https://github.com/jandira-tech/docx-validate/issues/707),
[#719](https://github.com/jandira-tech/docx-validate/issues/719): an
`invalid-*` declared value is `word-content-type-invalid`'s class alone).
Sessions 46,
49 and 50 (chunks s46e, s49a/s49b and s50c/s50d in
`docs/plans/evidence/word-probes-2026-10-24.md`) observed Word 16.115 offer
recovery for the main document itself — its Override typed
`application/octet-stream` (session 49), `APPLICATION/OCTET-STREAM`,
`application/octet-stream; charset=utf-8`, `image/png` or
`application/foo+xml` (session 50) — for the main document's `styles` part —
`word/styles.bin` typed `application/octet-stream` by a Default, whether the
part held non-XML bytes or well-formed styles XML, while the same XML typed
as styles opened clean (session 46), and `word/styles.xml` typed by its
Override with the same five spellings (session 49) — and for the
`settings`, `fontTable` and `numbering` parts the main document's own
relationships relate, typed `application/octet-stream` by their Override
(session 50) and typed the same four other spellings — `APPLICATION/OCTET-
STREAM`, `application/octet-stream; charset=utf-8`, `image/png` or
`application/foo+xml` — (session 53, s53b–s53d)
([#721](https://github.com/jandira-tech/docx-validate/issues/721),
[#740](https://github.com/jandira-tech/docx-validate/issues/740)). That
observed reach — the main document and the parts its own relationships part
relates by `styles`, `settings`, `fontTable` and `numbering` — carries the
Word recovery sentence and blocks under word-valid; every other segment the
check's spec-set walk reaches — a related webSettings or theme part, a
header, footer, footnotes, endnotes, comments or glossaryDocument part, and
the parts a reached glossary's own relationships relate, its styles
included — cites the content type ECMA-376 Part 1 (4th edition), clause
11.3 assigns it (14.2.7 for theme) as an
error under strict and a warning under word-valid. An undeclared type stays
with `ct-undeclared-part`, and an empty one with `word-content-type-invalid`
on the main document and on the styles, settings, fontTable and numbering
parts the main document relates
([#720](https://github.com/jandira-tech/docx-validate/issues/720),
[#740](https://github.com/jandira-tech/docx-validate/issues/740); sessions
22, 49, 50 and 53); every other segment's empty Override was not shown to Word
and stays unreported. An `invalid-*` declared value stays with
`word-content-type-invalid` alone; lenient warns. One
finding is reported per part, and when several relationship types reach it
the observed reach names the finding whenever it exists. Full, quick and
named runs agree.

The `schema` group contains structural checks, not an XSD walk. The `comments`
group runs `comment-markers` (range pairing and references to existing comments)
and `comment-threading` (marker counts, paragraph ids, parent ids, and durable ids
across comment extension parts). Threading reads the main document, the
comments part and all three extension parts after markup compatibility's
branch selection, as `comment-markers` reads its parts (#799, #808). Both
apply to DOCX/DOCM under every profile,
including lenient; they are separate from `WORD_RULE_CHECKS`. They retain their
own `comment-marker-parse` / `comment-thread-parse` diagnostics rather than
requiring `xml-syntax`. A named run that selects them also runs `xml-syntax`
and reports a malformed part once, through `xml-syntax`, as the full run does
([#385](https://github.com/jandira-tech/docx-validate/issues/385)). Under lenient that run stops where `xml-syntax` rejects a part, as the
full lenient run stops at a malformed part, so the comment checks then do not run
([#391](https://github.com/jandira-tech/docx-validate/issues/391)). Full validation runs them once and records their ids in
`executedChecks`; `--quick` does not select them. Like every other check, they
find a comment part under an OPC-equivalent name (`word/comments.XML`).

A dangling `w:commentReference` is a Word-valid error; a range start or end
naming a missing comment without that reference stays a warning. The
`comment-marker-missing` finding carries a registered reference fact, so its
verdict does not depend on message text. The `story` checks
`word-footnote-reference-missing` and `word-endnote-reference-missing` likewise
reject references whose ids name no matching note. Session 58 (s58a, two runs,
all four readings agree) showed recovery for the dangling comment and endnote
references and clean opens for their matching-id calibrations
([#784](https://github.com/jandira-tech/docx-validate/issues/784)).
`comment-markers` collects markers and comment ids in both the Transitional
and the Strict WordprocessingML namespaces, the ids attributes in the
marker's own namespace, so an ISO Strict package's comments are judged too:
session 63 (s63a in `docs/plans/evidence/word-probes-2026-10-24.md`, two
runs, all readings agree) offered recovery for a Strict `w:commentReference`
naming no comment its comments part defines, and opened the unchanged
Word-saved Strict calibration clean
([#817](https://github.com/jandira-tech/docx-validate/issues/817)), so the
Strict dangling reference blocks as its Transitional counterpart does. A
Strict orphan range start or end, now paired too, is unprobed. The ids are
pooled across the two namespaces, so a reference in one namespace satisfied
only by a comment in the other passes; mixed-namespace satisfaction, in
either direction, is unprobed.
No reference inside a notes part is judged, matching session 16's clean
open, and none inside a comments part either
([#852](https://github.com/jandira-tech/docx-validate/issues/852), below):
Word does not follow those references. The defined notes are read
only from the parts a `footnotes` / `endnotes` relationship of the main
document or of a reached glossary names, whatever the part is called: session
60 (s60a in `docs/plans/evidence/word-probes-2026-10-24.md`, two runs, all
readings agree) opened clean a used notes part stored as `word/notes-foot.xml`
and `word/notes-end.xml` and related by its relationship, a shape the
name-only reading rejected
([#793](https://github.com/jandira-tech/docx-validate/issues/793)), and
session 61 (s61a/s61b, same evidence, two runs, all readings agree) offered
recovery for a body reference whose note sat only in a `word/footnotes.xml` or
`word/endnotes.xml` no relationship names, while the same part with nothing
referencing it opened clean
([#797](https://github.com/jandira-tech/docx-validate/issues/797)). A notes
part the content-types stream types `application/octet-stream` defines no
note, whatever it is called: session 63 (s63c, same evidence, two runs, all
readings agree) offered recovery for a body footnote reference whose note sat
in `word/notes.bin` typed octet-stream by its Override, and session 64 (s64b,
same evidence, two runs, all readings agree) repeated that refusal for the
same row and for `word/notes.xml` and `word/footnotes.xml` alike, so the
type, not the name, makes Word refuse the part
([#829](https://github.com/jandira-tech/docx-validate/issues/829)). References
read through the main document use only the notes parts its own relationships
name: session 63
(s63c, same evidence, two runs, all readings agree) offered recovery for a
body footnote reference whose note only a reached glossary's notes part
defines, while the main-related matching calibration opened clean
([#828](https://github.com/jandira-tech/docx-validate/issues/828)); the finding's
message then names the glossary
([#838](https://github.com/jandira-tech/docx-validate/issues/838)). The endnotes
twin is unprobed; generalising from footnotes follows #793's footnote-to-endnote
reading. References read through a reached glossary keep the union of the main
document's and reached glossaries' definitions; that reading is unprobed.
Both note checks read references only from the story roots — every main document
and reached glossary, found by relationship wherever stored — so a story-named
part no relationship reaches is judged for
none ([#800](https://github.com/jandira-tech/docx-validate/issues/800)):
session 62 (s62a/s62d, same evidence, two runs, all readings agree) offered
recovery for a dangling footnote or endnote reference in a main document stored
as `word/main.xml`, and opened clean an unrelated `word/header9.xml` holding
one.
References in a read header or footer are not judged (session 62 s62d/s62e
for a header,
[#806](https://github.com/jandira-tech/docx-validate/issues/806);
session 63 s63d for a footer,
[#826](https://github.com/jandira-tech/docx-validate/issues/826)); nor is a
reference inside a notes part or a comments part
([#852](https://github.com/jandira-tech/docx-validate/issues/852)): session 16
(s16a) had opened an endnote reference inside a footnote clean, and sessions 65
and 66 (s65a, s66a, s65d, in
[docs/plans/evidence/word-probes-2026-10-24.md](docs/plans/evidence/word-probes-2026-10-24.md),
two runs, all readings agree) opened clean every nesting they stored — footnote
1 holding a footnote reference to 99 and endnote 1 an endnote reference to 99,
in-package or behind the escaping target (s65a), endnote 1 a footnote reference
to 99 with no footnotes part anywhere (s66a), and a comment a footnote reference
to 99 or to 1 only a reached glossary's notes part defines, the comments part
read by the main document or by a glossary (s66a, s65d) — so session 16's
special case is subsumed and gone, and a dangling reference in a notes or
comments part no longer blocks. A comments part both the main document and a
glossary read is refused, but by the shared reach, not the reference (s66a
refused it with no note reference at all), which is #850's rule. An endnote
reference in a comments part, and a note reference in a glossary's own notes
part, are unprobed.
Both note checks compare ids by integer value: session 62 (s62b, same
evidence, two runs, all readings agree) opened clean a footnote or endnote
reference spelled `01` or `+1` against a note spelled `1`
([#801](https://github.com/jandira-tech/docx-validate/issues/801)); any other
spelling compares as written, unprobed. When the package view cannot be read, or
its `officeDocument` relationship names no stored main document, the read model
and the relationship walks return nothing, so both note checks fall back to the
reading they used before #800 — references from the name-based story list but
its headers, footers, comments and notes parts, which are not judged (#806,
#826, #852), definitions from
the parts named `footnotes.xml` / `endnotes.xml` wherever stored — as the
comment checks fall back to a by-name lookup when no relationship names the
main document, and a
dangling reference still draws its finding beside the package defect's own
([#816](https://github.com/jandira-tech/docx-validate/issues/816)); no Word
session has shown a note reference in such a package, so the fallback reading is
unprobed.

A comment range whose start only an unselected `mc:Choice` holds, so the
effective view holds an unmatched end, opened clean, and word-valid accepts it with warnings (session 63 s63b, same
evidence, two runs, all readings agree; the selected-`w14` spelling is its
clean calibration,
[#820](https://github.com/jandira-tech/docx-validate/issues/820)); the
mirror, an end only such a Choice holds, is unprobed. A body
footnote reference whose note sits in `word/footnotes.xml` while the main
document's only `footnotes` relationship is External gets the recovery offer
and blocks (session 63 s63c, same evidence). When that relationship instead
targets outside the package, the note definitions resolve the target by
stopping its `..` segments at the package root, as RFC 3986's
remove_dot_segments does: session 63 (s63c, same evidence) and session 64
(s64a in `docs/plans/evidence/word-probes-2026-10-24.md`, two runs, all
readings agree) opened `word/`'s `../../footnotes.xml` and
`../../../footnotes.xml` clean with the note at `/footnotes.xml` — and the
endnotes twin `../../endnotes.xml` with its note at `/endnotes.xml` — while
the same escaping target with the note only in `word/footnotes.xml` got the
recovery offer and stays rejected, and the in-package `../footnotes.xml`
calibration opened clean
([#830](https://github.com/jandira-tech/docx-validate/issues/830)). Other
rules' relationship resolution keeps dropping an escaping target; those
readings are unprobed.

`word-shared-story-part` rejects a comments, footnotes or endnotes part that
the read model lists for both a main-document entry and a glossary entry
(`wordReadTypedTargets`) when each of the two roots is known to use it — a
reference in its readable effective XML; the model also keeps a reach whose
use it cannot decide, so a root whose XML fails the byte-level checks reports
no shared part, and its own XML diagnostics block the run
([#858](https://github.com/jandira-tech/docx-validate/issues/858)): Word session 66
(s66a in `docs/plans/evidence/word-probes-2026-10-24.md`, two runs, all four
readings agree) offered recovery for a `word/comments.xml` the main document
and a reached glossary both relate and reference — its comment holding
nothing but text — while the same part read by the main document alone or by
the glossary alone opened clean, and session 67 (s67a, same evidence, same
readings) showed the refusal needs both roots to use the part — related by
both but referenced by one, or by neither, opens clean — and covers notes
parts too: a `word/footnotes.xml` the glossary relates as `../footnotes.xml`
and both roots reference is refused, as is its endnotes twin, while the same
part referenced by the body only opens clean
([#850](https://github.com/jandira-tech/docx-validate/issues/850)). The
finding cites session 66 for comments and session 67 for footnotes and
endnotes, and reports the part through the main document's spelling. Shared
header and footer parts, a part two glossaries share, and a part one root
reads under one of these segments and the other under another are unprobed
and unjudged.

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
include the legacy core spelling and Strict extended/custom spellings. A properties
relationship inside an `mc:Fallback` or an understood `mc:Choice` names the part as a
direct child does ([#634](https://github.com/jandira-tech/docx-validate/issues/634)).
Targets resolve
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
Validation leaves source bytes unchanged unless you request `autoRepair` on an
unpacked directory; a packed input is repaired only on a discarded staged copy.

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

`opc-relationships-target-mode` requires a present `TargetMode` to be exactly
`Internal` or `External` in the package relationships part (`_rels/.rels`),
the main document's relationships part and a reached glossary's sidecar,
found by relationship wherever the story root is stored, with OPC part-name
comparison. Word 16.115 offered
recovery for `external` on a referenced chart or diagramData relationship,
`external`, `External ` (trailing space) or an empty value on an unreferenced
hyperlink in the main sidecar, and `external` on that hyperlink in the package
relationships part. The `Internal` chart and diagramData calibrations and
`External` hyperlinks opened clean. The same hyperlink with `external` in a
referenced header's own sidecar also opened clean; headers keep their verdict.
Those rows are from [#753](https://github.com/jandira-tech/docx-validate/issues/753),
session 55, `s55b`–`s55e` in `docs/plans/evidence/word-probes-2026-10-24.md`.
Session 57 (`s57f`) refused the same unreferenced hyperlink with `external`
in a reached glossary's sidecar, while its `External` calibration opened clean
([#775](https://github.com/jandira-tech/docx-validate/issues/775)). Other
relationships parts (footnotes, comments and so on) are unshown and keep their
verdict.
Full word-valid and named runs block the same findings; strict reports errors,
and lenient runs no Word rule. An absent `TargetMode` is not this check's finding.
A hyperlink naming a relationship with a misspelled present mode in this scope
is reported here alone, without a second `word-hyperlink-relationship` error
for that attribute ([#773](https://github.com/jandira-tech/docx-validate/issues/773)).

`opc-relationships-structure` requires every relationships part to have a
`Relationships` root in the package-relationships namespace and only `Relationship`
elements in that namespace as direct children after Word's markup compatibility
processing. It reports one finding for a wrong root, or one per surviving
offending child, at the stored part name. The root is judged on the processed
view in the package relationships part (`_rels/.rels`, compared by OPC part-name
equivalence): Word 16.115 opened a root that markup compatibility resolves to
exactly one `Relationships` element clean, whether an `mc:AlternateContent`
Fallback, a selected `mc:Choice` or an Ignorable `mc:ProcessContent` wrapper
holds it. It offered recovery for a root resolving to two `Relationship`
elements or to none (the `Relationships` element in an unselected Choice beside
an empty Fallback), which keep the finding because the processor returns the
stored view there
([#646](https://github.com/jandira-tech/docx-validate/issues/646),
[#640](https://github.com/jandira-tech/docx-validate/issues/640); Word probes
sessions 36 and 38, `s36c` and `s38a`). Combinations resolve the same way:
Word opened six clean in session 39 (`s39b`) — nested Fallback wrappers two
and three deep, an unselected Choice before a selected one, the first of two
selected Choices, a wrapper carrying `mc:MustUnderstand="w14"`, and a wrapped
`Relationships` element with an ignorable child
([#662](https://github.com/jandira-tech/docx-validate/issues/662)).
A root-wrapped sidecar is unobserved and keeps the stored-root verdict, as does
every part under lenient, whose XSD walk judges the stored root. Malformed XML is
left to `xml-syntax`. The check runs under Word-valid and strict, including quick and
named relationships selections; explicit lenient keeps its existing checks.
Word-valid blocks malformed package relationships and malformed relationships
parts under `word/`, and warns on
other sidecars, which have no Word evidence yet. Strict reports errors on all of
them; under strict and lenient the XSD walk still judges the stored
`_rels/.rels` root, so a root wrapper the Word rule accepts stays invalid there.
Main-part discovery uses the processed package relationships and ignores relationships
still nested inside foreign wrappers ([#433](https://github.com/jandira-tech/docx-validate/issues/433),
[#434](https://github.com/jandira-tech/docx-validate/issues/434)).
`opc-relationships-markup-compatibility` reports each stored child that processing
removes or replaces with only direct `Relationship` children (including an empty
replacement). Word-valid warns; strict keeps an error because stored OPC
relationships allow no such children. A foreign child left in a selected branch
still gets `opc-relationships-structure`. Word 16.115 opened each of the four
probed shapes of this class cleanly — the footer with its relationship in an
unknown-Choice Fallback, in a `w14` Choice, and beside an Ignorable foreign
child, and the package with its main relationship in a Fallback; other shapes
are unobserved. A foreign child left in that Fallback prompted recovery (Word probes
session 33, `s33d`/`s33e`; [#632](https://github.com/jandira-tech/docx-validate/issues/632)).
The same processed view drives reachability: the parts Word reads (the
Word-read parts, the story roots and the properties parts) are found through
the relationships markup compatibility selects, so a part reached only through
a relationship inside an `mc:Fallback` or an understood `mc:Choice` blocks
Word-valid exactly as its directly related twin does. Word 16.115 refused a
truncated main document behind a Fallback `officeDocument` relationship (open
error), and offered recovery for that Fallback main document's zero-byte
relationships part, for a truncated `docProps/core.xml` behind a Fallback
core-properties relationship, and for a footer typed as a header whose
relationship sat in the Fallback — each beside its direct-relationship twin,
which it refused identically (session 34, `s34a`/`s34b`;
[#634](https://github.com/jandira-tech/docx-validate/issues/634)). A
relationship nested inside a foreign wrapper that survives processing stays
unfollowed ([#434](https://github.com/jandira-tech/docx-validate/issues/434)).
Word also honours markup-compatibility errors in relationships parts:
`mc:MustUnderstand` of a namespace it does not understand on the `_rels/.rels`
root, and an `mc:Choice` whose `Requires` prefix is undeclared in
`word/_rels/document.xml.rels`, each prompted recovery (session 34, `s34c`/`s34d`;
[#636](https://github.com/jandira-tech/docx-validate/issues/636)). The
`word-markup-compatibility` check reports those findings (see below).
Hyperlink and toolbar resolution, duplicate counts and the validators' semantic
relationship checks read the processed view too, still taking direct children
only, under Word-valid and strict; for the semantic checks that view is
DOCX/DOCM only (#635, #643, above). Legacy lenient traversal remains
unchanged, and so does the repair inventory's raw counting — it selects
which parts count as styles parts through the same relationship walks the
style checks use (#745), but its counters still read the stored bytes.

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
A present `TargetMode` other than exactly `Internal` or `External` on an
otherwise accepted hyperlink relationship is left to
`opc-relationships-target-mode` wherever that check judges the relationships
part. This also applies to a named `word-hyperlink-relationship` run: select
`opc-relationships-target-mode` too to report that shape. An exactly `Internal`
or absent mode, another Type, an unresolved id, and a misspelled mode in a
header's sidecar still receive the hyperlink finding
([#773](https://github.com/jandira-tech/docx-validate/issues/773)). A reached
glossary's misspelled mode now belongs to the target-mode check (session 57,
`s57f`; [#775](https://github.com/jandira-tech/docx-validate/issues/775)).
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

`word-part-reference-relationship` requires every `c:chart r:id`, `dgm:relIds`
`r:dm` and `o:OLEObject r:id` in the selected markup-compatibility view of a
part Word reads to be present and non-empty and to name, in that part's own
relationships part and in either
relationships namespace, a relationship whose Type is exactly the one the
element reads: an internal `chart` relationship for `c:chart` (an External
chart relationship is refused too), an internal `diagramData` relationship
for `dgm:relIds r:dm`, and a `package` or `oleObject` relationship for
`o:OLEObject` — internal or External alike for `package`, which session 52
opened clean as an External (linked) package relationship (B1), and whose
External `oleObject` spelling was not shown to Word and stays unjudged. Word
16.115
offered recovery for every refused shape probed — an `image` relationship to
a real PNG or to a real chart part, a `diagramData`, `package` or `chart`
relationship, an External chart relationship with and without a part, an
External `hyperlink` relationship, the near-miss Type spellings
`…/relationships/chart/` and `…/relationships/Chart` naming a real chart, an
Id no relationship has, and an empty `r:id`/`r:dm` present on the element
(`dgm:relIds r:dm=""` drew an open error on one specimen and the recovery
offer on another) — on the minimal Transitional package and as
one-attribute mutations of the Word-clean ISO Strict `word-clean-strict01.docx`,
while the matching relationships opened clean
([#705](https://github.com/jandira-tech/docx-validate/issues/705),
[#728](https://github.com/jandira-tech/docx-validate/issues/728); sessions 45,
46 and 52, chunks s45b–s45k, s46a–s46c and s52a–s52d, in
[docs/plans/evidence/word-probes-2026-10-24.md](docs/plans/evidence/word-probes-2026-10-24.md);
`r:dm` alone naming the real relationship is the s52d confound and opens
clean). An id the element does not carry at all is still not judged.
The parts judged are the parts Word reads, found by relationship — every main
document and reached glossary wherever they are stored, plus the headers and
footers a section reference names and the comments, footnotes and endnotes
parts the document uses — not the parts whose name looks like a story: a
header the main document relates but no `w:headerReference` names, with its
relationships part and without one, and a glossary no relationship reaches,
all open clean, while the same chart reference in `word/headerFirst.xml` or in
the main document stored as `word/main.xml` gets the recovery offer
([#726](https://github.com/jandira-tech/docx-validate/issues/726); session 51,
chunk s51a, in the same evidence file). Partial named-check runs walk the same
read set, so their verdict is the full run's (#381).
The target part never matters: the relationship's Type decides, as the
header/footer rule's does. `r:lo`, `r:qs` and `r:cs` are deliberately not
checked — Word opened them clean naming the image and the chart relationship,
with the part's cached diagramDrawing rendered — and neither are `a:blip
r:embed` nor `v:imagedata o:relid`, which Word opened clean naming a
relationship of another type. A `c:chart` carrying an id attribute in both
relationships namespaces — `r:id` and the ISO Strict `rs:id` together,
whatever their values — cannot be opened: an open error with no recovery
offer (session 55, s55b and s55d;
[#756](https://github.com/jandira-tech/docx-validate/issues/756)), while
`rs:id` alone opens clean and draws the chart; the same pair on the other
references was not shown to Word and stays unjudged. A part with no
relationships part is judged
like one whose relationships part lacks the id — session 55 (s55a) saw
Word refuse a referenced `word/header1.xml` with no relationships part
holding each of the three references, as it does one whose relationships
part lacks the id, and open clean the same header with no `r:` reference.
Session 57 confirms the recovery offer for `c:chart r:id="rId1"` with no
relationships part in used footnotes and comments (`s57f`
`footnotes-nosidecar-chart`, `comments-nosidecar-chart`), a reached glossary
and the main document (`s57g` `glossary-nosidecar-chart`, `main-nosidecar-chart`);
all four text-only calibrations opened clean (`s57g`, two runs, all readings
agree; [#775](https://github.com/jandira-tech/docx-validate/issues/775)). Endnotes
with no relationships part were not shown and keep the same rule.
`rels-missing-sidecar` still reports the absent sidecar where its
walk fires — an `r:id`, `r:embed` or `r:link` attribute in a relationships
namespace, which the `dgm:relIds` shape's `r:dm`, `r:lo`, `r:qs` and
`r:cs` are not, so that shape's only finding is this rule's; an
unparseable sidecar is skipped for `xml-syntax` to report. The finding is an
error, and Word-blocking, under word-valid and strict; lenient runs no Word
rule. In a full word-valid or strict run, a dangling Transitional `r:id` on a
`c:chart` or an `o:OLEObject` is reported once, through this rule: every
legacy `rels-id-missing` copy of the same Id on the same element in the same
part is handed off — including the unselected twin's copy of the two an
element with an `mc:Choice` and an `mc:Fallback` twin yields, which the
raw-part walk sees twice and this rule once (#727; the key is the same
element name, namespace and local name, with the same Id in the same
part, not the element instance) — as the header/footer rule's is
(#612, #619); lenient keeps the legacy findings alone.

Under word-valid for DOCX/DOCM, `rels-missing-sidecar` blocks only when the
part is a story root, a member of the run's Word-read set, or a part that a
known Office relationship of an unobserved type names wherever it is stored,
and a counted relationship reference survives in its processed
markup-compatibility view. The Word-read set holds such parts only under
`word/`; Word was never shown one outside it with a reference and no sidecar,
so an unobserved type's reach keeps the blocking decision it had before #638.
The chart type is the one exception: session 36 observed that Word 16.115
never loads a chart that no drawing references — a chart part related from
the main document's sidecar (`…/relationships/chart`) but referenced by
nothing opened clean, stored under `word/` and outside it — so a part every
reaching relationship names through the chart type blocks only when a
Word-read part's effective XML names one of those relationship Ids
(`c:chart/@r:id` in a drawing, read in the processed view: a drawing only
inside an unselected `mc:Choice` opened clean and one inside a selected
Choice got the recovery offer, session 37 s37b). Session 44 opened clean
every part an image relationship reaches, under `word/` and outside it,
unreferenced or named by a picture's `a:blip r:embed`, and session 46 (s46d)
opened the same chart clean when a VML picture's `v:imagedata r:id` or a
linked picture's `a:blip r:link` names the relationship; session 52 (E1,
E2) opened the same no-sidecar chart clean again when a `w:name r:id` names
the reach and when a VML picture's `v:imagedata r:id` inside a selected
`mc:Fallback` does: an image reach
never loads the part, so `rels-missing-sidecar` always warns for it, however
its Id is named
([#706](https://github.com/jandira-tech/docx-validate/issues/706); #701's
extension is replaced by #705's rule, which blocks the `c:chart r:id`,
`dgm:relIds` and `o:OLEObject r:id` shapes naming an image relationship for
the reference's Type). Session 44 also opened clean an unreferenced
`diagramData`, `oleObject` or Microsoft `hdphoto` reach, so those types join
the rule; session 45 (S1) opened clean a picture's `a:blip r:embed` naming
one of them — session 52 (D1) opened the same clean hdphoto reach again, a
chart part with no sidecar an `hdphoto` relationship a picture's
`a:blip r:embed` names — so every other attribute in a relationships namespace (or
VML's `o:relid`) naming the relationship counts as their reference, and a
`dgm:relIds r:dm` keeps blocking
([#647](https://github.com/jandira-tech/docx-validate/issues/647);
[#702](https://github.com/jandira-tech/docx-validate/issues/702)). A part
any other relationship type also reaches keeps the reach-based verdict
([#645](https://github.com/jandira-tech/docx-validate/issues/645); chunks
s36b/s36d, s37b and s44a–s44c in the same evidence file).
Session 39 (s39a, s39c) showed Word the rule beyond the main document's
sidecar: an unused chart relationship from a referenced header's sidecar or
from a reached glossary's footnotes part or document, two unused chart
relationships to one part, a relationship targeting an OPC-equivalent
spelling of the chart's name, and a drawing only in an `mc:Fallback` Word
does not select all opened clean, and each shape's referenced twin got the
recovery offer ([#661](https://github.com/jandira-tech/docx-validate/issues/661)).
Session 35, chunks s35a/s35b, observed clean opens for ids in unreferenced
`word/xtra.xml` (both relationships namespaces), unreferenced `docProps/app.xml`,
an unreached glossary, on `w:name` in reached styles, and inside an unselected
`mc:Choice` (both relationships namespaces). These findings now warn under
word-valid; strict and lenient output is unchanged
([#638](https://github.com/jandira-tech/docx-validate/issues/638)). A footer
reference's Strict relationships id inside `w:sectPrChange` still blocks:
Word offered recovery for it in s35b.

The processed-reference test uses declaration sets generated from all bundled
XSDs by `scripts/generate-relationship-reference-elements.ts`, with a test
that regenerates them in memory to check for drift. The sets follow global
and local element declarations, element references, type extension and
restriction bases, and attribute groups. They union declarations sharing an
element name and map Transitional namespaces to their ISO Strict counterparts.
A non-empty `id`, `embed` or `link` in either relationships namespace stops
counting only when the schemas declare the element and no declaration gives
it that attribute. `w:name` declares only `w:val`, so its added relationship
id does not qualify. Undeclared elements and namespaces, including extension
namespaces such as `asvg`, keep counting; an unresolved type also cannot prove
absence. Charts, ambiguous names with a reference-bearing declaration and
unobserved extensions retain the previous blocking decision. This conservative
classification is not fresh Word evidence for those shapes. It loads or
compiles no XSD during word-valid validation.

A story root — the main document or a reached glossary — keeps precedence
over the header/footer basename tolerance, under any `.xml` name, even
`word/header1.xml`: session 33 observed recovery for both renamed roots
whose `.rels` was gone beside their story reference, while the renamed
controls opened clean
([#625](https://github.com/jandira-tech/docx-validate/issues/625); session 33,
chunk s33b, in the same evidence file). A genuine `word/header1.xml` or
`word/footer1.xml` bearing relationship references with no sidecar keeps its
word-valid warning — Word's own regenerated output opens clean carrying
that finding. A story root stored without a `.xml` suffix
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
adding one with that Id and Type when the part has none. Strict also keeps the XSD error for a
missing `r:id`. The finding is an error under
strict and Word-blocking under word-valid for DOCX/DOCM; lenient runs no Word
rule. In a full word-valid or strict run neither `rels-id-mismatch` nor
`rels-id-missing` repeats an element this check reported — the run drops
every copy of the same Id on the same element name (namespace URI and
local name, not the element instance) in the same part for a non-empty id,
including the unselected twin's copy of the two copies an element with an
`mc:Choice` and an `mc:Fallback` twin yields (#727), pairing the findings
by the element identity both checks register on the
issue (part path, namespace URI, local name and raw `r:id`; the namespace
the `r:id` was read in does not distinguish an element name's other
spelling), never by message text — while `validateAllRelationshipIds()` called on its own
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

`word-markup-compatibility` reports processor errors for the XML parts under
`word/` and for every relationships part, at the stored part name
([#636](https://github.com/jandira-tech/docx-validate/issues/636)). The judged
names are compared as OPC compares part names — canonical, ASCII case folded,
escapes of unreserved characters decoded — so `word/FOO.XML` and
`word/foo.%78ml` (an escaped `x`) are judged as `word/foo.xml` is, on both the
check and the named runs' root-Ignorable pairing
([#639](https://github.com/jandira-tech/docx-validate/issues/639)). An escaped
spelling keeps its own `opc-part-name-escaped-unreserved` rejection, so no
verdict changes; the escaped shapes themselves were not shown to Word.
Word 16.115 applies markup compatibility to relationships parts before judging their children
(session 33, `s33d`/`s33e`) and honours its errors there: `mc:MustUnderstand` of a
namespace it does not understand on the `_rels/.rels` root (session 34, `s34c`)
and an `mc:Choice` whose `Requires` prefix is undeclared in
`word/_rels/document.xml.rels` (session 34, `s34d`) each prompted recovery.
Word-valid therefore blocks `mce-malformed` and `mce-must-understand` in
`_rels/.rels` and in the sidecar of the main document and of every reached
glossary, found by relationship even when the main document is renamed
(sessions 34 and 36, [#644](https://github.com/jandira-tech/docx-validate/issues/644)). Sidecars that are not a story root's — stray,
styles, footer, settings, numbering and header — opened clean with a malformed
sibling or root `mc:MustUnderstand` (sessions 36 and 37), so Word-valid demotes
them to warnings, as it does in the remaining unobserved relationships parts;
strict keeps the error. Lenient runs no Word rule. Full, quick and named runs
report the same findings and severities.
If unwrapping a part's root through AlternateContent or ProcessContent would
leave zero or multiple element children, the processor reports error-severity
`mce-malformed` and returns the part unprocessed
([#640](https://github.com/jandira-tech/docx-validate/issues/640)). One element
continues to be unwrapped. The source is unchanged. The stored-root rules keep
their verdict, including blocking `opc-relationships-structure` for a wrong
relationships root and `word-document-root` for a wrong main document root —
except the package relationships part's root, which `opc-relationships-structure`
judges on the processed view
([#646](https://github.com/jandira-tech/docx-validate/issues/646)).
Word 16.115 offered recovery for a `_rels/.rels` root AlternateContent whose
Fallback holds two `Relationship` elements (session 36, `s36c`) and for one that
unwraps to none, its `Relationships` element in an unselected Choice beside an
empty Fallback (session 38, `s38a`); both keep their blocking
`opc-relationships-structure`. Roots that unwrap to the `Relationships` element
opened clean and pass Word-valid
([#646](https://github.com/jandira-tech/docx-validate/issues/646)). Other zero-
and multiple-element root shapes are unobserved. Word-valid and strict return
findings instead of throwing or reporting `check-crashed`; full, quick and named
runs agree on their shared checks. Lenient retains its prior findings, including
its legacy XSD diagnostic for multiple roots and its empty-view diagnostic for
zero roots.

Root validation, the processor, and the schema walk use these stable codes:

| Code                          | Meaning and profile behavior                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                     |
| ----------------------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ |
| `word-document-root`          | The main part has a foreign or missing effective root: Word refuses to open it or silently drops all content. Blocking in Word-valid and strict.                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                 |
| `mce-malformed`               | Invalid compatibility structure, including AlternateContent without a Choice, AlternateContent holding non-whitespace character data directly inside it (Word offered recovery for every position sessions 65 and 67 stored, in a used comments part's root and a body run; whitespace only, and text inside a Choice or the Fallback, opened clean; other parts and relationships parts are unprobed, #851), or root unwrapping that leaves zero or multiple elements. Blocking in Word-valid and strict in the XML parts under `word/`; in relationships parts in `_rels/.rels` and the sidecar of the main document or any reached glossary, found by relationship even when renamed (sessions 34 and 36, #644). Malformed siblings in stray, styles and footer (session 36) and settings, numbering and header (session 37) sidecars opened clean; other sidecars are unobserved and keep their Word-valid warning. Strict keeps all errors. |
| `mce-must-understand`         | A declared required namespace is not understood by the selected view. Blocking in Word-valid and strict in the XML parts under `word/`; in relationships parts in `_rels/.rels` and the sidecar of the main document or any reached glossary, found by relationship even when renamed (sessions 34 and 36, #644). Root MustUnderstand in stray (session 36), styles and footer (session 37) sidecars opened clean; other sidecars are unobserved and keep their Word-valid warning. Strict keeps all errors. Word and schema capabilities can differ.                                                                                                                                                                                                                                                                                                                                                                                            |
| `mce-unbound-prefix`          | A directive names an unbound prefix. Error in lenient/strict schema processing; warning in full Word-valid. A repeated root Ignorable finding is retained once as `ignorable-undeclared`.                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                        |
| `mce-invalid-token`           | An invalid ProcessContent QName. Error in lenient/strict schema processing; warning in full Word-valid.                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                          |
| `mce-root-ignored`            | An ignorable root disappears from the schema view. Informational for non-main parts; not emitted by the full Word-valid path.                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                    |
| `mce-main-root-ignored`       | The required main document root disappears from the schema view. Error in lenient/strict; the Word view reports `word-document-root` for a missing main root.                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                    |
| `mce-extension-not-ignorable` | Office extension attributes lack an Ignorable declaration and are removed from the schema view. Warning from schema preprocessing; not emitted by the full Word-valid path.                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                      |

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

`unpack` checks every entry before it writes one (#675). It refuses, writing
nothing, an entry whose stored name leaves the output directory, even one JSZip
would load inside it (`../evil.txt` loads as `evil.txt`; #676), and an entry
whose destination runs through a symbolic link already in the output directory.
The output directory itself may be a link. When a read or write fails part way,
`unpack` removes the files and folders it created, including the output
directory when it created it; a file it overwrote keeps the new bytes. Messages
print control and bidirectional formatting characters in entry names as `\u{…}`
escapes.

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
`opc-part-name-escaped-unreserved`, `opc-part-name-escaped-slash`,
`opc-part-name-word-refused`, and `opc-item-name-refused` first, then `xml-syntax` and the rest; Word-valid
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
`package-extract-failed`, as the full run's extraction reports it (#413). Each stored
item name that places no part is also its own error-severity `opc-item-name-refused`
finding, and a full or partial run of a packed file still judges the archive's other
entries, as a full run does after a failed extraction: it leaves the refused
central-directory records out, so an entry a refused `..` name hid in the package view
is judged by its own bytes. A partial run judges them with its selected checks only.
Neither judges anything more when the remaining names still collide (#413), when the
archive uses ZIP64, or when bytes sit between the central directory and its end record,
and a directory input judges only its names
([#629](https://github.com/jandira-tech/docx-validate/issues/629)). Directory inputs apply the same stored-name refusal before either full or partial validation: literal backslash names with empty or dot-only segments are refused with the same code and message (#424).
Direct `DOCXSchemaValidator.validate()` and `PPTXSchemaValidator.validate()` calls also return an error-severity
`package-extract-failed` for such refused directory names, with the same message as `validate()`; programming errors still propagate
([#548](https://github.com/jandira-tech/docx-validate/issues/548)). They also apply the same directory walk refusal
for broken symbolic links, non-regular files, unlistable directories and symbolic-link loops under every profile,
before running any part-content check ([#557](https://github.com/jandira-tech/docx-validate/issues/557)). For either refusal,
direct validators include the same stored-name findings, in the same order and at the same profile-dependent severities,
as `validate()` for the same path string, and list the stored-name checks they ran in `executedChecks`, as `validate()`'s
`metadata.executedChecks` does ([#572](https://github.com/jandira-tech/docx-validate/issues/572)). Its stored
entry names and archive bytes are still read. Selected `opc-equivalent-part-names`,
`opc-part-name-escaped-unreserved`, `opc-part-name-escaped-slash`,
`opc-part-name-word-refused`, and `opc-item-name-refused` checks and
selected envelope checks (such as `zip-bytes-after-eocd`) run and report what the
full run reports. Every other selected check needs the package's parts: it runs
on the entries a refused item name leaves, as above, and is otherwise listed in
`omittedChecks`
([#420](https://github.com/jandira-tech/docx-validate/issues/420),
[#629](https://github.com/jandira-tech/docx-validate/issues/629)). It is not a
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
A packed file (`.docx`/`.docm`/`.pptx`) is different: auto-repair runs on a
temporary staged copy that is discarded, so the source file's bytes never
change. The result then carries an `auto-repair-staged-copy` info finding —
and the CLI prints the same sentence on its own line, after
`Auto-repaired N issue(s)` when anything was repaired —
disclosing that the repairs and the repair plan describe the staged copy
([#649](https://github.com/jandira-tech/docx-validate/issues/649)). It
appears only when a repair ran: a packed `.xlsx`, which has no validators,
and a packed file that cannot be extracted get none
([#666](https://github.com/jandira-tech/docx-validate/issues/666)).

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

With `autoRepair` the run also reports repair outcomes for DOCX/DOCM and PPTX
([#622](https://github.com/jandira-tech/docx-validate/issues/622)). The plan
appears in the path API's `validate()` result and in the CLI's `--json`
output. The CLI's text mode lists error-severity issues only; the one other
finding it prints is the `auto-repair-staged-copy` disclosure, on its own
line. The byte APIs (`new Validate().run(bytes)`,
`jubarte.validate(bytes)`) have no `autoRepair` option, so neither surface
carries the plan
([#651](https://github.com/jandira-tech/docx-validate/issues/651)). Non-info
findings from before repair are grouped by path and code, then compared with
the same group from that run's post-repair validation
([#648](https://github.com/jandira-tech/docx-validate/issues/648)).
Each info-severity plan line gives the original count and says
"Auto-repair cleared N; M remain": `repair-plan` means none remain,
`repair-plan-partial` means some cleared and some remain, and
`repair-plan-unavailable` means none cleared. Remaining findings carry
"Fix manually: …" when guidance exists — one line lists every distinct
guidance its survivors carry — or "No manual repair guidance is available."
otherwise. Counts describe group outcomes, including indirect
repairs; they do not identify individual occurrences or promise future repairs.
New post-repair groups remain ordinary validation findings. No repairs are added.

The paragraph-id repair stamps `w14:paraId`/`w14:textId` on the paragraphs and
table rows that lack them. On each part it stamps, it also declares `w14` on the
part root and lists it in the root's `mc:Ignorable`, as Word writes it — reusing
the prefix the root binds to the markup-compatibility namespace, else declaring
`mc` — so the repaired part gains no `mce-extension-not-ignorable` finding. A
root that binds `w14` or `mc` to another namespace is left as it was. Word 16.115
opened the repaired output clean for a root with neither declaration, one listing
`w15` and one using a `compat` prefix, as it opened the earlier output with
`xmlns:w14` on each stamped element (Word probes session 40, `s40a`;
[#653](https://github.com/jandira-tech/docx-validate/issues/653)).

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
