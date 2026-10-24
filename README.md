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
can stop a lenient run before XSD. The trace describes the schema walk, not a
Microsoft Word invocation. Under `strict`, ISO Strict OOXML parts use the bundled
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
- `--list-checks` lists every registered id, group, format, and description without
  reading a file or starting XML/XSD engines. `--help` also needs no engine startup.
- `--strict` / `{ strict: true }` is shorthand for `profile: "strict"`. Strict
  rejects partial selection; `--quick` and `--checks` are mutually exclusive.
  `--strict` also rejects an explicit profile other than `strict`. API conflicts
  throw `ModeConflictError`; CLI configuration errors exit 1.
- Partial runs perform no XSD validation and report `omittedChecks`. A successful
  partial run with omissions prints “Selected checks passed”, followed by counts.
  It does not establish full-document validity. Selections run the registered
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

The complete named Word check catalogue is below. Except for
`word-markup-compatibility`, each id is also its finding code. Package checks are
`zip-bytes-after-eocd` and `zip-eocd-comment-overrun`; `xml-syntax` is the shared
XML prerequisite. All `word-*` checks apply to DOCX/DOCM only.

| Group                     | Named Word checks                                                                                                                                                                                                                                                                                                                                   |
| ------------------------- | --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `math`, `tracked-changes` | `word-math-run-insertion`, `word-math-run-deletion`                                                                                                                                                                                                                                                                                                 |
| `math`                    | `word-math-parse`, `word-math-spre-body`                                                                                                                                                                                                                                                                                                            |
| `story`                   | `word-shading-color`, `word-page-margin-attributes`, `word-run-misplaced`, `word-run-properties-misplaced`, `word-underline-misplaced`, `word-paragraph-properties-misplaced`, `word-section-properties-misplaced`, `word-footnote-reference-missing`, `word-numbering-picture-bullet`, `word-sdt-id-empty`, `word-text-contains-paragraph`         |
| `tables`                  | `word-grid-column-attributes`, `word-cell-margin-order`, `word-table-properties-required`, `word-table-properties-misplaced`, `word-row-without-cell`, `word-cell-ends-with-table`                                                                                                                                                                  |
| `drawings`                | `word-textbox-content-namespace`, `word-drawing-extent-missing`, `word-drawing-extent-misplaced`, `word-picture-nvpicpr-missing`, `word-picture-sppr-missing`, `word-empty-drawing`, `word-anchor-missing-attribute`, `word-position-offset-and-align`, `word-drawing-scalar-whitespace`, `word-drawing-arc-radius`, `word-drawing-extent-negative` |
| `content-types`           | `word-content-type-invalid`                                                                                                                                                                                                                                                                                                                         |
| `relationships`           | `word-duplicate-theme-relationship`, `word-duplicate-font-table-relationship`, `word-duplicate-web-settings-relationship`, `word-duplicate-settings-relationship`, `word-hyperlink-relationship`                                                                                                                                                    |
| `tracked-changes`         | `word-textbox-deleted-text`, `word-move-from-deleted-text`                                                                                                                                                                                                                                                                                          |
| `settings`                | `word-custom-property-filetime-empty`, `word-attached-toolbar-utf8`                                                                                                                                                                                                                                                                                 |
| `xml`                     | `word-markup-compatibility`, `word-document-root`, `word-choice-requires-undeclared`                                                                                                                                                                                                                                                                |
| `schema`                  | `word-unknown-element`, `word-link-without-namespace`                                                                                                                                                                                                                                                                                               |

The `schema` group contains structural checks, not an XSD walk. `comments` is a
recognized group with no registered checks yet; selecting it alone runs nothing
and reports all applicable ids as omitted.

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
original part path and available 1-based line/column. External entities are not
loaded or expanded; XSD imports resolve through allowed local schema roots.
Validation leaves source bytes unchanged unless you request `autoRepair`.

Content-type coverage resolves an exact `Override` before an extension `Default`.
An image covered only by `Override` needs no matching `Default`. Part-name
comparison is ASCII case-insensitive and decodes unreserved percent escapes;
reserved escapes stay distinct (`%40` differs from `@`), while spaces and other
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
other platforms or missing Word also return `unavailable`.

Before opening, the verifier observes Word once. If a dialog is already showing,
it returns `unavailable` without opening or closing anything: that dialog cannot
be attributed to this file. Verification needs exclusive use of Word: runs must
not overlap and Word must not be in use, because a dialog that appears after the
pre-open observation is attributed to this file (#331). macOS cleanup dismisses
the Word dialog showing at close, preferring Cancel, No, then OK; under exclusive
use that is the dialog this run's open raised. It does not relaunch Word after a
crash.
The source implementation (`src/lib/word-verifier.ts`) supports injected transport
and clock dependencies; public validation options accept a `wordVerifier` callback
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
registry ids a complete run executed (#319): lenient lists `xml-syntax`; Word-valid
and strict add the ZIP-envelope checks for a packed file and every Word rule whose
prerequisites passed. A partial run lists its resolved selection. It is not a
per-method execution trace. In the source scheduler, use `runChecks().outcomes`
for individual
`passed` / `failed` / `skipped` / `unavailable` statuses. `omittedChecks` names ids
left out by selection. `schemaCompleted` records completion of the eligible schema
walk without schema-load failures; it is separate from whether schema findings
made the document invalid. `wordVerification` is present only when requested.

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
and stale observations are excluded; floors of 358 clean opens and 67 rejections
prevent silent loss of evidence. The 24 jubarte 0.13.0 specimens live separately
under `tests/lane/` with their own integration suite.

Manifest `word` observations use `{ outcome, sha256, observedAt?, wordVersion? }`.
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
