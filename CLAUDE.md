# Project: docx-validate

OOXML validators and redline/comment helpers for `.docx` and `.pptx` files.

- Repo: `jandira-tech/docx-validate`
- npm: `docx-validate` (unscoped)
- Language: TypeScript, ESM only
- Runtime: Node + Bun
- Build: `vite-plus`
- Tests: `vitest`
- Package manager: `pnpm`

## Code conventions

- File names: kebab-case (`merge-runs.ts`).
- ES modules only; no CommonJS.
- Relative imports between siblings; no path aliases.
- XSD validation: `libxml2-wasm` (via the `XsdValidator` interface in `src/lib/xsd-validator.ts`; the native `libxmljs2` addon was dropped — it ships no Node 26/ABI-147 prebuild and no longer compiles against Node's V8 headers).
- DOM work: `@xmldom/xmldom` + `xpath`, always via `src/lib/xml-helpers.ts`. Do NOT call `DOMParser` directly.
- Temp dirs: `tmp` via `withTempDir(async (dir) => { ... })` (`src/lib/run-cli.ts`). Don't ad-hoc `os.tmpdir()` + `mkdtempSync`.
- CLIs: `commander`, wired via `runCli(...)`.
- Validation results: return the `ValidationResult` shape from `src/lib/types.ts`. No per-validator shapes.
- Tests: one vitest file per module, fixtures under `tests/fixtures/` (`broken/` synthetic single-defect specimens, `working/` valid/repaired, `eigen/` real-world Plate/SuperDoc specimens, `external/` vendored third-party). Name fixtures `<subject>.<comment-or-error>.docx`; `scripts/apply-fixture-names.ts` derives names from content and `scripts/update-manifest.ts` regenerates `tests/fixtures-all.manifest.json` (which the `fixtures-all-{strict,lenient}` suites pin). `fixtures-all-word-valid.test.ts` instead pins validity against each attributable DOCX fixture's hash-bound Word observation; unknown, timeout, and stale observations are excluded, with floors of 358 clean opens and 67 rejections. The 24 jubarte lane specimens have separate integration coverage under `tests/lane/`.

## Schemas — golden source of truth

`src/scripts/office/schemas/` holds the official OOXML XSD schemas. Reference these whenever implementing or validating XML elements.

- `ISO-IEC29500-4_2016/wml.xsd` — WordprocessingML
- `ISO-IEC29500-4_2016/dml-main.xsd` — DrawingML
- `ISO-IEC29500-4_2016/shared-math.xsd` — Math
- `ecma/fourth-edition/opc-*.xsd` — Open Packaging Conventions
- `mce/mc.xsd` — Markup Compatibility Extensions
- `microsoft/wml-*.xsd`, `microsoft/word12.xsd` — Microsoft extensions

`defaultSchemasDir()` resolves to that directory.

## Repository layout

```
src/
  index.ts                  — public barrel
  lib/                      — xml-helpers, types, run-cli
  scripts/
    accept-changes.ts
    comment.ts
    office/
      pack.ts, unpack.ts, validate.ts, soffice.ts
      schemas/              — bundled XSDs (see above)
      helpers/              — merge-runs, simplify-redlines
      validators/           — base, docx, pptx, redlining
tests/                      — vitest specs + fixtures (broken/ working/ eigen/ external/)
```

## Behavioural notes

These are intentional behaviours of the implementation, documented at the call sites and verified by the vitest specs.

1. **Validator results are structured, not printed.** `validate_*` methods return `ValidationResult { valid, issues[] }`. The CLI shim in `validate.ts` is the only place that emits human-readable output. Tests assert on issue codes/messages.

2. **`@xmldom/xmldom` text-node mutation requires `.data`, not `.nodeValue`.** Validators that rewrite text nodes (whitespace repair, template-tag stripping) must use `(textNode as Text).data = …`. Writes to `.nodeValue` are silently dropped at serialization. See `validators/base.ts` near `_preprocessXmlForXsd`.

3. **`unpack`/`pack` byte parity is best-effort.** JSZip emits ZIP entries deflated with its own compression settings, so byte-for-byte round-trips against external tools are not guaranteed. XML pretty-printing/indentation is preserved so contents round-trip identically; the outer ZIP envelope can differ. This affects file-checksum equality, not logical-equivalence checks.

4. **`IGNORED_VALIDATION_ERRORS` contains only `"hyphenationZone"`.** `opc-coreProperties.xsd` imports the Dublin Core schemas from the bundled `schemas/dublincore/` copies (`dc.xsd`, `dcterms.xsd`, `dcmitype.xsd` — see `schemas/SOURCES.md`), so the former suppressions for the missing remote import (`"purl.org/dc/terms"`, `"dublincore.org"`, and the libxmljs2-era `"Invalid XSD schema"`) were removed. Schema-load failures surface as error-severity `xsd-schema-load-failed` issues — never silently valid and never message-suppressed. The remaining `"hyphenationZone"` entry is inherited from the Python validator with no recorded reason; full lenient validation still applies it. Strict never suppresses schema findings through the list, and full DOCX/DOCM Word-valid runs no XSD, so it does not consult the list or emit `xsd-strict-skipped`.

5. **`compareParagraphCounts` returns a structured result.** `{ original, modified, delta, originalUsesStrictNamespace }`. With `verbose: true`, also prints `Paragraphs: N → M (+K)` to stdout before returning.

6. **XML diagnostics retain available source positions.** `checkXmlSyntax` parses original bytes without recovery and emits `xml-syntax`, `xml-undeclared-prefix`, or `xml-encoding` with the part path and available 1-based line/column. DOM decoding supports UTF-8 and UTF-16 via BOM/declaration. Strict schema findings retain available engine lines; when preprocessing changes the XML, the message identifies lines in the transformed validation view. Lenient keeps its capped, deduplicated XSD message path. External entities are not loaded or expanded, and XSD resources resolve through allowed local roots.

7. **`accept-changes.ts` wraps `runSoffice()` in `runSofficeWithTimeout(args, 30000)`.** Node's `child_process.spawn` has no built-in timeout that resolves cleanly with stdout/stderr, so the wrapper is a `setTimeout`-driven Promise.

8. **`validateWhitespacePreservation` uses `[ \t\n\r]`, not `\s`.** JavaScript's `\s` is broader (matches `\f`, `\v`, Unicode whitespace), so the explicit character class is used to keep the match set predictable.

9. **`RedliningValidator` class wraps free functions.** `RedliningValidator.validate()` and `.repair()` delegate to the free-function implementation. `repair()` returns `0`.

10. **Full `.xlsx` validation returns a silent structured error.** `runValidators` returns `code: "unsupported-file-type"` without writing to stderr; the CLI renders it. Other unrecognized suffixes fail target dispatch before this point. Partial selection can still run registered package/XML checks on XLSX.

11. **`repairWhitespacePreservation` prints a per-repair line when `verbose: true`.** `  Repaired: <file>: Added xml:space='preserve' to <tag>: <preview>` per element fixed.

12. **Namespace constants exported from `validators/docx.ts`:** `WORD_2006_NAMESPACE`, `WORD_STRICT_NAMESPACE`, and `WORD_PARAGRAPH_NAMESPACES` (a `readonly [string, string]` tuple).

13. **`validateNamespaces` walks all descendants for `xmlns:*` collection.** This mirrors lxml's `nsmap` propagation: namespace prefixes declared on child elements (not just the root) are collected.

14. **`pack()` accepts an `inferAuthorFunc` callback.** `inferAuthorFunc?: (unpackedDir: string, originalDocx: string) => Promise<string> | string`. When omitted, the default `inferAuthor` from `simplify-redlines.ts` is used.

15. **`validate()` exposes `xsdWalk` and optional `metadata`.** DOCX/DOCM default to `word-valid`; PPTX keeps `lenient` and its existing XSD validator, which walks the schemas under every profile, `word-valid` included, so a PPTX can report `schemaCompleted: true` under any profile (#340). For DOCX/DOCM, Word-valid document validation runs no XSD: `xsdWalk` is `{ started: false, parts: [], eligible: 0 }` and `metadata.schemaCompleted` is always false. Full lenient and strict runs walk eligible parts (lenient can return before XSD on malformed XML). `metadata` reports `profile`, the registry ids the run executed in `executedChecks` (#319), selection omissions in `omittedChecks`, `schemaCompleted`, and optional `wordVerification`. The byte APIs forward mode and verifier options and return metadata on the shared `ValidationResult`.

16. **Packaged CLI: `dist/cli.mjs` is the only live shim in the bundle.** `src/cli.ts` is the `docx-validate` bin (package.json `bin`), a second `pack.entry` that calls `runCli(import.meta.url, ...)`. The per-script `runCli` shims stay in `src/scripts/**` for `tsx` execution, but `vp pack`'s `inertCliShimsPlugin` rewrites their metaUrl to a never-matching `file:` URL — bundled modules share one `import.meta.url`, so a shim without the rewrite would fire on `node dist/index.mjs` and import (#314). New CLI entry points belong in `pack.entry` + `bin`, never as `runCli` calls the barrel can reach.

17. **Named Word rules run through the registry and scheduler.** `src/scripts/office/checks/index.ts` registers `WORD_RULE_CHECKS` for full `word-valid` and `strict` DOCX runs; lenient keeps the other validators. Check definitions declare `requires` and registration order puts prerequisites first. `xml-syntax` gates by part: malformed parts are hidden from dependent DOM rules while well-formed parts are still judged. Prerequisites run once; a failed gate with no part-scoped errors skips its dependents. `runChecks()` exposes per-check outcomes, including `unavailable` for archive-only checks on directories; top-level metadata is not that status trace.

18. **CLI and API selectors share `mode-policy.ts`.** `--strict` / `strict: true` selects the strict profile. `--quick` / `quick: true` selects `QUICK_CHECK_GROUPS` (`package`, `xml`, `content-types`, `relationships`, `tables`, `drawings`). `--checks` / `checks` selects registered ids or groups plus prerequisites, without the full validator's other methods. A partial run selecting a check that reports its own parse failures also runs `xml-syntax` as a companion without gating that check, and suppresses restated parse findings for rejected parts (#385); under lenient it stops where `xml-syntax` fails, as the full lenient run stops at a malformed part (#391). A partial run applies its profile as a full run does, so it never fails a package the full run passes (#381). Under lenient it runs neither the Word rules nor the ZIP-envelope checks: naming one, or only groups made of them, is a `ModeConflictError`. Under word-valid it reports a DOCX check's error that Word tolerates as a warning, through the `isWordBlockingIssue` predicate in `validators/word-blocking.ts` that the full run uses. Strict rejects partial selection and conflicting profiles; quick and named selection are mutually exclusive. API conflicts throw `ModeConflictError`. `--list-checks` and `--help` need no package or engine startup. Full CLI runs that walk schemas (`profileWalksSchemas(format, profile)`: DOCX/DOCM under lenient and strict, PPTX under every profile) compile a small XSD startup probe outside the document `xsdWalk`; Word-valid DOCX/DOCM, XLSX, and partial runs bypass it (#318, #341). Partial runs compile no XSD, list omitted check ids, name what has no id in an `Also not run:` line (the format validator's own rules, where a full run walks schemas the schema walk, and for a DOCX given `--original` the redlining validation, #353), and print `Selected checks passed`, never `All validations PASSED!`, even when no registered check was omitted (#339); `--json` renders `{ valid, issues, metadata }` for completed runs. See README for the complete named-check catalogue and advisory rules.

19. **Real-Word verification is optional and separate from portable success.** `--verify-word` / `verifyWord: true` uses `src/lib/word-verifier.ts` (or an injected `wordVerifier` callback). The built-in transport needs macOS, installed Microsoft Word, and host-terminal Accessibility automation permission. It opens an exact-byte temporary copy, requires the matching document to stay dialog-free for 3 seconds, and uses a default 45-second launch/observation/cleanup budget with up to 2 seconds reserved for close. Only `clean-open` satisfies verification; `repair-warning`, `open-error`, `crash`, `password-prompt`, `unknown-dialog`, `timeout`, and `unavailable` do not. Metadata records the hash, platform, wall-clock timestamp (`wallNow`, never the deadline clock; #329), elapsed observation, details, and available Word version/build. Requested-but-unsatisfied verification makes `valid` false even without error issues. Directories have no packed bytes and are unavailable. `crash` requires a Microsoft Word crash report written after this run's open; a vanished Word without one (an external kill or quit) is `unavailable` (#305). The version and availability calls race the same budget (#320). A verifier that throws yields `unavailable` instead of rejecting `validate` (#333). A dialog already showing before open makes the run unavailable without opening or closing anything (#315). Verification needs exclusive use of Word: runs must not overlap and Word must not be in use, because a dialog that appears after the pre-open observation is attributed to this file (#331). macOS close dismisses the Word dialog showing at close (Cancel, No, OK), which under exclusive use is the one this run's open raised, lets an open that overran its budget settle before closing (#328), closes only its copy without saving, and never quits or force-kills Word or relaunches it after a crash.

20. **Manifest Word evidence is hash-bound.** `src/lib/word-provenance.ts` reads `word: { outcome, sha256, observedAt?, wordVersion? }`; a hash mismatch is `unknown`. Legacy strings remain readable but have no hash to check. Unknown, timeout, crash and stale observations, and unmapped outcome spellings, do not count toward clean/rejected accuracy; only `open-error`, `unreadable-content-warning`, `password-required` and `unknown-dialog` are rejections, and `toManifestOutcome` maps verifier outcomes to these spellings (#316). `mergeWordObservation` retires superseded observations into `history` and preserves attributable same-byte evidence against timeout/unknown re-probes. `test:fixtures:refresh` uses the provenance-preserving `scripts/update-manifest.ts`; review expectation changes individually rather than treating current validator output as Word evidence.

## ISO OOXML Strict XSD validation

ISO OOXML defines two conformance classes: **Transitional** (backward-compatible, widely used) and **Strict** (pure ISO, uses different namespace URIs under `purl.oclc.org/ooxml/` instead of `schemas.openxmlformats.org/…`). Example: Strict WML uses `http://purl.oclc.org/ooxml/wordprocessingml/main` vs Transitional's `http://schemas.openxmlformats.org/wordprocessingml/2006/main`.

The official ISO Strict schemas are bundled under `src/scripts/office/schemas/iso-strict/` (ECMA-376 Part 1, 5th edition — see `schemas/SOURCES.md` for provenance, sha256s, and the two local fixes needed to make them compile: the xml.xsd import and three `default="off"` → `default="false"` attribute defaults in wml.xsd).

**Current behaviour:** `validateAgainstXsd` in `validators/base.ts` detects Strict conformance by checking each XML file's root element `namespaceURI` against the `http://purl.oclc.org/ooxml/` prefix (`_isStrictXmlFile`). Under `profile: "strict"`, Strict parts dispatch to the matching iso-strict schema by part identity (`STRICT_SCHEMA_MAPPINGS`); OPC parts (content types, rels, core properties) share one namespace across both conformance classes and always use the OPC schemas, and Strict parts with no mapped schema are skipped like any unmapped part. Under `lenient`, Strict parts keep the historical `info`-severity `xsd-strict-skipped` skip. Full DOCX/DOCM Word-valid runs no XSD for either conformance class and emits no such skip issue. Markup compatibility runs before the selected schema validates the part; for a Strict root it treats the `purl.oclc.org/ooxml/` counterparts of the understood Transitional namespaces as understood (`STRICT_UNDERSTOOD_NAMESPACES`) and keeps lax payloads in the Strict extension slots.

Also under `strict`: baseline subtraction against `--original` is bypassed (a schema defect is invalid even if the original had it — lenient keeps the XSD diff; DOCX/DOCM Word-valid runs no XSD), and every schema violation is emitted as its own `xsd-error` issue with the reported `line` (infrastructure findings retain codes such as `xsd-schema-load-failed`); when preprocessing rewrote the XML for validation, the issue message marks the line as referring to the transformed view.

The `STRICT_OOXML_NS_PREFIX` constant in `validators/base.ts` and the `WORD_STRICT_NAMESPACE` constant in `validators/docx.ts` are the canonical sources for these URIs.

The DOCX/DOCM Word-valid behavior described here does not change the PPTX pipeline. The Word and schema compatibility views use separate namespace capabilities. Word rules read `ValidationContext.effectiveXml()` after branch selection; raw-markup rules may read `xml()`. Unselected branches are not judged by structural Word rules. The schema processor preserves lax extension payloads and applies inherited compatibility directives on a clone. `mce-malformed` and `mce-must-understand` are Word-blocking; `mce-unbound-prefix` and `mce-invalid-token` are errors in lenient/strict schema processing but warnings in full Word-valid. Root Ignorable duplicates collapse to `ignorable-undeclared`. The schema-only `mce-root-ignored` is informational except for required main roots (`mce-main-root-ignored`, error); `mce-extension-not-ignorable` is a schema-preprocessing warning. Neither schema-only finding is emitted by full Word-valid.

Content-type checks share `src/lib/content-types.ts`: exact `Override` precedes extension `Default`, using canonical OPC part-name keys. Comparison decodes unreserved escapes, keeps reserved escapes distinct from literals, encodes other path characters, and compares ASCII case-insensitively. An Override-only image is covered; an actually uncovered part still fails with `ct-uncovered-part`.

## Known gaps

- **Of equivalent part names, only one entry is judged (#351).** `opc-equivalent-part-names` reports equivalent stored part names and multiple content-types stream spellings: an error under strict and under word-valid for DOCX/DOCM (Word 16.115 offers recovery for every duplicate spelling probed, `docs/plans/evidence/word-probes-2026-10-24.md`), a warning under lenient. The content of the dropped entries is not validated. The package view still prefers an exact normalized part-name match, otherwise the first OPC-equivalent entry in code-unit order; slash/backslash aliases keep the normalized spelling, otherwise the first stored spelling in code-unit order. The content-types stream keeps `[Content_Types].xml` when present, otherwise the first matching spelling in code-unit order. On a file system insensitive to case or Unicode normalization, full extraction still stops at `package-extract-failed` for collisions (#373, #376); the diagnosis runs on the archive names before extraction.

- **soffice end-to-end tests are gated on `process.env.SOFFICE_AVAILABLE`.** Both `tests/accept-changes.test.ts` and `tests/soffice.test.ts` skip real-LibreOffice invocations unless that env var is set; CI images typically don't include LibreOffice. Run with `SOFFICE_AVAILABLE=1 pnpm test` locally to exercise them.

- **Three `eslint-plugin-functional` rules are disabled in `.oxlintrc.json`:** `functional/immutable-data`, `functional/no-mixed-types`, `functional/prefer-readonly-type`. All three are type-aware and call `getParserServices()`, but oxlint's JS-plugin runtime hard-codes `parserServices: ObjectFreeze({})` (see `node_modules/oxlint/dist/lint.js:13580`), so they crash the AST walker on any file containing a type alias / interface or an `arr.push(...)` call. `prefer-property-signatures` is the only functional rule left enabled because it's syntactic-only. Re-enable the three when oxlint either provides parser services to JS plugins or honors `requiresTypeChecking: true` for skip.

## Verification

```bash
pnpm install
pnpm exec tsc --noEmit  # type-check
pnpm test --coverage    # vitest + line/branch reports
pnpm run test:coverage-gate  # full suite + per-module thresholds
pnpm run check          # vite-plus check
pnpm run build          # produce dist/
```

Coverage-enabled Vitest runs, including failed runs, print `Coverage: N% lines, M% branches` and write text, JSON-summary, HTML, and LCOV reports. `scripts/coverage-gate.ts` reads `coverage/coverage-summary.json`: the listed new modules need 80% lines / 70% branches; mode policy, scheduler, registry, and result aggregation need 90% / 85%. A gate with no matching coverage fails. CI runs the gate after full coverage tests; a focused run is not a substitute.

The installed-package smoke is opt-in: `PACKAGE_SMOKE=1 pnpm exec vitest run tests/package-artifact.integration.test.ts --coverage`. It checks consumer imports/validation, bundled Transitional/ISO Strict/Dublin Core schema assets, one packaged CLI help usage, and a silent library entry. The separate real-Word lane is `WORD_VERIFY=1 pnpm exec vitest run tests/word-verifier.word.test.ts --coverage` on macOS with Word and host-terminal Accessibility permission; portable tests do not make fresh Word observations.

## Tooling preferences

- Frontend tooling preference order: pnpm, biome, elysia (over npm/yarn/eslint/prettier).
- Favor Rust-backed tools.
- Async-first when writing supporting Python (uv, httpx, uvloop, pydantic, polars, pydantic_ai).

## Notes

- The vendored `docx/` folder is a third-party library snapshot (the `docx` npm package). Its `CLAUDE.md` is unrelated to this project — don't treat it as authoritative.
- `docx-templates/` is similarly a third-party reference snapshot.

## Desired Conversion Methodology

**Desired Conversion Methodology** means fully parsing the original document into a semantic, format-independent representation — extracting its content and the meaning of that content (including its structure, formatting, and relationships, and every other discrete constituent component of the document (each, an "**Element**")), and not its byte form, such that the document can be faithfully reconstructed into any target format.

Where an Element cannot be parsed semantically by available means — including, by default, DrawingML (`<w:drawing>`), VML pictures (`<w:pict>`), and embedded OLE objects (`<w:object>`) (each, a "**Non-Parseable Element**") — the Desired Conversion Methodology nonetheless applies in full to the entire document and to every Element enclosing that Non-Parseable Element; only the single, irreducible Element that resists semantic parsing may be preserved verbatim, and only after custom parsing techniques have been exhausted to reduce that residue to the smallest possible extent.

For the avoidance of doubt, the Elements enumerated above are not the sole Non-Parseable Elements; every time an Element is classified as a Non-Parseable Element, such classification must (i) emit a warning during code execution (at runtime), and (ii) be documented in `AGENTS.md`, `README.md` and `CLAUDE.md`.
