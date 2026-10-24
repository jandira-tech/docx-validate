# Migration guide

Notes for upgrading between breaking releases of `docx-validate`.
The project follows [SemVer](https://semver.org/spec/v2.0.0.html); under `0.x`,
minor versions carry breaking changes.

## Unreleased

### Breaking: DOCX/DOCM default to `word-valid`

Before this release, both `validate()` and the CLI defaulted to `lenient` through
`DEFAULT_PROFILE`. They now resolve the default by format: DOCX/DOCM use
`word-valid`; PPTX keeps `lenient` and its existing XSD validator. To preserve the previous DOCX policy, pass
`{ profile: "lenient" }` or `--profile lenient`. The same mode options apply to
`new Validate(options).run(bytes)` and `jubarte.validate(bytes, options)`; set
`suffix` on byte APIs when the input is not `.docx`.

For DOCX/DOCM, Word-valid uses structural Word checks without document XSD
validation. It can reject known
Word-opening defects previously accepted by lenient, and accepts some
schema-invalid constructs as warnings. For schema conformance, use
`{ strict: true }`, `{ profile: "strict" }`, or `--strict`. Word-valid always
returns `xsdWalk: { started: false, parts: [], eligible: 0 }` on the path API and
`metadata.schemaCompleted: false`; it no longer emits `xsd-strict-skipped` or
applies XSD message suppression or original-baseline subtraction. Full CLI
invocations still compile a small startup probe schema; this engine check is
outside the document trace. Help, discovery, and partial runs bypass it.

### Check selection and execution metadata

`quick: true` / `--quick` and `checks: ["tables"]` / `--checks tables` select
registered checks plus prerequisites, without XSD or the full validator's other
checks and severity remapping. Read `metadata.omittedChecks` before treating a
partial `valid: true` as a document verdict. `--list-checks` discovers ids and
groups. Strict cannot be combined with partial selection, quick cannot be
combined with named checks, and `--strict` cannot accompany a conflicting
profile. API conflicts throw `ModeConflictError`.

Results add optional `metadata` with `profile`, `executedChecks`, `omittedChecks`,
`schemaCompleted`, and requested `wordVerification`. The byte APIs return this
shared result. CLI `--json` prints `{ valid, issues, metadata }` for completed
runs, including warnings and locations, but omits the path API's `xsdWalk`,
`suffix`, and `repairs`. CLI configuration/startup errors still use stderr.

### Strict schemas, XML diagnostics, and coverage resolution

Strict now validates ISO Strict OOXML against bundled `iso-strict/` schemas;
lenient retains `xsd-strict-skipped`. Dublin Core imports are bundled locally.
Strict no longer subtracts `--original` schema errors or applies message
suppression, and repeated schema errors keep their own available line positions.
Messages identify positions in the preprocessed view when the view changed.
Schema/engine infrastructure failures are never removed by original-baseline
subtraction. Unmapped custom XML is not automatically a missing-schema error.

XML syntax uses original bytes without parser recovery, with UTF-8/UTF-16 support
and available 1-based line/column. Callers may now see `xml-undeclared-prefix` or
`xml-encoding` instead of generic `xml-syntax`. A BOM alone never causes failure.

Content-type checks now honor exact `Override` declarations before extension
`Default` declarations. Comparison uses canonical OPC keys: unreserved percent
escapes decode; reserved escapes remain distinct from their literal characters;
comparison is ASCII case-insensitive. Override-only images no longer produce
false missing-content-type errors.

Compatibility processing selects branches on a clone and preserves lax extension
payloads. Word rules judge the Word view rather than discarded branches. New
`mce-unbound-prefix` and `mce-invalid-token` findings are warnings in full
Word-valid and errors in lenient/strict schema processing. `mce-malformed` and
`mce-must-understand` remain Word-blocking. Loss of a required main root in the
schema view is `mce-main-root-ignored` (error), while an ignored non-main root
remains `mce-root-ignored` (info). See README for the complete current compatibility
code table and named Word-rule catalogue.

Full `.xlsx` validation now returns `unsupported-file-type` without library stderr
output; the CLI renders it. Unrecognized file suffixes still fail dispatch.

### Packaged CLI and optional Word verification

The package now ships `docx-validate` at `dist/cli.mjs`. Invoke that bin (or
`npx docx-validate`), rather than running `dist/index.mjs`, which is library-only
and starts no CLI. Source `tsx` script entrypoints still work.

Use `--verify-word` or `{ verifyWord: true }` to request an observed Word open.
The built-in verifier is macOS-only and requires installed Word and host-terminal
Accessibility automation permission. It opens a unique temporary copy of the
exact input bytes, observes the matching document, and closes without saving.
Only `clean-open` satisfies the request; all other outcomes make `valid` false,
even if portable checks pass and there are no error issues. Read
`metadata.wordVerification` for the outcome, file hash, platform, timestamp,
timing, details, and available Word version/build. Directories and unsupported
platforms return `unavailable`.

A pre-existing Word dialog now yields `unavailable` before any open/close, rather
than being misattributed to the input. Verification needs exclusive use of Word:
overlapping runs, or a user working in Word, can still have a dialog attributed to
the wrong file (#331). macOS cleanup dismisses the Word dialog showing at close; it
never quits or force-kills Word or relaunches it after a crash.

### Contributor coverage and fixture evidence

Run `pnpm run test:coverage-gate` for the full coverage suite and scoped module
thresholds (80% lines / 70% branches; 90% / 85% for policy, scheduling, registry,
and aggregation). Coverage-enabled runs print both percentages, including failed
runs. CI applies the same gate after tests.

Alongside `fixtures-all-{strict,lenient}`, `fixtures-all-word-valid.test.ts` pins
DOCX validity to Word observations bound to fixture hashes. Unknown, timeout, and
stale evidence does not count. Manifest `word` is now structured
`{ outcome, sha256, observedAt?, wordVersion? }`; use the provenance readers rather
than comparing it directly to an outcome string. Legacy strings remain readable
but cannot be checked for staleness. `test:fixtures:refresh` now uses
`update-manifest.ts`, preserving superseded observations in `history` and keeping
same-byte attributable evidence when a re-probe is inconclusive. Do not regenerate
expectations from validator output as proof of Word correctness.

## 0.53.0-rc.1

The low-level `validate()` result now requires `xsdWalk: { started, parts, eligible }`. Callers reading existing result fields keep working; typed test doubles or manually constructed `ValidateRunResult` objects must include the trace. `started: false, parts: [], eligible: 0` represents a path that did not enter XSD.

The `word-valid` profile adds checks grounded in recorded Word recovery cases, so documents previously accepted may now return blocking issue codes. Some markup-compatibility and undeclared-part edge cases remain under investigation; see `docs/really-broken/README.md`.

If upgrading from npm's prior `latest` version `0.5.0`, also read the `0.52.0` migration below: the intervening GitHub release changed the XSD engine and schema-load failure handling.

## 0.51.x → 0.52.0

### XSD engine: `libxmljs2` → `libxml2-wasm`

The native `libxmljs2` addon is gone. It ships no prebuild for Node 26
(ABI-147) and no longer compiles against Node's V8 headers, which made every
fresh install on current Node fail at the build step. The engine is now
[`libxml2-wasm`](https://github.com/albireox/libxml2-wasm) — same libxml2
semantics, no native compilation, no `node-gyp`.

- **No action** if you only call `validate()` / the validator classes with the
  default engine — the swap is internal to `createXsdValidator()`.
- **Action** if you injected a custom `XsdValidator`: nothing changes, the
  interface is identical. But note the default engine's new behaviours below.
- **Action** if your build scripts special-cased the `libxmljs2` postinstall
  (e.g. `allowBuilds` entries or `LD_LIBRARY_PATH` shims): remove them.

### `assertLibxmljsAvailable()` → `assertXsdValidationAvailable()`

The eager smoke-check on `BaseSchemaValidator` was renamed and is now
**async** (the wasm engine initialises asynchronously):

```ts
// before
BaseSchemaValidator.assertLibxmljsAvailable();

// after
await BaseSchemaValidator.assertXsdValidationAvailable();
```

The error prefix changed from `libxmljs2 required` to
`XSD engine unavailable (libxml2-wasm)`.

### Schema-load failures are errors, not info

A schema that cannot be loaded now produces
`severity: "error"`, `code: "xsd-schema-load-failed"`. Previously it produced
`severity: "info"`, `code: "xsd-schema-load-skipped"` and the document could
still validate as `valid: true`. A validator must not call a document valid
against a schema it could not load. At that release, the known-noisy remote
Dublin Core import was suppressed via
`IGNORED_VALIDATION_ERRORS`. The unreleased stack replaces it with local bundled
imports and removes that suppression; schema-load failures remain errors.

**Action**: tests or callers asserting on `xsd-schema-load-skipped`, or
relying on load failures being non-fatal, must switch to
`xsd-schema-load-failed` (error).

### New validation code: `ct-uncovered-part`

A residual OPC-coverage pass now flags parts matched by **no**
`<Override PartName>` and **no** `<Default Extension>` in
`[Content_Types].xml` (previously only declarable XML roots and known media
extensions were checked). Packages that quietly relied on the gap will now
fail with `ct-uncovered-part`. The current comparison uses canonical OPC part-name
keys as described above,
and treats only own-properties as known media extensions.

### New public export: `word-error-explanations`

`WORD_ERROR_EXPLANATIONS` / `explainWordError` are re-exported from the
package root, and the CLI prints a plain-language `↳ …` explanation under
each error when running with `--profile word-valid`. Additive; no action.

### Tooling (contributors)

- `pnpm` replaces `bun` as the package manager (`packageManager: pnpm@12.8.2`).
- `src/index.ts` is hand-maintained; the `barrel` / `barrel:check` scripts and
  the `barrelsby` dependency are gone (regeneration destroyed the curated
  surface). Edit the barrel by hand.
- Format settings live in `vite.config.ts` (`fmt` block); `.oxfmtrc.json` is
  gone. `pnpm run fmt` / `fmt:fix` now go through `vp fmt`.
- CI runs a Node 24 + 26 matrix.
