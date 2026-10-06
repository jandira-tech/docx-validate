# Changelog

All notable changes to `docx-validate` are documented here. Format follows
[Keep a Changelog](https://keepachangelog.com/en/1.1.0/) and the project
adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

## [0.52.0] — 2026-10-06

Breaking under `0.x` conventions: the native `libxmljs2` addon is replaced by
`libxml2-wasm` and schema-load failures change severity — see `MIGRATION.md`.

### Added

- **`word-error-explanations` public module.** `WORD_ERROR_EXPLANATIONS` / `explainWordError` — plain-language, real-Word-probing-grounded explanations for the 19 issue codes the `word-valid` profile treats as Word-blocking. The CLI prints them as a `↳ …` line under each error when running `--profile word-valid`. Re-exported from the package root.
- **`ct-uncovered-part` OPC coverage check.** Parts matched by no `<Override PartName>` and no `<Default Extension>` in `[Content_Types].xml` are now flagged (previously only declarable XML roots and known media extensions were checked — extensionless media slipped through). `PartName`s are URI-decoded before comparison, and media-extension matching uses own-property checks only (`x.constructor` is not a media type).
- **`getElementsByTagNameAll`** — stack-based DFS walker replacing quadratic live-`NodeList` iteration for wildcard collection; `getElementsByTagNameNSAll` now uses the same O(N) traversal (one shared implementation, two entry points). New `tests/xml-helpers.test.ts` pins the contract, including the native root semantics (Document root ⇒ `documentElement` included; Element root ⇒ self excluded).
- **`MIGRATION.md`** with per-version breaking-change notes.
- CI runs a Node 24 + 26 matrix.

### Changed

- **XSD engine: `libxmljs2` → `libxml2-wasm`.** The native addon ships no Node 26 / ABI-147 prebuild and no longer compiles against Node's V8 headers; the wasm build has identical libxml2 semantics with no native compilation step. Compiled schema validators are cached per schema path (the cache stores in-flight load promises, so concurrent first loads compile exactly one validator; failed loads evict and retry; source documents are disposed even when compilation throws).
- **Schema-load failures are errors, not info.** Unloadable schemas now produce `severity: "error"`, `code: "xsd-schema-load-failed"` (was `info` / `xsd-schema-load-skipped`). A validator must not call a document valid against a schema it could not load. The known-noisy Dublin Core import in `opc-coreProperties.xsd` stays suppressed via `IGNORED_VALIDATION_ERRORS`.
- `BaseSchemaValidator.assertLibxmljsAvailable()` → **async** `assertXsdValidationAvailable()` (wasm initialises asynchronously); error prefix is now `XSD engine unavailable (libxml2-wasm)`.
- **Toolchain: `pnpm` replaces `bun`** (`packageManager: pnpm@12.8.2`); all dependencies updated to latest.
- **`src/index.ts` is hand-maintained.** barrelsby could not reproduce the curated named exports and silently destroyed them on regeneration; the `barrel` / `barrel:check` scripts are gone and the barrel is edited by hand.
- **Format settings moved** from `.oxfmtrc.json` into `vite.config.ts`'s `fmt` block so `vp check` and `pnpm run fmt` cannot desync; `fmt` / `fmt:fix` now go through `vp fmt`.

### Fixed

- **CI workflows actually run again.** `pnpm/action-setup` now executes before `actions/setup-node`'s `cache: pnpm` resolution (the old order failed every pnpm job with "Unable to locate executable file: pnpm" — CI had not run green on `main` since 2026-06-02). The validator-diff workflow resolves pnpm via `package_json_file` (its path-scoped checkouts leave no root `package.json`).
- Validator performance on large documents: del-first `collectDeletedRunText` with Set dedup across both Word namespaces, native-DOM `validateDeletions` (no `xpath` descendant queries), O(N) paragraph counting that skips text-box subtrees during traversal, single-pass redlining stripping.

### Security

- **Zip Slip** blocked at extraction (`resolveSafeZipEntry`: `path.relative` containment, prefix-slip-safe).
- **Path traversal in relationship targets** contained lexically and via realpath-based symlink-escape detection (`containRelationshipTarget`), for both absolute and relative targets.
- **Document IDs use `crypto.randomInt`** instead of predictable `Math.random`-derived identifiers.

### Removed

- `libxmljs2` (runtime + `allowBuilds` entry), `barrelsby` (dev), `.oxfmtrc.json`, the `barrel` / `barrel:check` scripts.

## [0.51.0] — 2026-05-30

### Changed

- Version bump to `0.51.0`.

## [0.5.0] — 2026-05-27

### Added

- **Content-descriptive fixture corpus.** Every undescriptive `.docx` test fixture is now named `<subject>.<comment-or-error>.docx`, derived deterministically from its content. New tooling under `scripts/`: `fixture-fingerprint.ts` (JSZip + `validate()` content fingerprint), `derive-fixture-name.ts` (pure name derivation with `error-first` and `content-first` descriptor modes), and `apply-fixture-names.ts` (rename/sort driver with content-hash dedup, collision disambiguation, and `--into-categories` / `--category` / `--descriptor` flags). All three are unit-tested.
- **`tests/fixtures/eigen/`** — 148 real-world Plate/SuperDoc specimens (deduped from 151) imported from the former unreferenced top-level `fixtures/eigen-extended/`, renamed by content (`content-first`: distinguishing feature, falling back to the validation error code). Provenance in `tests/fixtures/eigen/README.md`. The `fixtures-all-{strict,lenient}` suites now pin all of them (`tests/fixtures-all.manifest.json`: 416 → 564 fixtures).

### Changed

- **`scripts/update-manifest.ts` preserves `word` metadata across regens.** When the Word-probe JSONL is absent (no LibreOffice in CI), the per-fixture `word` outcome now falls back to the previous manifest's value instead of being downgraded to `"unknown"`. Exposed as the testable `resolveWordOutcome` helper.
- Renamed three undescriptive `tests/fixtures/` strays (a random-named root specimen, `Ouch.docx`, and a path-as-filename artifact) to content-derived names.

### Removed

- `Sample Document.repaired.docx` (a tracked, regenerable repair artifact) and the stray npm `package-lock.json` (this is a bun-only repo) from the repo root.
- The dead top-level `fixtures/` directory (nothing referenced it).

## [0.1.3] — 2026-05-04

### Added

- **TypeScript type declarations now ship in the package.** `vp pack` emits `dist/index.d.mts` (27.87 kB / 9.37 kB gzipped) alongside `dist/index.mjs`. `package.json` exports the types via the `exports."."` conditions (`types` listed first per TS bundler-resolution rules) and a top-level `types` fallback. Consumers of every public re-export from `src/index.ts` — `validate`, the validator classes, side helpers, and the `XsdValidationOutcome` / `ParagraphCounts` / `RedliningOptions` interfaces — now get full IntelliSense.
- **`"sideEffects": false`** in `package.json` so bundlers can tree-shake the package aggressively. The barrel only re-exports pure functions, classes, types, and namespace constants — no top-level side effects, so the claim is safe.
- **publint and attw build-time gates** wired into `vite.config.ts` (`pack.publint: true`, `pack.attw: true`). publint fails the build on `package.json` shape violations (`exports` map order, missing files); attw verifies type resolution under `node10` / `node16` / `bundler` modes. The `attw.profile: "esm-only"` field documents that the `cjs-resolves-to-esm` warning is intentional — this is an ESM-only package by design (CLAUDE.md "ES modules only; no CommonJS").
- **Test coverage in CI.** `@vitest/coverage-v8` devDep + `test.coverage` block in `vite.config.ts` (text + html + lcov reporters). `ci.yml` runs `vitest run --coverage` and uploads `coverage/lcov.info` via `codecov/codecov-action@v5` (token sourced from `secrets.CODECOV_TOKEN`).
- **CodeQL security scanning** (`.github/workflows/codeql.yml`). Uses `javascript-typescript` language with `build-mode: none`, scoped to first-party source via `paths-ignore` (skips fixtures, vendored snapshots, build output, docs). Runs on push/PR to `main` only — no schedule.
- README badges: CI status, CodeQL, Codecov, Snyk Known Vulnerabilities, npm version + downloads + types, bundle size (min + gzip), tree-shakeable, dependency count, license, node version, GitHub repo.
- `llms.txt` follows [llmstxt.org](https://llmstxt.org/) format — title + summary + Docs / API / Tests / Optional sections with deep links into every public module.
- Public-surface exports for `XsdValidationOutcome`, `ParagraphCounts`, and `RedliningOptions` (previously declared `interface` without `export`, so consumers couldn't reference them by name).
- `knip.json` with `ignoreDependencies` for plugins knip can't statically resolve through oxlint's `jsPlugins[]` config (`barrelsby`, `eslint-plugin-functional`, `eslint-plugin-jsdoc`, `eslint-plugin-prefer-arrow`).
- `coverage/` added to `.gitignore`.
- `bun run release:publish` — single-step `git push --follow-tags && npm publish --access public`. Uses local npm credentials (`npm login` / `NPM_TOKEN`); the `publish.yml` Trusted-Publishing OIDC flow remains as a backup if you'd rather publish from CI. Named `release:publish` (not `publish`) to avoid the recursive lifecycle trigger when `npm publish` runs the `publish` script.
- Maintainer line in README pointing at [jandira.tech](https://www.jandira.tech) — Jandira Technologies + Cicero context.

### Changed

- `package.json` `description`: trailing typo fix (`"ESM. for the neurotic developer"` → `"ESM, for the neurotic developer."`).
- README header rewritten with badge row; summary tightened.
- typedoc no longer surfaces stale TS errors — unused imports (`mergeResults` in `validators/base.ts`, `serializeXml` in `validators/redlining.ts`) cleared.

### Internal

- cspell dictionary: added `barrelsby`, `llms`.
- `defaultSchemasDir()` in `validators/base.ts` falls back to both layouts (`../schemas` for the source tree, `./schemas` for the published `dist/` bundle) so the bundled package resolves XSDs against the copy emitted by the `copySchemasPlugin` in `vite.config.ts`.

## [0.1.2] — 2026-05-04

### Security

- **LD_PRELOAD shim removed from the published bundle.** The
  `acceptChanges()` LibreOffice flow uses a runtime-compiled C shim that
  `LD_PRELOAD`s into `soffice` to work around `AF_UNIX` restrictions in
  sandboxed VMs. That pattern legitimately matches supply-chain malware
  heuristics in automated scanners (Socket.dev, etc.) — and they're
  correct to flag _unknown_ packages doing this. The shim source remains
  in `src/scripts/office/soffice.ts` for developers who run the CLI
  directly, but `src/index.ts` no longer re-exports `soffice.ts` or
  `accept-changes.ts`, so the bundler tree-shakes both files (and
  `SHIM_SOURCE`, `LD_PRELOAD`, `RTLD_NEXT`, `dlsym`, etc.) out of
  `dist/index.mjs`. `grep -E "LD_PRELOAD|SHIM_SOURCE|RTLD_NEXT|dlsym"
dist/index.mjs` returns zero matches.
- Added `SECURITY.md` with the disclosure policy and a detailed
  explanation of the shim pattern, where it lives, and how to verify the
  published bundle is clean.

### Added

- **`bun run release` checklist wrapper** (`scripts/release.ts`) that
  reads `.release-checklist.md`, prompts `[y/N]` for each unchecked
  item, and aborts on the first "no" before forwarding to `bumpp`. The
  checklist itself lives in `.release-checklist.md` and isn't hard-coded
  — projects edit the file to match their actual workflow. `--yes` skips
  prompts for CI.
- `CHANGELOG.md` (this file).
- `llms.txt` package summary for LLM-friendly tooling.
- CodeQL workflow (`.github/workflows/codeql.yml`).
- Codecov coverage upload step in `ci.yml`.
- README badges for CI / CodeQL / Codecov / npm / bundle size / types.

### Changed

- README "Programmatic use" section updated with a clearly-marked "Not
  in the package surface: LibreOffice helpers" subsection pointing at
  SECURITY.md and the source-checkout path for callers that need
  `acceptChanges()` / `runSoffice()`.
- `llms.txt` synced with the post-shim-removal API surface.

### Internal

- `package.json` `barrel` script now passes `--exclude` patterns for
  `scripts/office/soffice\.ts$` and `scripts/accept-changes\.ts$` so
  regenerating the barrel preserves the surface trim.

## [0.1.1] — earlier

Initial publishable release. See `git log --oneline` for the full set of
commits — the changelog starts here.

[0.52.0]: https://github.com/jandira-tech/docx-validate/compare/v0.51.0...v0.52.0
[0.51.0]: https://github.com/jandira-tech/docx-validate/compare/v0.5.0...v0.51.0
[0.5.0]: https://github.com/jandira-tech/docx-validate/compare/v0.1.3...v0.5.0
[0.1.3]: https://github.com/jandira-tech/docx-validate/compare/v0.1.2...v0.1.3
[0.1.2]: https://github.com/jandira-tech/docx-validate/compare/v0.1.1...v0.1.2
[0.1.1]: https://github.com/jandira-tech/docx-validate/releases/tag/v0.1.1
