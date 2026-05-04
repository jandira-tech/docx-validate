# Changelog

All notable changes to `docx-validate` are documented here. Format follows
[Keep a Changelog](https://keepachangelog.com/en/1.1.0/) and the project
adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

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

[0.1.3]: https://github.com/jandira-tech/docx-validate/compare/v0.1.2...v0.1.3
[0.1.2]: https://github.com/jandira-tech/docx-validate/compare/v0.1.1...v0.1.2
[0.1.1]: https://github.com/jandira-tech/docx-validate/releases/tag/v0.1.1
