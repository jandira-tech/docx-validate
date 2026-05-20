# Project: docx-validate

OOXML validators and redline/comment helpers for `.docx` and `.pptx` files.

- Repo: `jandira-tech/docx-validate`
- npm: `docx-validate` (unscoped)
- Language: TypeScript, ESM only
- Runtime: Node + Bun
- Build: `vite-plus`
- Tests: `vitest`
- Package manager: `bun`

## Code conventions

- File names: kebab-case (`merge-runs.ts`).
- ES modules only; no CommonJS.
- Relative imports between siblings; no path aliases.
- XSD validation: `libxmljs2`.
- DOM work: `@xmldom/xmldom` + `xpath`, always via `src/lib/xml-helpers.ts`. Do NOT call `DOMParser` directly.
- Temp dirs: `tmp` via `withTempDir(async (dir) => { ... })` (`src/lib/run-cli.ts`). Don't ad-hoc `os.tmpdir()` + `mkdtempSync`.
- CLIs: `commander`, wired via `runCli(...)`.
- Validation results: return the `ValidationResult` shape from `src/lib/types.ts`. No per-validator shapes.
- Tests: one vitest file per module, fixtures under `tests/fixtures/`.

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
tests/                      — vitest specs + fixtures
```

## Behavioural notes

These are intentional behaviours of the implementation, documented at the call sites and verified by the vitest specs.

1. **Validator results are structured, not printed.** `validate_*` methods return `ValidationResult { valid, issues[] }`. The CLI shim in `validate.ts` is the only place that emits human-readable output. Tests assert on issue codes/messages.

2. **`@xmldom/xmldom` text-node mutation requires `.data`, not `.nodeValue`.** Validators that rewrite text nodes (whitespace repair, template-tag stripping) must use `(textNode as Text).data = …`. Writes to `.nodeValue` are silently dropped at serialization. See `validators/base.ts` near `_preprocessXmlForXsd`.

3. **`unpack`/`pack` byte parity is best-effort.** JSZip emits ZIP entries deflated with its own compression settings, so byte-for-byte round-trips against external tools are not guaranteed. XML pretty-printing/indentation is preserved so contents round-trip identically; the outer ZIP envelope can differ. This affects file-checksum equality, not logical-equivalence checks.

4. **`IGNORED_VALIDATION_ERRORS` includes `"Invalid XSD schema"` and `"purl.org/dc/terms"`.** libxmljs2 swallows underlying schema-load failures (e.g. unresolved `<xs:import namespace=".../dc/terms"/>` in `opc-coreProperties.xsd`) into a single opaque error string. Both strings are filtered so the docProps/core.xml false-positive stays suppressed.

5. **`compareParagraphCounts` returns a structured result.** `{ original, modified, delta, originalUsesStrictNamespace }`. With `verbose: true`, also prints `Paragraphs: N → M (+K)` to stdout before returning.

6. **XSD line numbers are not propagated.** `@xmldom` does not surface `sourceline`. Validation issues report file paths only; tests assert on path + message, not line number.

7. **`accept-changes.ts` wraps `runSoffice()` in `runSofficeWithTimeout(args, 30000)`.** Node's `child_process.spawn` has no built-in timeout that resolves cleanly with stdout/stderr, so the wrapper is a `setTimeout`-driven Promise.

8. **`validateWhitespacePreservation` uses `[ \t\n\r]`, not `\s`.** JavaScript's `\s` is broader (matches `\f`, `\v`, Unicode whitespace), so the explicit character class is used to keep the match set predictable.

9. **`RedliningValidator` class wraps free functions.** `RedliningValidator.validate()` and `.repair()` delegate to the free-function implementation. `repair()` returns `0`.

10. **`validate.ts` for unsupported suffixes returns a structured error.** For `.xlsx` and similar, `runValidators` writes to `process.stderr` and returns a `ValidationResult` with `code: "unsupported-file-type"` — it does NOT throw.

11. **`repairWhitespacePreservation` prints a per-repair line when `verbose: true`.** `Repaired: <file>: Added xml:space='preserve' to <tag>: <preview>` per element fixed.

12. **Namespace constants exported from `validators/docx.ts`:** `WORD_2006_NAMESPACE`, `WORD_STRICT_NAMESPACE`, and `WORD_PARAGRAPH_NAMESPACES` (a `readonly [string, string]` tuple).

13. **`validateNamespaces` walks all descendants for `xmlns:*` collection.** This mirrors lxml's `nsmap` propagation: namespace prefixes declared on child elements (not just the root) are collected.

14. **`pack()` accepts an `inferAuthorFunc` callback.** `inferAuthorFunc?: (unpackedDir: string, originalDocx: string) => Promise<string> | string`. When omitted, the default `inferAuthor` from `simplify-redlines.ts` is used.

## ISO OOXML Strict XSD validation

ISO OOXML defines two conformance classes: **Transitional** (backward-compatible, widely used) and **Strict** (pure ISO, uses different namespace URIs under `purl.oclc.org/ooxml/` instead of `schemas.openxmlformats.org/…`). Example: Strict WML uses `http://purl.oclc.org/ooxml/wordprocessingml/main` vs Transitional's `http://schemas.openxmlformats.org/wordprocessingml/2006/main`.

The bundled XSD schemas under `src/scripts/office/schemas/` are Transitional-only. Applying them to Strict documents produces ~30 false-positive `xsd-*` errors per file (every element is flagged as from an unknown namespace).

**Current behaviour:** `validateAgainstXsd` in `validators/base.ts` detects Strict conformance by checking each XML file's root element `namespaceURI` against `STRICT_OOXML_NAMESPACES`. When detected, XSD validation is skipped for the entire document and a single `info`-severity `xsd-strict-skipped` issue is recorded. This keeps `result.valid === true` for structurally well-formed Strict documents while being honest about the gap.

**Future improvement:** Ship the ISO OOXML Strict XSD schemas alongside the Transitional ones and add a second dispatch path in `_getSchemaPath` / `_validateSingleFileXsd`. Deferred — requires bundling ~50 MB of additional schema files and building a Strict-aware preprocessing pipeline.

The `STRICT_OOXML_NAMESPACES` constant in `validators/base.ts` and the `WORD_STRICT_NAMESPACE` constant in `validators/docx.ts` are the canonical sources for these URIs.

## Known gaps

- **soffice end-to-end tests are gated on `process.env.SOFFICE_AVAILABLE`.** Both `tests/accept-changes.test.ts` and `tests/soffice.test.ts` skip real-LibreOffice invocations unless that env var is set; CI images typically don't include LibreOffice. Run with `SOFFICE_AVAILABLE=1 bun run test` locally to exercise them.

- **Three `eslint-plugin-functional` rules are disabled in `.oxlintrc.json`:** `functional/immutable-data`, `functional/no-mixed-types`, `functional/prefer-readonly-type`. All three are type-aware and call `getParserServices()`, but oxlint's JS-plugin runtime hard-codes `parserServices: ObjectFreeze({})` (see `node_modules/oxlint/dist/lint.js:13580`), so they crash the AST walker on any file containing a type alias / interface or an `arr.push(...)` call. `prefer-property-signatures` is the only functional rule left enabled because it's syntactic-only. Re-enable the three when oxlint either provides parser services to JS plugins or honors `requiresTypeChecking: true` for skip.

## Verification

```bash
bun install
bunx tsc --noEmit       # type-check
bun run test            # vitest
bun run check           # vite-plus check
bun run build           # produce dist/
```

## Tooling preferences

- Frontend tooling preference order: bun, biome, elysia (over npm/pnpm/eslint/prettier).
- Favor Rust-backed tools.
- Async-first when writing supporting Python (uv, httpx, uvloop, pydantic, polars, pydantic_ai).

## Notes

- The vendored `docx/` folder is a third-party library snapshot (the `docx` npm package). Its `AGENTS.md` is unrelated to this project — don't treat it as authoritative.
- `docx-templates/` is similarly a third-party reference snapshot.
