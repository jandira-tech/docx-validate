# Migration guide

Notes for upgrading between breaking releases of `docx-validate`.
The project follows [SemVer](https://semver.org/spec/v2.0.0.html); under `0.x`,
minor versions carry breaking changes.

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
against a schema it could not load. The known-noisy bundled case
(`opc-coreProperties.xsd`'s remote Dublin Core import) stays suppressed via
the message filter in `IGNORED_VALIDATION_ERRORS`, so `docProps/core.xml`
does not regress.

**Action**: tests or callers asserting on `xsd-schema-load-skipped`, or
relying on load failures being non-fatal, must switch to
`xsd-schema-load-failed` (error).

### New validation code: `ct-uncovered-part`

A residual OPC-coverage pass now flags parts matched by **no**
`<Override PartName>` and **no** `<Default Extension>` in
`[Content_Types].xml` (previously only declarable XML roots and known media
extensions were checked). Packages that quietly relied on the gap will now
fail with `ct-uncovered-part`. The comparison decodes URI-escaped
`PartName`s and treats only own-properties as known media extensions.

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
