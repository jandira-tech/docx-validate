# docx-validate

OOXML validators and redline/comment helpers for `.docx` and `.pptx` files. XSD-backed, TypeScript, ESM, runs under Node and Bun.

- Repo: [jandira-tech/docx-validate](https://github.com/jandira-tech/docx-validate)
- npm: `docx-validate`

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
        base.ts                  — XSD schema loader (libxmljs2)
        docx.ts                  — DOCX validator
        pptx.ts                  — PPTX validator
        redlining.ts             — w:ins / w:del / w:moveTo etc. validator
tests/                           — vitest specs
```

## Conventions

- File names: kebab-case (`merge-runs.ts`, not `merge_runs.ts` or `mergeRuns.ts`).
- Modules: ES modules (`import`/`export`), no CommonJS.
- Imports: relative paths between sibling modules; no path aliases.
- XSD validation goes through `libxmljs2` (it ships its own types).
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
- `libxmljs2` — XSD validation
- `commander` — CLI argument parsing
- `tmp` — temp directory lifecycle

Dev:

- `@types/tmp`, `@types/node`, `typescript`, `vite-plus`, `bumpp`

## Running CLIs

```bash
bunx tsx src/scripts/office/validate.ts <path>
bunx tsx src/scripts/office/unpack.ts <path>
bunx tsx src/scripts/office/pack.ts <dir>
bunx tsx src/scripts/accept-changes.ts <path>
bunx tsx src/scripts/comment.ts <path>
```

## Development

```bash
bun install
bun run test       # vitest
bun run check      # type-check
bun run build      # vite-plus build
```

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

The bundled XSDs declare ~30 `targetNamespace` values; `NS` currently
exposes 13. The other 17 are unused at call sites, so adding them would
just be dead surface area. XSD validation is independent: `libxmljs2`
reads the `.xsd` files directly, so `NS` membership has no effect on
what gets validated.

When a future validator needs `c:chart` (DrawingML chart) or `v:shape`
(VML), add the constant to `NS` next to the validator that uses it —
don't pre-populate speculatively. The schemas under
`src/scripts/office/schemas/` remain the source of truth for *element
definitions*; `NS` is the source of truth for *namespace URIs the TS
code references at runtime*.

## License

Apache-2.0 — see [LICENSE](./LICENSE).
