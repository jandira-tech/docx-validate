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

OOXML validators and redline/comment helpers for `.docx` and `.pptx` files. XSD-backed, TypeScript, ESM, runs under Node and Bun for the neurotic developer.

- Repo: [jandira-tech/docx-validate](https://github.com/jandira-tech/docx-validate)
- npm: `docx-validate`
- Maintainer: [jandira.tech](https://www.jandira.tech) — We are building legal tech. Jandira Technologies is the studio behind tools like [Cicero](https://www.cicero.im) (a legal workbench that turns messy inputs into redlines, issue lists, and memos), PII redaction models for Brazilian Portuguese, and AI/contract-drafting benchmarks. `docx-validate` falls out of that work — when redlines have to round-trip through Word, validating the OOXML directly beats trusting the renderer.

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

## Programmatic use

Everything the CLIs do is also available as plain function/class imports — no
shell required. The package's barrel (`src/index.ts`) is auto-generated from
the source tree, so anything exported by a `src/**` file is reachable from
the package root.

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

Strict profile (flags BOM-prefixed parts and other tolerated-but-non-canonical
constructs):

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
(`bunx tsx src/scripts/accept-changes.ts <input> <output>`) rather than
importing from the published package.

### Result shape — `ValidationResult`

```ts
interface ValidationIssue {
    severity: "error" | "warning" | "info";
    message: string;
    path?: string; // file path inside the unpacked dir, when applicable
    code?: string; // stable string ID; safe to switch on
}

interface ValidationResult {
    valid: boolean; // true when every issue is severity !== "error"
    issues: ValidationIssue[];
}
```

`ValidateRunResult` extends this with `suffix` (e.g. `".docx"`) and `repairs`
(number of issues auto-repaired when `autoRepair: true`).

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
`src/scripts/office/schemas/` remain the source of truth for _element
definitions_; `NS` is the source of truth for _namespace URIs the TS
code references at runtime_.

## License

Apache-2.0 — see [LICENSE](./LICENSE).
