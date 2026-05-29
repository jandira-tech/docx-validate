# Four-Class Architecture — Design Spec

**Date:** 2026-05-29
**Branch:** `spec/four-class-architecture` ← `feat/inventory-fingerprint`
**Status:** Draft, brainstormed with user, awaiting implementation-plan handoff

> One library. One bundle. Four top-level classes (`Validate`, `Repair`, `Normalize`, `Measure`).
> Browser-safe core + strictly-segregated Node conveniences, no duplication.
> The XSD engine moves from `libxmljs2` (Node-native) to `libxml2-wasm` (works in both runtimes).
> Eventually exposed under the `jubarte.validate / read / write / repair / normalize / measure` namespace
> when consumed by the jubarte-first AST hub.

---

## §0 — TL;DR

Today docx-validate is a Node-only library because of `libxmljs2`'s native bindings. The four-class
refactor (a) lifts the existing scattered validate/repair logic into four named top-level classes,
(b) unifies the scattered "normalize" operations (mergeRuns, simplifyRedlines, whitespace-preserve
injection) into a single `Normalize` class, (c) adds a per-file `Measure` class adapted from
jubarte-first's audit pipeline, and (d) swaps the XSD engine to a WASM port so the entire library
runs in both Node and the browser from one source.

Node-only conveniences (file-path helpers, CLI, LibreOffice flow, corpus drivers) move to
`src/node/` and are exposed via a separate `./node` package entry. They wrap the core; they never
duplicate it. A test enforces the segregation by ripgrep-ing for `node:*` imports outside `src/node/`.

The refactor lands as three stacked PRs (A: introduce wasm validator interface; B: rework
`BaseValidator` to use it and drop `libxmljs2`; C: add the four top-level classes + `jubarte.*`
namespace + the `src/node/` segregation).

---

## §1 — Architecture

One library, one bundle, four top-level classes plus functional helpers. No `*.node.ts` /
`*.browser.ts` paired files. Platform difference is handled by **one** swapped dependency:
`libxmljs2` → `libxml2-wasm`.

### 1.1 Directory layout

```
src/
  validate.ts          — Validate class               ← platform-agnostic
  repair.ts            — Repair class                 ← platform-agnostic
  normalize.ts         — Normalize class              ← platform-agnostic
  measure.ts           — Measure.runOne (per-file)    ← platform-agnostic
  index.ts             — browser-safe barrel: the 4 classes + jubarte.*
  lib/
    xml-helpers.ts                                    ← already clean
    xsd-validator.ts   — libxml2-wasm wrapper         ← runs everywhere
    types.ts
  scripts/office/      — internal validators          ← audited to be node:*-free
                         (BaseValidator etc. — the 4 top-level classes delegate here)

  node/                ← NODE-ONLY, strictly segregated, zero duplication
    index.ts           — node barrel
    validate-file.ts   — validateFile(path) → reads bytes, calls core Validate
    repair-file.ts     — same pattern
    normalize-file.ts  — same pattern
    measure-corpus.ts  — Measure.runCorpus driver (reads manifest from disk)
    accept-changes.ts  — LibreOffice flow (uses child_process)
    pack.ts / unpack.ts — file-tree variants; thin wrappers around in-memory ZIP core
    cli.ts             — commander wiring
    lib/temp.ts        — withTempDir (uses os.tmpdir)
```

### 1.2 Package exports

```json
{
  "exports": {
    ".":      { "import": "./dist/index.js",      "types": "./dist/index.d.ts" },
    "./node": { "import": "./dist/node/index.js", "types": "./dist/node/index.d.ts" }
  }
}
```

Browser app: `import { Validate, jubarte } from "docx-validate"` — bytes-only, zero `node:*`.
Node app: `import { Validate, validateFile, acceptChanges } from "docx-validate/node"` — everything.

### 1.3 No-duplication contract

`src/node/*` files MUST import each class from `src/` and only add I/O — never re-implement logic.
A test (delivered in PR A) enforces it:

```ts
// tests/no-node-imports-in-core.test.ts
it("no src/ file outside src/node/ imports node:*", () => {
  // ripgrep ^import .* from ['"]node:  across src/ excluding src/node/
  // expected: zero matches
});
```

Any drift fails CI.

---

## §2 — Class scopes

All four classes take `Uint8Array`. File-path overloads live in `src/node/` (wrappers, not core).

```ts
class Validate {
  constructor(opts?: { xsdValidator?: XsdValidator; schemasDir?: string });
  async run(bytes: Uint8Array): Promise<ValidationResult>;
}

class Repair {
  // Exactly the surface that BaseValidator.repair() + DocxValidator.repair() implement today,
  // lifted into a class: whitespace xml:space="preserve" injection, structural fixes.
  // Returns repaired bytes + repair count + diagnostics.
  constructor(opts?: { xsdValidator?: XsdValidator });
  async run(bytes: Uint8Array): Promise<{
    bytes: Uint8Array;
    repairs: number;
    diagnostics: Issue[];
  }>;
}

class Normalize {
  // New unified surface. Idempotent. Covers what's scattered today:
  //   mergeRuns, simplifyRedlines, whitespace-preserve injection.
  // Distinct from Repair: assumes valid input; the goal is a canonical form, not fixup.
  async run(bytes: Uint8Array): Promise<{ bytes: Uint8Array; changed: boolean }>;
}

class Measure {
  // New in docx-validate (adapted from jubarte-first's audit-a-runner).
  // Per-file: read → write → re-read → compare → classify.
  async runOne(bytes: Uint8Array): Promise<MeasureResult>;
}
```

### 2.1 Functional namespace (the eventual `jubarte.*` shape)

```ts
export const jubarte = {
  validate:  (bytes: Uint8Array) => new Validate().run(bytes),
  repair:    (bytes: Uint8Array) => new Repair().run(bytes),
  normalize: (bytes: Uint8Array) => new Normalize().run(bytes),
  measure:   (bytes: Uint8Array) => new Measure().runOne(bytes),
  // read / write come from the AST hub (jubarte-first project) and are composed
  // into this namespace when the AST hub consumes docx-validate as a dependency.
};
```

Power users get the classes (for `xsdValidator` injection, custom schemas dirs). The 90% case is
the functional helpers.

### 2.2 Node-only conveniences

```ts
// src/node/validate-file.ts
export async function validateFile(path: string): Promise<ValidationResult> {
  const bytes = await fs.readFile(path);
  return new Validate().run(bytes);
}

// src/node/measure-corpus.ts
export async function runCorpus(manifest: ManifestEntry[]): Promise<CorpusReport> {
  // reads each fixture path, calls Measure.runOne, aggregates classification + metrics
}

// src/node/accept-changes.ts — unchanged behavior; just relocated
// src/node/pack.ts / unpack.ts — file-tree variants
// src/node/cli.ts — commander wiring; delegates to the 4 classes via the file helpers
```

---

## §3 — The libxml2-wasm swap

This is the load-bearing change that enables "one binary." Today `BaseValidator._validateSingleFileXsd`
calls `libxmljs2` directly. Tomorrow it calls a `XsdValidator` interface; the wasm impl is the
default factory.

### 3.1 Interface

```ts
// src/lib/xsd-validator.ts — the ONLY place that touches the XSD engine
export interface XsdValidator {
  validate(xml: string, schemaPath: string): Promise<Issue[]>;
}

export async function createXsdValidator(): Promise<XsdValidator> {
  const { parseXmlString, XmlDocument } = await import("libxml2-wasm");
  // wraps the wasm engine in the interface above
}
```

`createXsdValidator()` is memoized; one WASM init per process.

### 3.2 Integration into `BaseValidator`

```ts
class BaseValidator {
  constructor(opts: { xsdValidator?: XsdValidator; ... }) {
    this.xsdValidator = opts.xsdValidator ?? defaultXsdValidator;
  }
  async _validateSingleFileXsd(...) {
    return this.xsdValidator.validate(xml, schemaPath);
  }
}
```

The direct `libxmljs2` call goes away. `BaseValidator`'s public behavior is unchanged.

### 3.3 Behavior parity (must hold under wasm)

These are CLAUDE.md-documented behaviors that must survive the swap; PR A delivers parity tests:

- Note 2: `@xmldom/xmldom` text-node mutation via `.data` (not `.nodeValue`). This is independent of
  libxmljs2/libxml2-wasm — `@xmldom/xmldom` is untouched.
- Note 4: `IGNORED_VALIDATION_ERRORS` filtering ("Invalid XSD schema" + "purl.org/dc/terms") still
  applies. Strings come from the underlying libxml2 C library — same strings under WASM.
- ISO Strict skip: `STRICT_OOXML_NAMESPACES` detection is in TypeScript, not the XSD engine.
  Untouched.
- Note 8: `[ \t\n\r]` regex in `validateWhitespacePreservation` — independent.

PR A also delivers a parity smoke test comparing `libxmljs2` vs `libxml2-wasm` output across the
fixture corpus before `libxmljs2` is dropped in PR B.

---

## §4 — What changes vs. what stays

### Changes

- **Deps:** drop `libxmljs2`, add `libxml2-wasm`.
- **New files:** `src/validate.ts`, `repair.ts`, `normalize.ts`, `measure.ts`, `lib/xsd-validator.ts`.
- **New tree:** `src/node/` with the relocated Node-only files.
- **Reworked:** `BaseValidator` takes injected `XsdValidator`; no direct `libxmljs2` import there.
- **Barrel:** `src/index.ts` exports the 4 classes + `jubarte` namespace; current exports stay (but
  any `node:*`-touching internals get moved under `src/node/index.ts`).
- **Package.json exports:** dual entry (`.` and `./node`).
- **CLI shim:** `src/node/cli.ts` (moved from `src/scripts/office/validate.ts`) delegates to
  `new Validate(...).run(...)`.

### Stays untouched

- All existing validator behavior: XSD scope, ignored errors, Strict-namespace skip, Note 8 regex.
- `ValidationResult` shape (Note 1).
- `accept-changes.ts` — same logic, relocated to `src/node/accept-changes.ts`.
- `pack`/`unpack` byte-parity best-effort behavior (Note 3) — the in-memory ZIP core preserves it.
- All test fixtures.
- `compareParagraphCounts` shape (Note 5).
- Note 12 namespace constants on `validators/docx.ts`.
- Note 13's descendant `xmlns:*` walk.
- Note 14's `inferAuthorFunc` callback on `pack()`.

---

## §5 — Migration & TDD plan (high-level)

Three stacked PRs, each over the previous, each independently green. Full step-by-step plan comes
via the writing-plans skill in a follow-up.

### PR A — `feat/xsd-validator-interface` ← `spec/four-class-architecture`

- Add `src/lib/xsd-validator.ts` (interface + libxml2-wasm impl).
- Add `libxml2-wasm` to deps; keep `libxmljs2` for now.
- Add `tests/no-node-imports-in-core.test.ts` (the segregation enforcement test).
- Add `tests/xsd-validator-parity.test.ts` (wasm vs. libxmljs2 across the existing fixture corpus).
- No consumer changes yet. `BaseValidator` still calls `libxmljs2` directly.

### PR B — `feat/wasm-xsd-cutover` ← PR A

- Rework `BaseValidator` to take injected `XsdValidator`; default factory is wasm.
- Drop `libxmljs2` from deps.
- All existing tests must pass.
- Audit `src/scripts/office/` for `node:*` imports; relocate anything that touches them to
  `src/node/` with thin wrappers staying in place if backwards-compat is needed.

### PR C — `feat/four-class-surface` ← PR B

- Add `src/validate.ts`, `repair.ts`, `normalize.ts`, `measure.ts`.
- Add `src/index.ts` barrel with the 4 classes + `jubarte` namespace.
- Add `src/node/` tree and `package.json` `exports` dual entry.
- Move CLI to `src/node/cli.ts`.
- New tests: orchestrator-level tests for each class; behavior covered by existing validator tests.

Each PR follows TDD red-green; each PR is over the precedent per repo policy
(CLAUDE.md: "Each PR must be over the precedent PR. Test ALWAYS. No exceptions.").

---

## §6 — Open risks

1. **libxml2-wasm XSD parity.** Both engines wrap the same C library, so outputs should match, but
   the WASM build occasionally lags on optional features (XInclude, schema imports across
   namespaces). PR A's parity smoke test on the fixture corpus catches divergence before PR B's
   cutover. If divergence is found, PR B is paused while the wasm gap is fixed or worked around.

2. **`pack`/`unpack` Node-FS coupling.** Today these traverse the filesystem for the unpacked tree.
   The in-memory ZIP path exists (JSZip) but the file-tree variant is what tests use. Splitting
   into pure-bytes core (in `src/`) + file-tree wrapper (in `src/node/`) is real work. Sized in
   the PR C implementation plan, not deferred.

3. **`Measure` runtime cost.** jubarte-first's audit-a-runner takes minutes over 609 fixtures.
   `Measure.runOne` is per-file and cheap; `runCorpus` is opt-in and Node-only. CI uses sampled
   subsets.

4. **WASM init in the browser bundle.** `libxml2-wasm` ships a `.wasm` payload (~2-5 MB). vite-plus
   must be configured to inline-or-fetch it correctly. Verified in PR C's browser smoke test.

5. **Tree-shaking the unused class.** A browser app importing only `Validate` should not pay for
   `Measure`'s round-trip code. The four-class file layout supports this; verified by a
   bundle-size assertion in PR C.

---

## §7 — Future: `jubarte.read` / `jubarte.write`

The `read` / `write` AST verbs come from the jubarte-first project, not docx-validate. When
jubarte-first consumes docx-validate as a dependency, it composes:

```ts
// jubarte-first/src/index.ts (sketch)
import { jubarte as validators } from "docx-validate";
import { docxToAst, astToDocx } from "./ast-hub";

export const jubarte = {
  ...validators,           // validate, repair, normalize, measure
  read:  docxToAst,
  write: astToDocx,
};
```

The naming convention (`jubarte.*` as the public surface) is reserved here so the merge is
mechanical when it lands. No coordination work is required in docx-validate beyond keeping the
namespace name stable.
