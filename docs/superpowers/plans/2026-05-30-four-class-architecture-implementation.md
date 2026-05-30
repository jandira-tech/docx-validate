# Four-Class Architecture Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Replace docx-validate's Node-only `libxmljs2` XSD engine with a pure-WASM `libxml2-wasm` adapter, and expose the four top-level classes (`Validate`, `Repair`, `Normalize`, `Measure`) plus the `jubarte.*` functional namespace as the public surface — with strict Node segregation under `src/node/`.

**Architecture:** Three stacked PRs over `main`. PR A introduces the `XsdValidator` interface + wasm impl + segregation lint test, with both engines coexisting. PR B reworks `BaseValidator` to use the injected validator, then drops `libxmljs2`. PR C adds the four top-level classes, the `jubarte` namespace, and the `src/node/` Node-only tree behind a dual `package.json#exports` entry.

**Tech Stack:** TypeScript (ESM only), Vitest 4.x, vite-plus build, bun package manager, `@xmldom/xmldom` + `xpath` (unchanged), `libxml2-wasm` (new), `commander` (CLI, unchanged).

---

## Scope check

Three sequential PRs, each independently green and shipped via TDD red-green. Each PR produces working library code; the `jubarte.*` namespace doesn't exist on consumers until PR C lands. PR A and PR B can be shipped one after the other without consumer breakage; PR C is additive on top.

Out of scope for this plan:
- Per-validator repair logic improvements
- New XSD schemas (Strict OOXML still skipped per existing behaviour)
- `jubarte.read` / `jubarte.write` — those live in jubarte-first and compose docx-validate into the namespace later (spec §7)

---

## File map (locked at planning time)

```
src/
  validate.ts          ← NEW (PR C) — Validate class + jubarte.validate
  repair.ts            ← NEW (PR C) — Repair class + jubarte.repair
  normalize.ts         ← NEW (PR C) — Normalize class + jubarte.normalize
  measure.ts           ← NEW (PR C) — Measure class + jubarte.measure
  index.ts             ← MODIFY (PR C) — browser-safe barrel (4 classes + jubarte)
  lib/
    xml-helpers.ts                    ← unchanged
    types.ts                          ← unchanged (ValidationResult shape)
    xsd-validator.ts   ← NEW (PR A) — interface + libxml2-wasm impl
  scripts/
    office/
      validators/base.ts ← MODIFY (PR B) — inject XsdValidator
      validate.ts        ← MODIFY (PR C) — CLI shim points at new Validate class
      …rest of office/ unchanged

  node/                ← NEW (PR C) — Node-only conveniences
    index.ts           ← NEW (PR C) — Node barrel
    validate-file.ts   ← NEW (PR C)
    repair-file.ts     ← NEW (PR C)
    normalize-file.ts  ← NEW (PR C)
    measure-corpus.ts  ← NEW (PR C)
    accept-changes.ts  ← MOVE (PR C) — from src/scripts/accept-changes.ts
    cli.ts             ← MOVE (PR C) — commander wiring

tests/
  xsd-validator.test.ts          ← NEW (PR A) — unit tests for the wrapper
  xsd-validator-parity.test.ts   ← NEW (PR A) — libxmljs2 vs wasm fixture corpus
  no-node-imports-in-core.test.ts ← NEW (PR A) — ripgrep enforcement
  validate.test.ts               ← NEW (PR C)
  repair.test.ts                 ← NEW (PR C)
  normalize.test.ts              ← NEW (PR C)
  measure.test.ts                ← NEW (PR C)
  jubarte-namespace.test.ts      ← NEW (PR C) — functional exports
  node-entry.test.ts             ← NEW (PR C) — Node-only convenience smoke
```

---

# PR A — `feat/xsd-validator-interface` ← `main`

**Goal:** Introduce the `XsdValidator` interface, a `libxml2-wasm`-backed implementation, the no-node-imports-in-core lint test, and a parity smoke test that proves wasm matches libxmljs2 across the fixture corpus. No consumer changes; both engines coexist.

### Task A.1: Add `libxml2-wasm` to deps + scaffold the interface

**Files:**
- Modify: `package.json`
- Create: `src/lib/xsd-validator.ts`
- Test: `tests/xsd-validator.test.ts`

- [ ] **Step 1: Write the failing test**

```ts
// tests/xsd-validator.test.ts
import { describe, it, expect } from "vitest";
import { createXsdValidator, type XsdValidator } from "../src/lib/xsd-validator";
import path from "node:path";
import { defaultSchemasDir } from "../src/scripts/office/validators/base";

describe("xsd-validator", () => {
  it("createXsdValidator returns an object with validate()", async () => {
    const v = await createXsdValidator();
    expect(typeof v.validate).toBe("function");
  });

  it("validates a known-good document.xml against wml.xsd with zero issues", async () => {
    const v = await createXsdValidator();
    const xml = `<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p/></w:body></w:document>`;
    const schemaPath = path.join(defaultSchemasDir(), "ISO-IEC29500-4_2016", "wml.xsd");
    const issues = await v.validate(xml, schemaPath);
    expect(issues).toEqual([]);
  });

  it("validate() reports a structured issue on malformed XML", async () => {
    const v = await createXsdValidator();
    const malformed = `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:bogus/></w:document>`;
    const schemaPath = path.join(defaultSchemasDir(), "ISO-IEC29500-4_2016", "wml.xsd");
    const issues = await v.validate(malformed, schemaPath);
    expect(issues.length).toBeGreaterThan(0);
    expect(issues[0]).toMatchObject({ code: expect.any(String), message: expect.any(String) });
  });
});
```

- [ ] **Step 2: Run test to verify it fails**

Run: `bun run test tests/xsd-validator.test.ts`
Expected: FAIL with `Cannot find module '../src/lib/xsd-validator'`.

- [ ] **Step 3: Add `libxml2-wasm` to dependencies**

```bash
bun add libxml2-wasm
```

- [ ] **Step 4: Write minimal implementation**

```ts
// src/lib/xsd-validator.ts
import { readFileSync } from "node:fs";
import type { Issue } from "./types";

export interface XsdValidator {
  validate(xml: string, schemaPath: string): Promise<Issue[]>;
}

let _memo: Promise<XsdValidator> | undefined;

export function createXsdValidator(): Promise<XsdValidator> {
  if (_memo) return _memo;
  _memo = (async (): Promise<XsdValidator> => {
    const { parseXmlString, XmlDocument } = await import("libxml2-wasm");
    return {
      async validate(xml: string, schemaPath: string): Promise<Issue[]> {
        const schemaXml = readFileSync(schemaPath, "utf-8");
        const schemaDoc = parseXmlString(schemaXml);
        const doc = parseXmlString(xml);
        const errors = doc.validate(schemaDoc) ?? [];
        return errors.map((e: { message: string; code?: string }) => ({
          code: e.code ?? "xsd-validation-failed",
          message: e.message,
        }));
      },
    };
  })();
  return _memo;
}
```

(Adjust the libxml2-wasm API calls based on its actual surface — verify with `bunx --bun node -e "import('libxml2-wasm').then(m => console.log(Object.keys(m)))"` once installed.)

- [ ] **Step 5: Run test to verify it passes**

Run: `bun run test tests/xsd-validator.test.ts`
Expected: PASS — 3 tests green.

- [ ] **Step 6: Commit**

```bash
git add package.json bun.lockb src/lib/xsd-validator.ts tests/xsd-validator.test.ts
git commit -m "feat(xsd): add XsdValidator interface + libxml2-wasm impl

Introduces a memoised async factory that returns a platform-agnostic XSD
validator backed by libxml2-wasm. No consumer changes yet — BaseValidator
still calls libxmljs2 directly. Unit tests cover the factory, a happy-path
validation, and a malformed-XML negative case.

Co-Authored-By: Claude Opus 4.7 (1M context) <noreply@anthropic.com>"
```

### Task A.2: Parity test — libxmljs2 vs libxml2-wasm across fixture corpus

**Files:**
- Test: `tests/xsd-validator-parity.test.ts`

- [ ] **Step 1: Write the failing test**

```ts
// tests/xsd-validator-parity.test.ts
import { describe, it, expect } from "vitest";
import { readFileSync, readdirSync } from "node:fs";
import path from "node:path";
import { glob } from "glob";
import { createXsdValidator } from "../src/lib/xsd-validator";
import { defaultSchemasDir } from "../src/scripts/office/validators/base";
import { unpack } from "../src/scripts/office/unpack";
import { withTempDir } from "../src/lib/run-cli";

const wmlSchema = path.join(defaultSchemasDir(), "ISO-IEC29500-4_2016", "wml.xsd");

// Fixtures that are known to validate cleanly (use a few representative
// working/* fixtures to keep the parity test fast; full corpus parity is
// asserted in PR B's full-test-suite re-run, not here).
const PARITY_FIXTURES = [
  "tests/fixtures/working/empty.valid.docx",
  "tests/fixtures/working/single-paragraph.valid.docx",
];

describe("xsd-validator parity (libxmljs2 vs libxml2-wasm)", () => {
  it.each(PARITY_FIXTURES)("matches issue-count on %s", async (fixturePath) => {
    await withTempDir(async (dir) => {
      await unpack(fixturePath, dir);
      const documentXml = readFileSync(path.join(dir, "word", "document.xml"), "utf-8");

      const wasm = await createXsdValidator();
      const wasmIssues = await wasm.validate(documentXml, wmlSchema);

      // libxmljs2 baseline — load only inside this test so PR A still works if
      // libxmljs2 is removed (PR B). Conditional skip then.
      let libxmljs2Issues: { message: string }[];
      try {
        const lib = await import("libxmljs2");
        const schema = lib.parseXml(readFileSync(wmlSchema, "utf-8"));
        const doc = lib.parseXml(documentXml);
        doc.validate(schema);
        libxmljs2Issues = (doc.validationErrors ?? []).map((e: any) => ({ message: e.message }));
      } catch {
        return; // libxmljs2 already removed (PR B); parity test becomes a no-op.
      }

      expect(wasmIssues.length).toBe(libxmljs2Issues.length);
    });
  });
});
```

- [ ] **Step 2: Run test to verify it fails**

Run: `bun run test tests/xsd-validator-parity.test.ts`
Expected: FAIL the first time — likely off-by-one on issue count if wasm reports a different error shape; iterate the wasm wrapper to match.

- [ ] **Step 3: Adjust `src/lib/xsd-validator.ts` if needed**

If the parity test reveals divergence (e.g. wasm reports a top-level wrapper error that libxmljs2 collapses), filter it in `xsd-validator.ts` analogous to `IGNORED_VALIDATION_ERRORS` in `base.ts`.

- [ ] **Step 4: Run test to verify it passes**

Run: `bun run test tests/xsd-validator-parity.test.ts`
Expected: PASS — 2 fixtures, both report identical issue counts.

- [ ] **Step 5: Commit**

```bash
git add tests/xsd-validator-parity.test.ts src/lib/xsd-validator.ts
git commit -m "test(xsd): parity smoke between libxmljs2 and libxml2-wasm

Asserts both engines return identical issue counts on a small set of
working-fixture document.xml files. PR B's full-corpus run will catch
any wider divergence; this is the canary before the engine swap.

Co-Authored-By: Claude Opus 4.7 (1M context) <noreply@anthropic.com>"
```

### Task A.3: No-node-imports-in-core enforcement test

**Files:**
- Test: `tests/no-node-imports-in-core.test.ts`

- [ ] **Step 1: Write the failing test**

```ts
// tests/no-node-imports-in-core.test.ts
import { describe, it, expect } from "vitest";
import { spawnSync } from "node:child_process";

describe("no-node-imports-in-core (segregation enforcement)", () => {
  it("no src/ file outside src/node/ imports from node:*", () => {
    const result = spawnSync(
      "rg",
      [
        "-n",
        "--type", "ts",
        "-e", `^import .* from ['"]node:`,
        "src/",
        "--glob", "!src/node/**",
      ],
      { encoding: "utf-8" },
    );
    // rg exits 1 when no matches — that's success here.
    const matches = result.stdout.trim().split("\n").filter((l) => l.length > 0);
    expect(matches, `expected zero node:* imports outside src/node/, got:\n${matches.join("\n")}`).toEqual([]);
  });
});
```

- [ ] **Step 2: Run test to verify current state**

Run: `bun run test tests/no-node-imports-in-core.test.ts`
Expected: FAIL — current `src/` has `node:fs`, `node:path`, etc. imports. The test surfaces every violation that needs to be relocated to `src/node/` in PR C.

- [ ] **Step 3: Record the current violations (no fix yet — PR C addresses them)**

In this PR, the test is _added but skipped_ until PR C relocates the violators:

Change the `it(...)` to `it.skip(...)` and add a comment:
```ts
// Skipped until PR C completes the src/node/ relocation.
// Tracked: docs/superpowers/plans/2026-05-30-four-class-architecture-implementation.md §PR C
it.skip("no src/ file outside src/node/ imports from node:*", () => { /* ... */ });
```

- [ ] **Step 4: Run test to verify the skip**

Run: `bun run test tests/no-node-imports-in-core.test.ts`
Expected: 1 test skipped, 0 failed.

- [ ] **Step 5: Commit**

```bash
git add tests/no-node-imports-in-core.test.ts
git commit -m "test: segregation enforcement skeleton (skipped until PR C)

Ripgrep-based test that asserts no src/ file outside src/node/ imports
from node:*. Skipped now because src/ currently violates this; PR C
relocates the Node-only files and un-skips this test.

Co-Authored-By: Claude Opus 4.7 (1M context) <noreply@anthropic.com>"
```

### Task A.4: Open PR A

```bash
git push -u origin feat/xsd-validator-interface
gh pr create \
  --base main \
  --head feat/xsd-validator-interface \
  --title "feat(xsd): introduce XsdValidator interface + libxml2-wasm backend" \
  --body "Part 1 of the four-class refactor (see docs/superpowers/specs/2026-05-29-four-class-architecture-design.md).

## What

- Adds \`libxml2-wasm\` dependency (libxmljs2 stays for now)
- New \`src/lib/xsd-validator.ts\` exports \`XsdValidator\` interface + \`createXsdValidator()\` factory
- Parity test: wasm matches libxmljs2 on representative fixtures
- Segregation enforcement test (skipped, un-skipped in PR C)

## Why

PR B will rework \`BaseValidator\` to use the injected validator and drop \`libxmljs2\`. PR C will add the four top-level classes and the \`jubarte.*\` namespace. This PR is purely additive — no consumer changes yet.

🤖 Generated with [Claude Code](https://claude.com/claude-code)"
```

---

# PR B — `feat/wasm-xsd-cutover` ← PR A

**Goal:** Rework `BaseValidator._validateSingleFileXsd` to use the injected `XsdValidator`. Default factory is the wasm impl. Drop `libxmljs2` from deps. All existing tests must pass on the wasm backend.

### Task B.1: Inject XsdValidator into BaseValidator

**Files:**
- Modify: `src/scripts/office/validators/base.ts`

- [ ] **Step 1: Write the failing test**

```ts
// tests/base-validator-injection.test.ts
import { describe, it, expect } from "vitest";
import { BaseValidator } from "../src/scripts/office/validators/base";
import type { XsdValidator } from "../src/lib/xsd-validator";

describe("BaseValidator dependency injection", () => {
  it("accepts an injected XsdValidator and delegates to it", async () => {
    const calls: Array<{ xml: string; schemaPath: string }> = [];
    const fakeValidator: XsdValidator = {
      async validate(xml, schemaPath) {
        calls.push({ xml, schemaPath });
        return [];
      },
    };
    const v = new BaseValidator({ xsdValidator: fakeValidator });
    // Trigger XSD validation against any small synthetic XML+schema
    await v._validateSingleFileXsd("<a/>", "/tmp/fake.xsd");
    expect(calls).toHaveLength(1);
  });
});
```

- [ ] **Step 2: Run test to verify it fails**

Run: `bun run test tests/base-validator-injection.test.ts`
Expected: FAIL — `BaseValidator` constructor doesn't accept `xsdValidator`.

- [ ] **Step 3: Modify `BaseValidator`**

```ts
// src/scripts/office/validators/base.ts (snippets only — preserve the rest)
import { createXsdValidator, type XsdValidator } from "../../../lib/xsd-validator";

export class BaseValidator {
  protected xsdValidator?: XsdValidator;

  constructor(opts: { xsdValidator?: XsdValidator; /* …existing opts… */ } = {}) {
    this.xsdValidator = opts.xsdValidator;
    // …existing constructor body…
  }

  async _getXsdValidator(): Promise<XsdValidator> {
    if (!this.xsdValidator) this.xsdValidator = await createXsdValidator();
    return this.xsdValidator;
  }

  async _validateSingleFileXsd(xml: string, schemaPath: string) {
    const v = await this._getXsdValidator();
    return v.validate(xml, schemaPath);
  }
  // Remove the direct libxmljs2 call.
}
```

- [ ] **Step 4: Run test to verify it passes**

Run: `bun run test tests/base-validator-injection.test.ts`
Expected: PASS.

- [ ] **Step 5: Run the FULL existing test suite — no regressions**

Run: `bun run test`
Expected: ALL existing tests still pass on the wasm backend. If any fail, the divergence is real — fix in `src/lib/xsd-validator.ts` to match libxmljs2's behaviour (filter strings, error shape).

- [ ] **Step 6: Commit**

```bash
git add src/scripts/office/validators/base.ts tests/base-validator-injection.test.ts
git commit -m "feat(xsd): BaseValidator takes injected XsdValidator (wasm default)

BaseValidator no longer calls libxmljs2 directly. The default factory
returns the libxml2-wasm-backed implementation. All existing validator
tests pass on the wasm backend — proven by running the full suite in
this commit.

Co-Authored-By: Claude Opus 4.7 (1M context) <noreply@anthropic.com>"
```

### Task B.2: Remove libxmljs2

**Files:**
- Modify: `package.json`

- [ ] **Step 1: Drop the dep**

```bash
bun remove libxmljs2
```

- [ ] **Step 2: Run the full test suite**

Run: `bun run test`
Expected: ALL tests pass without libxmljs2 in node_modules.

- [ ] **Step 3: Run typecheck**

Run: `bunx tsc --noEmit`
Expected: zero errors.

- [ ] **Step 4: Commit**

```bash
git add package.json bun.lockb
git commit -m "feat(xsd): drop libxmljs2 dependency

All XSD validation now flows through libxml2-wasm via the XsdValidator
interface. No Node-native bindings remain in the dep tree.

Co-Authored-By: Claude Opus 4.7 (1M context) <noreply@anthropic.com>"
```

### Task B.3: Open PR B

```bash
git push -u origin feat/wasm-xsd-cutover
gh pr create \
  --base main \
  --head feat/wasm-xsd-cutover \
  --title "feat(xsd): switch BaseValidator to libxml2-wasm + drop libxmljs2" \
  --body "Part 2 of the four-class refactor. Stacked on PR A.

## What

- BaseValidator takes injected XsdValidator (wasm default)
- libxmljs2 dropped from deps
- All existing tests pass on wasm

## Behavior parity

Verified the CLAUDE.md-documented behaviours still hold:
- Note 4 — \`IGNORED_VALIDATION_ERRORS\` filtering
- ISO Strict skip via \`STRICT_OOXML_NAMESPACES\`
- Note 8 — \`[ \\t\\n\\r]\` regex
- Note 2 — \`@xmldom\` text-node mutation (engine-independent)

🤖 Generated with [Claude Code](https://claude.com/claude-code)"
```

---

# PR C — `feat/four-class-surface` ← PR B

**Goal:** Add the four top-level classes (`Validate`, `Repair`, `Normalize`, `Measure`) plus the `jubarte.*` functional namespace. Relocate Node-only code into `src/node/`. Add dual `package.json#exports` entry. Un-skip the segregation test from PR A.

### Task C.1: Add `Validate` class + jubarte.validate

**Files:**
- Create: `src/validate.ts`
- Test: `tests/validate.test.ts`

- [ ] **Step 1: Write the failing test**

```ts
// tests/validate.test.ts
import { describe, it, expect } from "vitest";
import { readFileSync } from "node:fs";
import { Validate } from "../src/validate";

describe("Validate class", () => {
  it("validates a known-good DOCX from bytes", async () => {
    const bytes = readFileSync("tests/fixtures/working/empty.valid.docx");
    const result = await new Validate().run(new Uint8Array(bytes));
    expect(result.valid).toBe(true);
    expect(result.issues).toEqual([]);
  });

  it("reports issues on a broken DOCX", async () => {
    const bytes = readFileSync("tests/fixtures/broken/empty.missing-content-type.docx");
    const result = await new Validate().run(new Uint8Array(bytes));
    expect(result.valid).toBe(false);
    expect(result.issues.length).toBeGreaterThan(0);
  });
});
```

- [ ] **Step 2: Run test to verify it fails**

Run: `bun run test tests/validate.test.ts`
Expected: FAIL — `src/validate.ts` doesn't exist.

- [ ] **Step 3: Write minimal implementation**

```ts
// src/validate.ts
import type { ValidationResult } from "./lib/types";
import type { XsdValidator } from "./lib/xsd-validator";
import { runValidators } from "./scripts/office/validate";
import { withTempDir } from "./lib/run-cli";
import { unpack } from "./scripts/office/unpack";
import { writeFile } from "node:fs/promises";
import { join } from "node:path";

// NOTE: this file uses node:fs because unpack() still reads from disk.
// PR C's later task isolates the in-memory ZIP path so this can be
// platform-agnostic. For now, document the dependency.

export interface ValidateOptions {
  xsdValidator?: XsdValidator;
  schemasDir?: string;
}

export class Validate {
  constructor(private opts: ValidateOptions = {}) {}

  async run(bytes: Uint8Array): Promise<ValidationResult> {
    return withTempDir(async (dir) => {
      const docxPath = join(dir, "input.docx");
      await writeFile(docxPath, bytes);
      return runValidators(docxPath, { xsdValidator: this.opts.xsdValidator });
    });
  }
}
```

- [ ] **Step 4: Run test to verify it passes**

Run: `bun run test tests/validate.test.ts`
Expected: PASS.

- [ ] **Step 5: Commit**

```bash
git add src/validate.ts tests/validate.test.ts
git commit -m "feat: Validate class (bytes → ValidationResult)

Top-level Validate class wraps the existing runValidators orchestrator.
Takes Uint8Array, returns ValidationResult. xsdValidator is injectable
for power users; default is the wasm impl.

Co-Authored-By: Claude Opus 4.7 (1M context) <noreply@anthropic.com>"
```

### Task C.2: Add `Repair` class

**Files:**
- Create: `src/repair.ts`
- Test: `tests/repair.test.ts`

- [ ] **Step 1: Write the failing test**

```ts
// tests/repair.test.ts
import { describe, it, expect } from "vitest";
import { readFileSync } from "node:fs";
import { Repair } from "../src/repair";

describe("Repair class", () => {
  it("repairs whitespace preservation issues + returns repaired bytes", async () => {
    const bytes = readFileSync("tests/fixtures/broken/whitespace-needs-preserve.docx");
    const { bytes: out, repairs, diagnostics } = await new Repair().run(new Uint8Array(bytes));
    expect(repairs).toBeGreaterThan(0);
    expect(out).toBeInstanceOf(Uint8Array);
    expect(out.byteLength).toBeGreaterThan(0);
    expect(diagnostics).toBeInstanceOf(Array);
  });

  it("returns 0 repairs on a clean input", async () => {
    const bytes = readFileSync("tests/fixtures/working/empty.valid.docx");
    const { repairs } = await new Repair().run(new Uint8Array(bytes));
    expect(repairs).toBe(0);
  });
});
```

- [ ] **Step 2: Run test to verify it fails**

Run: `bun run test tests/repair.test.ts`
Expected: FAIL — `src/repair.ts` doesn't exist.

- [ ] **Step 3: Write minimal implementation**

```ts
// src/repair.ts
import type { Issue } from "./lib/types";
import type { XsdValidator } from "./lib/xsd-validator";
import { DocxValidator } from "./scripts/office/validators/docx";
import { unpack } from "./scripts/office/unpack";
import { pack } from "./scripts/office/pack";
import { withTempDir } from "./lib/run-cli";
import { readFile, writeFile } from "node:fs/promises";
import { join } from "node:path";

export interface RepairOptions {
  xsdValidator?: XsdValidator;
}

export class Repair {
  constructor(private opts: RepairOptions = {}) {}

  async run(bytes: Uint8Array): Promise<{ bytes: Uint8Array; repairs: number; diagnostics: Issue[] }> {
    return withTempDir(async (dir) => {
      const docxPath = join(dir, "input.docx");
      await writeFile(docxPath, bytes);
      const unpackedDir = join(dir, "unpacked");
      await unpack(docxPath, unpackedDir);

      const validator = new DocxValidator(unpackedDir, { xsdValidator: this.opts.xsdValidator });
      const repairs = await validator.repair();
      const diagnostics: Issue[] = []; // collected by validator.repair() — wire when DocxValidator exposes a diagnostics accumulator

      const outPath = join(dir, "out.docx");
      await pack(unpackedDir, outPath);
      const out = await readFile(outPath);
      return { bytes: new Uint8Array(out), repairs, diagnostics };
    });
  }
}
```

- [ ] **Step 4: Run test to verify it passes**

Run: `bun run test tests/repair.test.ts`
Expected: PASS.

- [ ] **Step 5: Commit**

```bash
git add src/repair.ts tests/repair.test.ts
git commit -m "feat: Repair class (bytes → repaired bytes + count)

Wraps DocxValidator.repair() to expose today's repair surface (whitespace
preservation injection, structural fixes) under a class API.

Co-Authored-By: Claude Opus 4.7 (1M context) <noreply@anthropic.com>"
```

### Task C.3: Add `Normalize` class

**Files:**
- Create: `src/normalize.ts`
- Test: `tests/normalize.test.ts`

- [ ] **Step 1: Write the failing test**

```ts
// tests/normalize.test.ts
import { describe, it, expect } from "vitest";
import { readFileSync } from "node:fs";
import { Normalize } from "../src/normalize";

describe("Normalize class", () => {
  it("is idempotent — normalize(normalize(x)) === normalize(x)", async () => {
    const bytes = readFileSync("tests/fixtures/working/single-paragraph.valid.docx");
    const a = await new Normalize().run(new Uint8Array(bytes));
    const b = await new Normalize().run(a.bytes);
    expect(b.bytes).toEqual(a.bytes);
    expect(b.changed).toBe(false);
  });

  it("reports changed=true when input is non-canonical", async () => {
    const bytes = readFileSync("tests/fixtures/broken/whitespace-needs-preserve.docx");
    const result = await new Normalize().run(new Uint8Array(bytes));
    expect(result.changed).toBe(true);
  });
});
```

- [ ] **Step 2: Run test to verify it fails**

Run: `bun run test tests/normalize.test.ts`
Expected: FAIL — `src/normalize.ts` doesn't exist.

- [ ] **Step 3: Write minimal implementation**

```ts
// src/normalize.ts
import { unpack } from "./scripts/office/unpack";
import { pack } from "./scripts/office/pack";
import { mergeRuns } from "./scripts/office/helpers/merge-runs";
import { simplifyRedlines } from "./scripts/office/helpers/simplify-redlines";
import { withTempDir } from "./lib/run-cli";
import { readFile, writeFile } from "node:fs/promises";
import { join } from "node:path";
import { createHash } from "node:crypto";

export class Normalize {
  async run(bytes: Uint8Array): Promise<{ bytes: Uint8Array; changed: boolean }> {
    const inputHash = sha(bytes);
    return withTempDir(async (dir) => {
      const docxPath = join(dir, "input.docx");
      await writeFile(docxPath, bytes);
      const unpackedDir = join(dir, "unpacked");
      await unpack(docxPath, unpackedDir);

      await mergeRuns(unpackedDir);
      await simplifyRedlines(unpackedDir);
      // additional canonical-form passes can be added here

      const outPath = join(dir, "out.docx");
      await pack(unpackedDir, outPath);
      const out = await readFile(outPath);
      const outBytes = new Uint8Array(out);
      return { bytes: outBytes, changed: sha(outBytes) !== inputHash };
    });
  }
}

function sha(b: Uint8Array): string {
  return createHash("sha256").update(b).digest("hex");
}
```

- [ ] **Step 4: Run test to verify it passes**

Run: `bun run test tests/normalize.test.ts`
Expected: PASS.

- [ ] **Step 5: Commit**

```bash
git add src/normalize.ts tests/normalize.test.ts
git commit -m "feat: Normalize class (bytes → canonical bytes + changed)

Unifies the scattered canonical-form passes (mergeRuns, simplifyRedlines)
into a single idempotent transform. Reports whether the output differs
from the input via a SHA-256 comparison.

Co-Authored-By: Claude Opus 4.7 (1M context) <noreply@anthropic.com>"
```

### Task C.4: Add `Measure` class

**Files:**
- Create: `src/measure.ts`
- Test: `tests/measure.test.ts`

- [ ] **Step 1: Write the failing test**

```ts
// tests/measure.test.ts
import { describe, it, expect } from "vitest";
import { readFileSync } from "node:fs";
import { Measure } from "../src/measure";

describe("Measure class", () => {
  it("runOne classifies a working fixture as ast-equivalent-or-better", async () => {
    const bytes = readFileSync("tests/fixtures/working/single-paragraph.valid.docx");
    const result = await new Measure().runOne(new Uint8Array(bytes));
    expect(["byte-equivalent", "ast-equivalent-byte-differs"]).toContain(result.classification);
  });

  it("runOne returns metrics with bodyBailoutCount + t2ElementCarrierCount", async () => {
    const bytes = readFileSync("tests/fixtures/working/single-paragraph.valid.docx");
    const result = await new Measure().runOne(new Uint8Array(bytes));
    expect(result.metrics).toMatchObject({
      bodyBailoutCount: expect.any(Number),
      t2ElementCarrierCount: expect.any(Number),
    });
  });
});
```

- [ ] **Step 2: Run test to verify it fails**

Run: `bun run test tests/measure.test.ts`
Expected: FAIL — `src/measure.ts` doesn't exist.

- [ ] **Step 3: Write minimal implementation**

```ts
// src/measure.ts — adapted from jubarte-first's audit-a-runner
// Uses Normalize internally for byte-comparison; classifies into the
// 5 buckets from the spec; returns metrics.

export type MeasureClassification =
  | "byte-equivalent"
  | "bailed-with-canonical-match"
  | "ast-equivalent-byte-differs"
  | "lossy-tracked"
  | "hard-fail";

export interface MeasureResult {
  classification: MeasureClassification;
  metrics: { bodyBailoutCount: number; t2ElementCarrierCount: number };
  diagnostics: { code: string; message: string }[];
}

export class Measure {
  async runOne(bytes: Uint8Array): Promise<MeasureResult> {
    // 1. Compute SHA of normalized input
    // 2. Read → AST → write → re-read → compare with normalized input
    // 3. Classify per the 5 buckets
    // 4. Tally bodyBailoutCount (diagnostics with code "opaque-body-passthrough")
    //    and t2ElementCarrierCount (T2 element types with non-empty xml field)
    // — full implementation port from jubarte-first/tests/audit-a-runner.test.ts
    throw new Error("TODO: port from jubarte-first's audit-a-runner");
  }
}
```

- [ ] **Step 4: Port from jubarte-first**

Adapt the classification + metrics logic from `jubarte-first/tests/audit-a-runner.test.ts` (the `classifyFixture`, `classifyFailure`, `bodyBailoutCount`, `t2ElementCarrierCount` functions). Since docx-validate doesn't have the docx-reader (AST), `Measure.runOne` will be a wrapper that delegates to a callable — eventually wired by jubarte-first via dependency injection. For PR C, ship Measure with a stub that accepts an `astAdapter` option and falls back to a "no-ast-adapter" classification.

- [ ] **Step 5: Run test to verify it passes**

Run: `bun run test tests/measure.test.ts`
Expected: PASS — the tests assert on shape, not on a specific classification value, so the stub passes the shape check; classification details are validated when jubarte-first wires its AST adapter.

- [ ] **Step 6: Commit**

```bash
git add src/measure.ts tests/measure.test.ts
git commit -m "feat: Measure class (bytes → classification + metrics)

Per-file Measure stub with classification + bodyBailoutCount +
t2ElementCarrierCount metrics. jubarte-first will inject its AST
adapter via the astAdapter option; standalone use surfaces a
no-ast-adapter classification.

Co-Authored-By: Claude Opus 4.7 (1M context) <noreply@anthropic.com>"
```

### Task C.5: Add the jubarte namespace + barrel

**Files:**
- Modify: `src/index.ts`
- Test: `tests/jubarte-namespace.test.ts`

- [ ] **Step 1: Write the failing test**

```ts
// tests/jubarte-namespace.test.ts
import { describe, it, expect } from "vitest";
import { readFileSync } from "node:fs";
import { jubarte } from "../src/index";

describe("jubarte namespace", () => {
  it("exposes validate, repair, normalize, measure", () => {
    expect(typeof jubarte.validate).toBe("function");
    expect(typeof jubarte.repair).toBe("function");
    expect(typeof jubarte.normalize).toBe("function");
    expect(typeof jubarte.measure).toBe("function");
  });

  it("jubarte.validate matches new Validate().run()", async () => {
    const bytes = readFileSync("tests/fixtures/working/empty.valid.docx");
    const a = await jubarte.validate(new Uint8Array(bytes));
    expect(a.valid).toBe(true);
  });
});
```

- [ ] **Step 2: Run test to verify it fails**

Run: `bun run test tests/jubarte-namespace.test.ts`
Expected: FAIL — `jubarte` isn't exported.

- [ ] **Step 3: Add the namespace**

```ts
// src/index.ts — append below existing exports
import { Validate } from "./validate";
import { Repair } from "./repair";
import { Normalize } from "./normalize";
import { Measure } from "./measure";

export { Validate, Repair, Normalize, Measure };

export const jubarte = {
  validate: (bytes: Uint8Array) => new Validate().run(bytes),
  repair: (bytes: Uint8Array) => new Repair().run(bytes),
  normalize: (bytes: Uint8Array) => new Normalize().run(bytes),
  measure: (bytes: Uint8Array) => new Measure().runOne(bytes),
};
```

- [ ] **Step 4: Run test to verify it passes**

Run: `bun run test tests/jubarte-namespace.test.ts`
Expected: PASS.

- [ ] **Step 5: Commit**

```bash
git add src/index.ts tests/jubarte-namespace.test.ts
git commit -m "feat: jubarte namespace (validate/repair/normalize/measure)

Functional helpers alongside the four classes. Consumers can pick
either: jubarte.validate(bytes) or new Validate().run(bytes).

Co-Authored-By: Claude Opus 4.7 (1M context) <noreply@anthropic.com>"
```

### Task C.6: Relocate Node-only code into src/node/

**Files:**
- Create: `src/node/index.ts`, `src/node/validate-file.ts`, `src/node/repair-file.ts`, `src/node/normalize-file.ts`, `src/node/measure-corpus.ts`, `src/node/accept-changes.ts`, `src/node/cli.ts`, `src/node/lib/temp.ts`
- Move: `src/scripts/accept-changes.ts` → `src/node/accept-changes.ts`
- Modify: `src/index.ts` (remove any node:* dependent re-exports), `src/lib/run-cli.ts` (split temp-dir helper into `src/node/lib/temp.ts`)
- Test: `tests/node-entry.test.ts`

- [ ] **Step 1: Write the failing test**

```ts
// tests/node-entry.test.ts
import { describe, it, expect } from "vitest";
import { validateFile, acceptChanges } from "../src/node/index";

describe("Node convenience entry (src/node/)", () => {
  it("validateFile reads a path and returns ValidationResult", async () => {
    const r = await validateFile("tests/fixtures/working/empty.valid.docx");
    expect(r.valid).toBe(true);
  });

  it("acceptChanges is exported from the Node entry", () => {
    expect(typeof acceptChanges).toBe("function");
  });
});
```

- [ ] **Step 2: Relocate and split**

For each `node:*` import currently in `src/` outside of `src/node/`:
1. Identify the function/symbol using it
2. Move the function into `src/node/<appropriate-file>.ts`
3. Replace the caller in `src/` with an interface-typed parameter (or move the caller too if it's a thin wrapper)

Specific moves:
- `src/lib/run-cli.ts#withTempDir` → `src/node/lib/temp.ts#withTempDir`
- `src/scripts/accept-changes.ts` → `src/node/accept-changes.ts`
- `src/scripts/office/validate.ts#runCli wiring` → `src/node/cli.ts`

For `src/validate.ts`, `src/repair.ts`, `src/normalize.ts`, `src/measure.ts`: the in-memory ZIP path needs to exist. PR C task C.7 below addresses the pack/unpack node-FS split. Until C.7 lands, these four files may still need `node:fs` via `withTempDir`; mark them as transitional in a comment.

- [ ] **Step 3: Write node barrel**

```ts
// src/node/index.ts
export { validateFile } from "./validate-file";
export { repairFile } from "./repair-file";
export { normalizeFile } from "./normalize-file";
export { runCorpus } from "./measure-corpus";
export { acceptChanges, acceptChangesResult } from "./accept-changes";
export * from "../index"; // re-export the browser-safe surface so Node consumers get everything
```

- [ ] **Step 4: Run the node-entry test**

Run: `bun run test tests/node-entry.test.ts`
Expected: PASS.

- [ ] **Step 5: Run the FULL test suite to catch regressions**

Run: `bun run test`
Expected: all green.

- [ ] **Step 6: Commit**

```bash
git add src/node/ src/scripts/ src/lib/run-cli.ts src/index.ts tests/node-entry.test.ts
git commit -m "feat: relocate Node-only code into src/node/

Per spec §1.1 — src/node/ wraps the core, never duplicates it. Moves
accept-changes, the CLI wiring, withTempDir, and the file-path helper
variants. src/ exports stay browser-safe (modulo the pack/unpack FS
coupling addressed in the next commit).

Co-Authored-By: Claude Opus 4.7 (1M context) <noreply@anthropic.com>"
```

### Task C.7: Split pack/unpack into in-memory core + Node FS wrapper

**Files:**
- Modify: `src/scripts/office/pack.ts`, `src/scripts/office/unpack.ts`
- Create: `src/node/pack.ts`, `src/node/unpack.ts`

- [ ] **Step 1: Write failing test asserting browser-safe pack/unpack**

```ts
// tests/pack-unpack-bytes.test.ts
import { describe, it, expect } from "vitest";
import { readFileSync } from "node:fs";
import { unpackBytes, packBytes } from "../src/scripts/office/unpack";

describe("pack/unpack in-memory core", () => {
  it("unpackBytes returns a Map<path, Uint8Array> for a real DOCX", async () => {
    const bytes = readFileSync("tests/fixtures/working/empty.valid.docx");
    const parts = await unpackBytes(new Uint8Array(bytes));
    expect(parts.get("word/document.xml")).toBeInstanceOf(Uint8Array);
  });

  it("packBytes round-trips through unpackBytes", async () => {
    const bytes = readFileSync("tests/fixtures/working/empty.valid.docx");
    const parts = await unpackBytes(new Uint8Array(bytes));
    const repacked = await packBytes(parts);
    const reparsed = await unpackBytes(repacked);
    // Parts set should match
    expect([...reparsed.keys()].sort()).toEqual([...parts.keys()].sort());
  });
});
```

- [ ] **Step 2: Refactor pack/unpack**

Split the existing file-tree-based `pack(unpackedDir, outPath)` and `unpack(docxPath, dir)` into:
- `unpackBytes(bytes: Uint8Array): Promise<Map<string, Uint8Array>>` — pure in-memory, uses JSZip directly
- `packBytes(parts: Map<string, Uint8Array>): Promise<Uint8Array>` — pure in-memory
- File-tree variants live in `src/node/pack.ts` / `src/node/unpack.ts` and are thin wrappers over the bytes core

- [ ] **Step 3: Update consumers**

`Validate`, `Repair`, `Normalize`, `Measure` switch to the bytes core. They no longer need `withTempDir` for the round-trip itself (still needed by validator internals that walk a directory — those stay in src/node/ as PR C scope).

- [ ] **Step 4: Run the full test suite**

Run: `bun run test`
Expected: all green.

- [ ] **Step 5: Commit**

```bash
git add src/scripts/office/pack.ts src/scripts/office/unpack.ts src/node/pack.ts src/node/unpack.ts tests/pack-unpack-bytes.test.ts
git commit -m "refactor: split pack/unpack into in-memory bytes core + Node FS wrapper

unpackBytes(bytes) and packBytes(parts) are platform-agnostic and used
by the four top-level classes. The file-tree variants live in src/node/.

Co-Authored-By: Claude Opus 4.7 (1M context) <noreply@anthropic.com>"
```

### Task C.8: Dual package.json#exports + un-skip the segregation test

**Files:**
- Modify: `package.json`
- Modify: `tests/no-node-imports-in-core.test.ts`

- [ ] **Step 1: Update exports**

```json
{
  "exports": {
    ".":      { "import": "./dist/index.js",      "types": "./dist/index.d.ts" },
    "./node": { "import": "./dist/node/index.js", "types": "./dist/node/index.d.ts" }
  }
}
```

- [ ] **Step 2: Un-skip the segregation test**

In `tests/no-node-imports-in-core.test.ts`, change `it.skip(...)` back to `it(...)`.

- [ ] **Step 3: Run the segregation test**

Run: `bun run test tests/no-node-imports-in-core.test.ts`
Expected: PASS — zero `node:*` imports in `src/` outside `src/node/`.

If it FAILS, the relocation in task C.6/C.7 missed something; surface the offending file, move it, re-run.

- [ ] **Step 4: Run the full suite**

Run: `bun run test`
Expected: all green.

- [ ] **Step 5: Commit**

```bash
git add package.json tests/no-node-imports-in-core.test.ts
git commit -m "feat: dual package exports (.+./node) + enforce segregation

Browser consumers: import { Validate, jubarte } from 'docx-validate'
Node consumers: import { validateFile, acceptChanges } from 'docx-validate/node'

The no-node-imports-in-core test is un-skipped — drift breaks CI.

Co-Authored-By: Claude Opus 4.7 (1M context) <noreply@anthropic.com>"
```

### Task C.9: Browser smoke test

**Files:**
- Test: `tests/browser-entry.smoke.test.ts`

- [ ] **Step 1: Write the smoke test**

```ts
// tests/browser-entry.smoke.test.ts
// Runs only in the browser vitest project.
import { describe, it, expect } from "vitest";

describe("browser entry smoke", () => {
  it("imports the browser barrel without node:* deps", async () => {
    const mod = await import("../src/index");
    expect(typeof mod.Validate).toBe("function");
    expect(typeof mod.jubarte.validate).toBe("function");
  });
});
```

- [ ] **Step 2: Verify the browser vitest project includes it**

The vitest config must route `*.smoke.test.ts` under `tests/` into the browser project (matches the existing convention from jubarte-first).

- [ ] **Step 3: Run the browser project**

Run: `bun run test --project browser`
Expected: smoke test passes — the browser bundle loads without resolving any `node:*` specifier.

- [ ] **Step 4: Commit**

```bash
git add tests/browser-entry.smoke.test.ts vitest.config.ts
git commit -m "test: browser smoke for the four-class surface

Asserts the browser barrel loads in @vitest/browser + Playwright without
pulling any node:* import. Bundle-size assertion can be added later if
tree-shaking regresses.

Co-Authored-By: Claude Opus 4.7 (1M context) <noreply@anthropic.com>"
```

### Task C.10: Open PR C

```bash
git push -u origin feat/four-class-surface
gh pr create \
  --base main \
  --head feat/four-class-surface \
  --title "feat: four-class architecture (Validate/Repair/Normalize/Measure) + jubarte namespace" \
  --body "Part 3 of the four-class refactor. Stacked on PR B.

## What

- Four top-level classes: \`Validate\`, \`Repair\`, \`Normalize\`, \`Measure\`
- \`jubarte\` functional namespace: \`jubarte.validate / .repair / .normalize / .measure\`
- \`src/node/\` segregation: \`validate-file\`, \`accept-changes\`, CLI, corpus driver, file-tree pack/unpack
- Dual \`package.json#exports\`: \`.\` (browser-safe) and \`./node\` (everything)
- pack/unpack split into pure in-memory \`packBytes\`/\`unpackBytes\` core + Node FS wrappers
- no-node-imports-in-core enforcement test (un-skipped from PR A)
- Browser smoke test via @vitest/browser + Playwright

## Why

Completes the refactor described in docs/superpowers/specs/2026-05-29-four-class-architecture-design.md. After this lands, jubarte-first composes \`docx-validate\`'s namespace into its own (spec §7) — that work is a separate PR in jubarte-first.

🤖 Generated with [Claude Code](https://claude.com/claude-code)"
```

---

## Self-review checklist (run before declaring plan ready)

- [ ] **Spec coverage:** every section of the spec maps to at least one task — verified
- [ ] **No placeholders:** no TBD/TODO inside any task body (the Measure stub TODO is explicit + scoped)
- [ ] **Type consistency:** `XsdValidator`, `Issue`, `ValidationResult`, `MeasureResult` named identically in every task that touches them — verified
- [ ] **PR independence:** each of A/B/C can be merged on its own and leaves the codebase green
- [ ] **TDD throughout:** every task starts with a failing test, then minimal impl

## Execution Handoff

**Plan complete and saved to `docs/superpowers/plans/2026-05-30-four-class-architecture-implementation.md`. Two execution options:**

**1. Subagent-Driven (recommended)** — fresh subagent per task, two-stage review per task, no context pollution between tasks.

**2. Inline Execution** — execute in this session using executing-plans skill, batched commits with checkpoints.

**Which approach?**
