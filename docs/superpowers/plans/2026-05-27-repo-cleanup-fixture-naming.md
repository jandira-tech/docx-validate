# Repo Cleanup + Content-Descriptive Fixture Naming Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Remove repo-root cruft and give every undescriptive `.docx` fixture a deterministic content-derived name (`<subject>.<comment-or-error>.docx`), folding the 151 unreferenced `fixtures/eigen-extended/` specimens into the exercised `tests/fixtures/` corpus.

**Architecture:** A pure, unit-tested name-derivation function (`deriveName`) consumes a content fingerprint (`fingerprint`, built on JSZip + the existing `validate()`); a thin driver (`apply-fixture-names.ts`) wires them to `git mv`, dedup, and category-sorting. `update-manifest.ts` is hardened to preserve `word` metadata across regens. Three stacked PRs: tooling → rename strays → eigen triage.

**Tech Stack:** TypeScript (ESM), bun, vitest, JSZip, `@xmldom/xmldom` via `src/lib/xml-helpers.ts`, `bunx tsx` for scripts.

---

## File structure

| File                                              | Responsibility                                                                                      |
| ------------------------------------------------- | --------------------------------------------------------------------------------------------------- |
| `scripts/derive-fixture-name.ts`                  | **Pure.** `FixtureFingerprint` + `DerivedName` types, `slugify`, `deriveName`. The TDD core.        |
| `scripts/fixture-fingerprint.ts`                  | IO. `fingerprint(docxPath)` → `FixtureFingerprint` (JSZip read + `validate()`).                     |
| `scripts/apply-fixture-names.ts`                  | Driver CLI: fingerprint → derive → dedup → `git mv` (in-place or `--into-categories`), `--dry-run`. |
| `scripts/update-manifest.ts`                      | **Modify.** Preserve prior `word` values when probe data is absent.                                 |
| `tests/derive-fixture-name.test.ts`               | Unit tests for `deriveName`/`slugify`.                                                              |
| `tests/fixture-fingerprint.test.ts`               | Integration test for `fingerprint` against real fixtures.                                           |
| `tests/update-manifest-word-preservation.test.ts` | Unit test for `word` preservation.                                                                  |

The branch `chore/repo-cleanup` (off `main`) already exists and holds the design-spec commit. PR 1 continues on it.

---

## PR 1 — tooling + root cruft (branch: `chore/repo-cleanup`)

### Task 1: Remove root cruft

**Files:**

- Delete: `Sample Document.repaired.docx` (git-tracked)
- Delete: `package-lock.json` (untracked, gitignored npm leftover)

- [ ] **Step 1: Confirm nothing references them**

Run: `grep -rn "Sample Document.repaired" src tests scripts; git ls-files package-lock.json`
Expected: no `src/tests/scripts` hits; `git ls-files` prints nothing (untracked).

- [ ] **Step 2: Delete**

```bash
git rm "Sample Document.repaired.docx"
rm -f package-lock.json
```

- [ ] **Step 3: Verify build/tests still green**

Run: `bunx tsc --noEmit && bun run test`
Expected: PASS (these files were not imported anywhere).

- [ ] **Step 4: Commit**

```bash
git add -A
git commit -m "chore: remove tracked repair artifact and stray npm lockfile"
```

---

### Task 2: Harden `update-manifest.ts` to preserve `word` metadata

**Files:**

- Modify: `scripts/update-manifest.ts` (the `main()` `word` derivation, ~line 99-107)
- Test: `tests/update-manifest-word-preservation.test.ts`

The current code sets `word: probeRecord?.word?.outcome ?? "unknown"`. With the probe file gone, regen downgrades every `word` to `"unknown"`. We add a fallback to the existing manifest's value. To make this testable, extract the resolution into an exported pure function.

- [ ] **Step 1: Write the failing test**

Create `tests/update-manifest-word-preservation.test.ts`:

```typescript
/*
 * Copyright 2026 Jandira Technologies, LLC
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *     http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */

import { describe, expect, it } from "vitest";

import { resolveWordOutcome } from "../scripts/update-manifest";

describe("resolveWordOutcome", () => {
  const prior = new Map<string, string>([["broken/a.docx", "clean-open"]]);

  it("prefers a fresh probe outcome", () => {
    expect(resolveWordOutcome("broken/a.docx", "recovered", prior)).toBe("recovered");
  });

  it("falls back to the prior manifest value when no probe", () => {
    expect(resolveWordOutcome("broken/a.docx", undefined, prior)).toBe("clean-open");
  });

  it("is 'unknown' when neither probe nor prior value exists", () => {
    expect(resolveWordOutcome("broken/new.docx", undefined, prior)).toBe("unknown");
  });
});
```

- [ ] **Step 2: Run test to verify it fails**

Run: `bun run test -- update-manifest-word-preservation`
Expected: FAIL — `resolveWordOutcome` is not exported.

- [ ] **Step 3: Implement `resolveWordOutcome` and wire it into `main`**

In `scripts/update-manifest.ts`, add the exported helper (near the other top-level functions):

```typescript
/**
 * Resolve the `word` outcome for a fixture: a fresh probe result wins; else
 * the value carried in the previous manifest is preserved; else "unknown".
 * Keeps regen idempotent w.r.t. Word-probe metadata when the probe JSONL is
 * absent (no LibreOffice in CI).
 */
export function resolveWordOutcome(
  relativePath: string,
  probeOutcome: string | undefined,
  priorWord: Map<string, string>,
): string {
  return probeOutcome ?? priorWord.get(relativePath) ?? "unknown";
}

function readPriorManifestWord(): Map<string, string> {
  const map = new Map<string, string>();
  if (!existsSync(MANIFEST)) return map;
  try {
    const prior = JSON.parse(readFileSync(MANIFEST, "utf-8")) as {
      entries?: { relativePath: string; word?: string }[];
    };
    for (const entry of prior.entries ?? []) {
      if (entry.word) map.set(entry.relativePath, entry.word);
    }
  } catch {
    // Malformed prior manifest — treat as no prior data.
  }
  return map;
}
```

Then in `main()`, replace the `word` derivation block:

```typescript
const probeResults = readProbeResults();
const priorWord = readPriorManifestWord();
const entries: ManifestEntry[] = [];

for (const file of files) {
  const relativePath = path.relative(FIXTURES_ROOT, file);
  process.stderr.write(`Processing: ${relativePath}\n`);

  const strict = await runValidator(file, "strict");
  const lenient = await runValidator(file, "lenient");

  const probeRecord = probeResults.get(relativePath);
  const wordOutcome = resolveWordOutcome(relativePath, probeRecord?.word?.outcome, priorWord);

  entries.push({ relativePath, strict, lenient, word: wordOutcome });
}
```

- [ ] **Step 4: Run test to verify it passes**

Run: `bun run test -- update-manifest-word-preservation`
Expected: PASS (3 tests).

- [ ] **Step 5: Commit**

```bash
git add scripts/update-manifest.ts tests/update-manifest-word-preservation.test.ts
git commit -m "fix: preserve word metadata when regenerating fixtures manifest"
```

---

### Task 3: `deriveName` — pure name derivation

**Files:**

- Create: `scripts/derive-fixture-name.ts`
- Test: `tests/derive-fixture-name.test.ts`

- [ ] **Step 1: Write the failing test**

Create `tests/derive-fixture-name.test.ts`:

```typescript
/*
 * Copyright 2026 Jandira Technologies, LLC
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *     http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */

import { describe, expect, it } from "vitest";

import { type FixtureFingerprint, deriveName, slugify } from "../scripts/derive-fixture-name";

const base: FixtureFingerprint = {
  strictErrorCodes: [],
  lenientErrorCodes: [],
  insCount: 0,
  delCount: 0,
  commentCount: 0,
  firstCommentText: null,
  tableCount: 0,
  hasTextBox: false,
  hasHeaderFooter: false,
  titleText: null,
  contentHash: "deadbeef",
};

describe("slugify", () => {
  it("lowercases, hyphenates, and caps word count", () => {
    expect(slugify("Q4 2025 NPS Survey Results Extra", 4)).toBe("q4-2025-nps-survey");
  });
  it("returns empty string for punctuation-only input", () => {
    expect(slugify("!!! ???")).toBe("");
  });
});

describe("deriveName", () => {
  it("routes a file with strict errors to broken/ using the first error code", () => {
    const d = deriveName({ ...base, strictErrorCodes: ["tables-broken-rels", "another-code"] });
    expect(d.category).toBe("broken");
    expect(d.descriptor).toBe("another-code"); // alphabetically first
    expect(d.fileName).toBe("document.another-code.docx");
  });

  it("describes insertions-only tracked changes", () => {
    const d = deriveName({ ...base, insCount: 3 });
    expect(d.category).toBe("working");
    expect(d.fileName).toBe("document.suggesting-insertions.docx");
  });

  it("describes deletions-only tracked changes", () => {
    const d = deriveName({ ...base, delCount: 2 });
    expect(d.fileName).toBe("document.suggesting-deletions.docx");
  });

  it("describes mixed tracked changes", () => {
    const d = deriveName({ ...base, insCount: 1, delCount: 1 });
    expect(d.fileName).toBe("document.suggesting-mixed-edits.docx");
  });

  it("describes a comment using its slugified gist", () => {
    const d = deriveName({
      ...base,
      commentCount: 1,
      firstCommentText: "Please review this clause",
    });
    expect(d.fileName).toBe("document.comment-please-review-this-clause.docx");
  });

  it("falls back to a structural descriptor for a clean table doc", () => {
    const d = deriveName({ ...base, tableCount: 2 });
    expect(d.fileName).toBe("document.table.docx");
  });

  it("uses the slugified title as the subject when present", () => {
    const d = deriveName({ ...base, titleText: "Master Services Agreement", insCount: 1 });
    expect(d.fileName).toBe("master-services-agreement.suggesting-insertions.docx");
  });

  it("falls back to plain-paragraphs for an otherwise featureless clean doc", () => {
    expect(deriveName(base).fileName).toBe("document.plain-paragraphs.docx");
  });
});
```

- [ ] **Step 2: Run test to verify it fails**

Run: `bun run test -- derive-fixture-name`
Expected: FAIL — module not found.

- [ ] **Step 3: Implement `scripts/derive-fixture-name.ts`**

```typescript
/*
 * Copyright 2026 Jandira Technologies, LLC
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *     http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */

/**
 * Pure, deterministic derivation of a content-descriptive fixture name from a
 * docx content fingerprint. See
 * docs/superpowers/specs/2026-05-27-repo-cleanup-fixture-naming-design.md.
 */

export interface FixtureFingerprint {
  /** Distinct, sorted strict-profile error codes. */
  strictErrorCodes: string[];
  /** Distinct, sorted lenient-profile error codes. */
  lenientErrorCodes: string[];
  insCount: number;
  delCount: number;
  commentCount: number;
  firstCommentText: string | null;
  tableCount: number;
  hasTextBox: boolean;
  hasHeaderFooter: boolean;
  /** dc:title or first non-empty paragraph text, if any. */
  titleText: string | null;
  /** sha256 of word/document.xml, for dedup. */
  contentHash: string;
}

export interface DerivedName {
  category: "broken" | "working";
  subjectSlug: string;
  descriptor: string;
  /** `${subjectSlug}.${descriptor}.docx` */
  fileName: string;
}

/** Lowercase kebab slug, capped at `maxWords` hyphen-delimited words. */
export function slugify(text: string, maxWords = 4): string {
  const words = text
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, "-")
    .replace(/^-+|-+$/g, "")
    .split("-")
    .filter(Boolean);
  return words.slice(0, maxWords).join("-");
}

function descriptorFor(fp: FixtureFingerprint): string {
  if (fp.strictErrorCodes.length > 0) {
    return [...fp.strictErrorCodes].sort()[0];
  }
  if (fp.insCount > 0 && fp.delCount > 0) return "suggesting-mixed-edits";
  if (fp.insCount > 0) return "suggesting-insertions";
  if (fp.delCount > 0) return "suggesting-deletions";
  if (fp.commentCount > 0) {
    const gist = slugify(fp.firstCommentText ?? "", 4);
    return gist ? `comment-${gist}` : "comment";
  }
  if (fp.tableCount > 0) return "table";
  if (fp.hasTextBox) return "text-box";
  if (fp.hasHeaderFooter) return "header-footer";
  return "plain-paragraphs";
}

export function deriveName(fp: FixtureFingerprint): DerivedName {
  const category: "broken" | "working" = fp.strictErrorCodes.length > 0 ? "broken" : "working";
  const subjectSlug = (fp.titleText && slugify(fp.titleText)) || "document";
  const descriptor = descriptorFor(fp);
  return { category, subjectSlug, descriptor, fileName: `${subjectSlug}.${descriptor}.docx` };
}
```

- [ ] **Step 4: Run test to verify it passes**

Run: `bun run test -- derive-fixture-name`
Expected: PASS (all cases).

- [ ] **Step 5: Commit**

```bash
git add scripts/derive-fixture-name.ts tests/derive-fixture-name.test.ts
git commit -m "feat: deterministic content-descriptive fixture name derivation"
```

---

### Task 4: `fingerprint` — content fingerprint from a docx

**Files:**

- Create: `scripts/fixture-fingerprint.ts`
- Test: `tests/fixture-fingerprint.test.ts`

- [ ] **Step 1: Write the failing test**

Create `tests/fixture-fingerprint.test.ts`. Uses real fixtures that exist on `main`: `broken/tables.missing-namespace.docx` (a known strict failure) and `external/mammoth-js/strict-format.docx` (a structurally-valid doc).

```typescript
/*
 * Copyright 2026 Jandira Technologies, LLC
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *     http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */

import path from "node:path";
import { fileURLToPath } from "node:url";

import { describe, expect, it } from "vitest";

import { fingerprint } from "../scripts/fixture-fingerprint";

const HERE = path.dirname(fileURLToPath(import.meta.url));
const FIXTURES = path.join(HERE, "fixtures");

describe("fingerprint", () => {
  it("captures strict errors and a content hash for a known-broken fixture", async () => {
    const fp = await fingerprint(path.join(FIXTURES, "broken/tables.missing-namespace.docx"));
    expect(fp.strictErrorCodes.length).toBeGreaterThan(0);
    expect(fp.contentHash).toMatch(/^[0-9a-f]{64}$/);
    // sorted + de-duplicated
    expect([...fp.strictErrorCodes].sort()).toEqual(fp.strictErrorCodes);
  }, 20000);

  it("reports zero strict errors for a structurally-valid fixture", async () => {
    const fp = await fingerprint(path.join(FIXTURES, "external/mammoth-js/strict-format.docx"));
    expect(fp.strictErrorCodes).toEqual([]);
    expect(fp.contentHash).toMatch(/^[0-9a-f]{64}$/);
  }, 20000);
});
```

- [ ] **Step 2: Run test to verify it fails**

Run: `bun run test -- fixture-fingerprint`
Expected: FAIL — module not found.

> Note: if `external/mammoth-js/strict-format.docx` reports strict errors on this branch, swap the second fixture for any file whose manifest entry shows `strict.valid === true` (check `tests/fixtures-all.manifest.json`). Adjust the assertion, don't weaken the contract.

- [ ] **Step 3: Implement `scripts/fixture-fingerprint.ts`**

```typescript
/*
 * Copyright 2026 Jandira Technologies, LLC
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *     http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */

import { createHash } from "node:crypto";
import { promises as fs } from "node:fs";

import JSZip from "jszip";

import { NS } from "../src/lib/types";
import { getElementsByTagNameNSAll, parseXml } from "../src/lib/xml-helpers";
import { validate } from "../src/scripts/office/validate";
import type { FixtureFingerprint } from "./derive-fixture-name";

const DC = "http://purl.org/dc/elements/1.1/";

async function errorCodes(file: string, profile: "strict" | "lenient"): Promise<string[]> {
  const result = await validate(file, { profile });
  return Array.from(
    new Set(
      result.issues
        .filter((i) => i.severity === "error")
        .map((i) => i.code)
        .filter((c): c is string => Boolean(c)),
    ),
  ).sort();
}

function firstParagraphText(doc: Document): string | null {
  const paras = getElementsByTagNameNSAll(doc, NS.W, "p");
  for (const p of paras) {
    const text = getElementsByTagNameNSAll(p, NS.W, "t")
      .map((t) => t.textContent ?? "")
      .join("")
      .trim();
    if (text) return text;
  }
  return null;
}

export async function fingerprint(docxPath: string): Promise<FixtureFingerprint> {
  const buf = await fs.readFile(docxPath);
  const zip = await JSZip.loadAsync(buf);

  const docXml = (await zip.file("word/document.xml")?.async("string")) ?? "";
  const contentHash = createHash("sha256").update(docXml).digest("hex");
  const doc = parseXml(docXml);

  const insCount = getElementsByTagNameNSAll(doc, NS.W, "ins").length;
  const delCount = getElementsByTagNameNSAll(doc, NS.W, "del").length;
  const tableCount = getElementsByTagNameNSAll(doc, NS.W, "tbl").length;
  const hasTextBox = docXml.includes("txbxContent") || docXml.includes("textbox");
  const hasHeaderFooter = zip.file(/word\/(header|footer)\d*\.xml$/).length > 0;

  let commentCount = 0;
  let firstCommentText: string | null = null;
  const commentsXml = await zip.file("word/comments.xml")?.async("string");
  if (commentsXml) {
    const cdoc = parseXml(commentsXml);
    const comments = getElementsByTagNameNSAll(cdoc, NS.W, "comment");
    commentCount = comments.length;
    if (comments[0]) {
      const text = getElementsByTagNameNSAll(comments[0], NS.W, "t")
        .map((t) => t.textContent ?? "")
        .join("")
        .trim();
      firstCommentText = text || null;
    }
  }

  let titleText: string | null = null;
  const coreXml = await zip.file("docProps/core.xml")?.async("string");
  if (coreXml) {
    const core = parseXml(coreXml);
    const title = getElementsByTagNameNSAll(core, DC, "title")[0]?.textContent?.trim();
    if (title) titleText = title;
  }
  if (!titleText) titleText = firstParagraphText(doc);

  const [strictErrorCodes, lenientErrorCodes] = await Promise.all([
    errorCodes(docxPath, "strict"),
    errorCodes(docxPath, "lenient"),
  ]);

  return {
    strictErrorCodes,
    lenientErrorCodes,
    insCount,
    delCount,
    commentCount,
    firstCommentText,
    tableCount,
    hasTextBox,
    hasHeaderFooter,
    titleText,
    contentHash,
  };
}
```

- [ ] **Step 4: Run test to verify it passes**

Run: `bun run test -- fixture-fingerprint`
Expected: PASS (2 tests).

- [ ] **Step 5: Commit**

```bash
git add scripts/fixture-fingerprint.ts tests/fixture-fingerprint.test.ts
git commit -m "feat: docx content fingerprint for fixture naming"
```

---

### Task 5: `apply-fixture-names.ts` driver

**Files:**

- Create: `scripts/apply-fixture-names.ts`
- Test: `tests/apply-fixture-names.test.ts`

The driver fingerprints each `.docx` under the given paths, derives a target name, dedups by `contentHash` (keeping the lexicographically-first source path), and prints/applies the moves. Export `planMoves` (pure given fingerprints) for testing; keep `git mv` execution in `main`.

- [ ] **Step 1: Write the failing test**

Create `tests/apply-fixture-names.test.ts`:

```typescript
/*
 * Copyright 2026 Jandira Technologies, LLC
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *     http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */

import { describe, expect, it } from "vitest";

import { type FingerprintedFile, planMoves } from "../scripts/apply-fixture-names";

const fp = (overrides: Partial<FingerprintedFile["fingerprint"]> = {}) => ({
  strictErrorCodes: [] as string[],
  lenientErrorCodes: [] as string[],
  insCount: 0,
  delCount: 0,
  commentCount: 0,
  firstCommentText: null,
  tableCount: 0,
  hasTextBox: false,
  hasHeaderFooter: false,
  titleText: null,
  contentHash: "h0",
  ...overrides,
});

describe("planMoves", () => {
  it("renames in place (keeps source dir) when intoCategories is false", () => {
    const moves = planMoves(
      [
        {
          sourcePath: "tests/fixtures/word-strict/Ouch.docx",
          fingerprint: fp({ insCount: 1, contentHash: "a" }),
        },
      ],
      { intoCategories: false, fixturesRoot: "tests/fixtures" },
    );
    expect(moves).toEqual([
      {
        from: "tests/fixtures/word-strict/Ouch.docx",
        to: "tests/fixtures/word-strict/document.suggesting-insertions.docx",
      },
    ]);
  });

  it("routes into category dirs when intoCategories is true", () => {
    const moves = planMoves(
      [
        {
          sourcePath: "fixtures/eigen-extended/Untitled (1).docx",
          fingerprint: fp({ strictErrorCodes: ["x-code"], contentHash: "b" }),
        },
      ],
      { intoCategories: true, fixturesRoot: "tests/fixtures" },
    );
    expect(moves[0].to).toBe("tests/fixtures/broken/document.x-code.docx");
  });

  it("disambiguates colliding target names with a numeric suffix", () => {
    const moves = planMoves(
      [
        { sourcePath: "a.docx", fingerprint: fp({ insCount: 1, contentHash: "c1" }) },
        { sourcePath: "b.docx", fingerprint: fp({ insCount: 1, contentHash: "c2" }) },
      ],
      { intoCategories: true, fixturesRoot: "tests/fixtures" },
    );
    expect(moves.map((m) => m.to)).toEqual([
      "tests/fixtures/working/document.suggesting-insertions.docx",
      "tests/fixtures/working/document.suggesting-insertions-2.docx",
    ]);
  });

  it("drops content duplicates, keeping the first source path", () => {
    const moves = planMoves(
      [
        { sourcePath: "b.docx", fingerprint: fp({ insCount: 1, contentHash: "dup" }) },
        { sourcePath: "a.docx", fingerprint: fp({ insCount: 1, contentHash: "dup" }) },
      ],
      { intoCategories: true, fixturesRoot: "tests/fixtures", dedup: true },
    );
    expect(moves).toHaveLength(1);
    expect(moves[0].from).toBe("a.docx"); // lexicographically first kept
  });
});
```

- [ ] **Step 2: Run test to verify it fails**

Run: `bun run test -- apply-fixture-names`
Expected: FAIL — module not found.

- [ ] **Step 3: Implement `scripts/apply-fixture-names.ts`**

```typescript
/*
 * Copyright 2026 Jandira Technologies, LLC
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *     http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */

/**
 * Fingerprint a set of .docx files, derive content-descriptive names, and
 * `git mv` them — either in place (rename only) or sorted into
 * tests/fixtures/<category>/. Dedups by content hash. Supports --dry-run.
 *
 *   bunx tsx scripts/apply-fixture-names.ts [--into-categories] [--dedup] \
 *     [--dry-run] [--fixtures-root tests/fixtures] <path>...
 */

import { execFileSync } from "node:child_process";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { readdirSync, statSync } from "node:fs";

import { type FixtureFingerprint, deriveName } from "./derive-fixture-name";
import { fingerprint } from "./fixture-fingerprint";

export interface FingerprintedFile {
  sourcePath: string;
  fingerprint: FixtureFingerprint;
}

export interface PlanOptions {
  intoCategories: boolean;
  fixturesRoot: string;
  dedup?: boolean;
}

export interface Move {
  from: string;
  to: string;
}

/** Pure: turn fingerprinted files into a deduped, collision-free move list. */
export function planMoves(files: FingerprintedFile[], opts: PlanOptions): Move[] {
  const seenHash = new Set<string>();
  const takenTargets = new Set<string>();
  const moves: Move[] = [];

  // Deterministic order; dedup keeps the lexicographically-first source.
  const sorted = [...files].sort((a, b) => a.sourcePath.localeCompare(b.sourcePath));

  for (const file of sorted) {
    if (opts.dedup) {
      if (seenHash.has(file.fingerprint.contentHash)) continue;
      seenHash.add(file.fingerprint.contentHash);
    }
    const derived = deriveName(file.fingerprint);
    const dir = opts.intoCategories
      ? path.join(opts.fixturesRoot, derived.category)
      : path.dirname(file.sourcePath);

    let candidate = path.join(dir, derived.fileName);
    let n = 2;
    while (takenTargets.has(candidate)) {
      candidate = path.join(dir, `${derived.subjectSlug}.${derived.descriptor}-${n}.docx`);
      n += 1;
    }
    takenTargets.add(candidate);
    moves.push({ from: file.sourcePath, to: candidate });
  }
  return moves;
}

function collectDocx(target: string, out: string[]): void {
  const stat = statSync(target);
  if (stat.isDirectory()) {
    for (const entry of readdirSync(target)) {
      if (entry.startsWith("~$")) continue;
      collectDocx(path.join(target, entry), out);
    }
  } else if (target.toLowerCase().endsWith(".docx")) {
    out.push(target);
  }
}

async function main(): Promise<void> {
  const argv = process.argv.slice(2);
  const intoCategories = argv.includes("--into-categories");
  const dedup = argv.includes("--dedup");
  const dryRun = argv.includes("--dry-run");
  const rootIdx = argv.indexOf("--fixtures-root");
  const fixturesRoot = rootIdx >= 0 ? argv[rootIdx + 1] : "tests/fixtures";
  const flagValues = new Set([fixturesRoot]);
  const targets = argv.filter(
    (a, i) =>
      !a.startsWith("--") && !(i > 0 && argv[i - 1] === "--fixtures-root") && !flagValues.has(a),
  );

  const docxPaths: string[] = [];
  for (const t of targets) collectDocx(t, docxPaths);

  const files: FingerprintedFile[] = [];
  for (const p of docxPaths) {
    files.push({ sourcePath: p, fingerprint: await fingerprint(p) });
  }

  const moves = planMoves(files, { intoCategories, fixturesRoot, dedup });
  const dropped = files.length - moves.length;

  for (const m of moves) {
    process.stdout.write(`${m.from}  ->  ${m.to}\n`);
    if (!dryRun) {
      execFileSync("git", ["mv", m.from, m.to]);
    }
  }
  process.stdout.write(
    `\n${moves.length} renamed, ${dropped} dropped as duplicates${dryRun ? " (dry-run)" : ""}\n`,
  );
}

if (import.meta.url === `file://${process.argv[1]}`) {
  main().catch((err) => {
    process.stderr.write(`Error: ${err}\n`);
    process.exit(1);
  });
}
```

- [ ] **Step 4: Run test to verify it passes**

Run: `bun run test -- apply-fixture-names`
Expected: PASS (4 tests).

- [ ] **Step 5: Full verification + commit**

Run: `bunx tsc --noEmit && bun run test`
Expected: whole suite PASS.

```bash
git add scripts/apply-fixture-names.ts tests/apply-fixture-names.test.ts
git commit -m "feat: fixture rename/sort driver with dedup and collision handling"
```

- [ ] **Step 6: Open PR 1**

```bash
git push -u origin chore/repo-cleanup
gh pr create --base main --title "chore: repo cleanup tooling + root cruft removal" \
  --body "First of three stacked PRs. Removes root cruft, hardens manifest regen to preserve word metadata, and adds the tested fingerprint/derive/driver tooling used by the next two PRs. Spec: docs/superpowers/specs/2026-05-27-repo-cleanup-fixture-naming-design.md"
```

---

## PR 2 — rename the 3 `tests/fixtures/` strays (branch: `chore/repo-cleanup-renames`)

Stacked on PR 1.

### Task 6: Rename strays in place + regen manifest

**Files:**

- Rename: `tests/fixtures/vfdsdfcACawesd.docx`
- Rename: `tests/fixtures/word-strict/second-pass/Ouch.docx`
- Rename: `tests/fixtures/word-strict/tests:fixtures:broken:sample-document.broken-tables.docx-repaired.docx`
- Modify: `tests/fixtures-all-strict.test.ts` (doc-comment line ~20)
- Modify (generated): `tests/fixtures-all.manifest.json`

- [ ] **Step 1: Branch off PR 1**

```bash
git checkout chore/repo-cleanup
git checkout -b chore/repo-cleanup-renames
```

- [ ] **Step 2: Dry-run the driver on the three strays**

```bash
bunx tsx scripts/apply-fixture-names.ts --dry-run \
  "tests/fixtures/vfdsdfcACawesd.docx" \
  "tests/fixtures/word-strict/second-pass/Ouch.docx" \
  "tests/fixtures/word-strict/tests:fixtures:broken:sample-document.broken-tables.docx-repaired.docx"
```

Expected: three `from -> to` lines, renamed in place (same dirs), with content-derived basenames. Review they read sensibly.

- [ ] **Step 3: Confirm no test hard-codes these basenames (other than the doc-comment)**

Run: `grep -rn "vfdsdfcACawesd\|Ouch.docx\|tests:fixtures:broken:sample-document" src tests scripts | grep -v fixtures-all.manifest.json`
Expected: only `tests/fixtures-all-strict.test.ts:20` (the doc-comment).

- [ ] **Step 4: Apply the renames**

Re-run the Step 2 command without `--dry-run`. Note the three new basenames it prints.

- [ ] **Step 5: Update the test doc-comment**

In `tests/fixtures-all-strict.test.ts`, replace `plus the lone vfdsdfcACawesd.docx).` with `plus the renamed root specimen <new-basename>).` using the new basename printed in Step 4.

- [ ] **Step 6: Regenerate the manifest**

Run: `bunx tsx scripts/update-manifest.ts`
Then inspect: `git diff --stat tests/fixtures-all.manifest.json` and `git diff tests/fixtures-all.manifest.json | grep '"word"' | head`
Expected: only `relativePath` keys for the three renamed files changed; `word` values unchanged elsewhere (hardening from Task 2). `totalFixtures` unchanged.

- [ ] **Step 7: Verify**

Run: `bun run test -- fixtures-all && bunx tsc --noEmit`
Expected: both strict + lenient fixtures suites PASS.

- [ ] **Step 8: Commit + PR**

```bash
git add -A
git commit -m "chore: rename undescriptive tests/fixtures strays to content-derived names"
git push -u origin chore/repo-cleanup-renames
gh pr create --base chore/repo-cleanup --title "chore: rename tests/fixtures strays" \
  --body "Second of three stacked PRs (over chore/repo-cleanup). Renames vfdsdfcACawesd / Ouch / the colon-path file to content-derived names and regenerates the manifest."
```

---

## PR 3 — eigen-extended triage (branch: `chore/repo-cleanup-eigen`)

Stacked on PR 2.

### Task 7: Dedup, rename, sort eigen-extended into the corpus

**Files:**

- Move+rename: all of `fixtures/eigen-extended/*.docx` → `tests/fixtures/{broken,working}/`
- Delete: empty `fixtures/` directory
- Modify (generated): `tests/fixtures-all.manifest.json`
- Modify: `tests/fixtures/broken/README.md`, `tests/fixtures/working/README.md`
- Modify: `tests/fixtures-all-strict.test.ts` + `tests/fixtures-all-lenient.test.ts` doc-comments

- [ ] **Step 1: Branch off PR 2**

```bash
git checkout chore/repo-cleanup-renames
git checkout -b chore/repo-cleanup-eigen
```

- [ ] **Step 2: Dry-run the driver on the whole eigen folder**

```bash
bunx tsx scripts/apply-fixture-names.ts --into-categories --dedup --dry-run \
  fixtures/eigen-extended
```

Expected: ~148 `from -> to` lines into `tests/fixtures/broken/` and `tests/fixtures/working/`, plus `~3 dropped as duplicates`. Skim for obviously-wrong names; the logic is deterministic so anomalies indicate a fingerprint edge case worth a follow-up, not a blocker.

- [ ] **Step 3: Pre-flight collision check against existing corpus**

The driver only dedups within its input. Confirm no derived target collides with a pre-existing file:

```bash
bunx tsx scripts/apply-fixture-names.ts --into-categories --dedup --dry-run fixtures/eigen-extended \
  | grep ' -> ' | sed 's/.* -> //' | while read t; do [ -e "$t" ] && echo "COLLISION: $t"; done
```

Expected: no `COLLISION:` lines. If any appear, they are pre-existing corpus files — re-run is non-destructive; resolve by bumping the colliding eigen file's name manually after apply, or extend `planMoves` to seed `takenTargets` from disk (follow-up).

- [ ] **Step 4: Delete the 3 content-duplicate files first (so dedup choice is explicit in history)**

The driver drops dups silently from the move list, leaving them unstaged in `fixtures/eigen-extended/`. Apply the moves, then remove whatever remains.

- [ ] **Step 5: Apply**

Re-run Step 2 without `--dry-run`. This `git mv`s survivors into the category dirs.

- [ ] **Step 6: Remove duplicate leftovers and the now-empty source dir**

```bash
git rm -r fixtures/eigen-extended   # removes the 3 dedup leftovers still sitting here
rmdir fixtures 2>/dev/null || true
```

Expected: `fixtures/` no longer exists; the duplicate files are deleted in this commit.

- [ ] **Step 7: Regenerate the manifest**

Run: `bunx tsx scripts/update-manifest.ts`
Then: `grep '"totalFixtures"' tests/fixtures-all.manifest.json`
Expected: `totalFixtures` increased by the number of survivors (~145). New entries carry `"word": "unknown"` (no probe data — acceptable, not asserted by tests); pre-existing entries keep their prior `word` values.

- [ ] **Step 8: Append provenance to the category READMEs**

To `tests/fixtures/broken/README.md` and `tests/fixtures/working/README.md`, append a section:

```markdown
## Imported real-world specimens (eigen)

The files below were imported from the former `fixtures/eigen-extended/`
collection (real-world Plate/SuperDoc exports). They were renamed
deterministically from their content via `scripts/apply-fixture-names.ts`
(see `docs/superpowers/specs/2026-05-27-repo-cleanup-fixture-naming-design.md`).
Their expected validator outcomes are pinned in `tests/fixtures-all.manifest.json`.
```

- [ ] **Step 9: Update the fixtures-all doc-comments**

In both `tests/fixtures-all-strict.test.ts` and `tests/fixtures-all-lenient.test.ts`, extend the source list in the header comment to mention the imported eigen specimens now living under `broken/` and `working/`.

- [ ] **Step 10: Full verification**

Run: `bunx tsc --noEmit && bun run test && bun run check`
Expected: all green. The `fixtures-all-{strict,lenient}` suites now pin every imported specimen.

- [ ] **Step 11: Commit + PR**

```bash
git add -A
git commit -m "chore: fold eigen-extended specimens into tests/fixtures with content-derived names"
git push -u origin chore/repo-cleanup-eigen
gh pr create --base chore/repo-cleanup-renames --title "chore: import eigen-extended fixtures into the corpus" \
  --body "Third of three stacked PRs (over chore/repo-cleanup-renames). Dedups, renames, and sorts the 151 unreferenced eigen-extended specimens into tests/fixtures/{broken,working}/, regenerates the manifest, and updates READMEs. Removes the dead top-level fixtures/ dir."
```

---

## Self-review notes

- **Spec coverage:** root cruft (Task 1) ✓; manifest `word` preservation (Task 2) ✓; naming convention (Task 3) ✓; fingerprint (Task 4) ✓; dedup + sort + collision (Task 5) ✓; 3 strays (Task 6) ✓; eigen triage + READMEs + manifest bump (Task 7) ✓.
- **Determinism caveat:** final eigen names are produced at runtime by the tested driver, not hardcoded here — because they depend on file content this plan does not pre-read. The _logic_ is fully specified and unit-tested in PR 1; PR 2/3 are application + review + verification. This is intentional, not a placeholder.
- **Type consistency:** `FixtureFingerprint`/`DerivedName` defined in Task 3 are imported unchanged by Tasks 4–5; `planMoves`/`Move`/`FingerprintedFile` defined in Task 5 are used unchanged in Tasks 6–7.
- **Known follow-up (out of scope):** `planMoves` only dedups/disambiguates within its input set; cross-corpus collisions are surfaced by the Step-3 pre-flight check rather than auto-resolved.
