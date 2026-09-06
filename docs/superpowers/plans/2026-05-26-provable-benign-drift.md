# Provable-Benign Repair Drift + Repair-Validity Fixes — Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Fix two repair-validity defects (duplicate `w14:paraId`; `diff-docx` crash on unparseable parts) and build a `collectVisibleProjection` proof engine that demonstrates each non-loss repair-drift is invisible to an ordinary Word user.

**Architecture:** A new order-preserving "visible projection" of a DOCX captures only reader-perceptible content (text, visible formatting, tables, image targets, page geometry, hyperlinks, headers/footers, comments) and normalizes away plumbing (paraId/rsid/rels-ids/content-types/namespaces/whitespace). The strict invariant is: **repair must not change the visible projection.** A `prove-benign` script repairs a copy of each fixture and asserts projection invariance, emitting `BENIGN_PROOF.md`.

**Tech Stack:** TypeScript (ESM), `@xmldom/xmldom` via `src/lib/xml-helpers.ts`, JSZip, `vitest`, `bun`.

**Spec:** `docs/superpowers/specs/2026-05-26-provable-benign-drift-design.md`
**Branch:** stacked on `feat/inventory-fingerprint` (work directly on it).

---

## File structure

| File                                                       | Responsibility                                                              |
| ---------------------------------------------------------- | --------------------------------------------------------------------------- |
| `src/scripts/office/validators/docx.ts`                    | **(modify)** Fix `repairMissingParaIds` existence check (Commit 1).         |
| `scripts/diff-docx.ts`                                     | **(modify)** Catch per-input failures; report instead of throw (Commit 2).  |
| `src/scripts/office/validators/docx-visible-projection.ts` | **(new)** `collectVisibleProjection` + `diffVisibleProjections` (Commit 3). |
| `scripts/prove-benign.ts`                                  | **(new)** Repair-and-project survey → `BENIGN_PROOF.md` (Commit 4).         |
| `tests/validators-docx.test.ts`                            | paraId regression test (Commit 1).                                          |
| `tests/diff-docx.cli.test.ts`                              | hardening test (Commit 2).                                                  |
| `tests/docx-visible-projection.test.ts`                    | **(new)** projection unit tests (Commit 3).                                 |

## Reference — existing helpers (reuse, do not reinvent)

From `src/lib/xml-helpers.ts`: `parseXml`, `serializeXml`.
From `src/scripts/office/validators/docx-diagnostics.ts` patterns (replicate the same style if you need them locally): word-namespace iteration via a `WORD_NAMESPACES` set, `directWordChild`, `wordChildAttr`. The two transitional/strict word namespaces are:
`http://schemas.openxmlformats.org/wordprocessingml/2006/main` and `http://purl.oclc.org/ooxml/wordprocessingml/main`.

---

# Commit 1 — Fix duplicate `w14:paraId`

## Task 1: paraId existence check uses a qualified-name fallback

**Root cause:** In `repairMissingParaIds` (`src/scripts/office/validators/docx.ts` ~lines 2413–2414) the existence check is `elem.getAttributeNS(W14_NAMESPACE, "paraId")` where `W14_NAMESPACE = "http://schemas.microsoft.com/office/word/2010/wordml"`. Documents authored against the **older** `w14` namespace (`http://schemas.microsoft.com/office/word/2008/9/12/wordml`, e.g. `external/open-xml-sdk/mcdoc.docx`) bind `w14:paraId` to that older URI, so the 2010-namespace lookup returns `""`. Repair then stamps a second `w14:paraId` via `setAttributeNS(W14_NAMESPACE, "w14:paraId", …)`, producing **two** `w14:paraId` attributes (duplicate qualified name) → invalid XML Word cannot open.

**Files:**

- Modify: `src/scripts/office/validators/docx.ts` (`repairMissingParaIds`, ~lines 2413–2414)
- Test: `tests/validators-docx.test.ts`

- [ ] **Step 1: Write the failing test**

Add to `tests/validators-docx.test.ts` (it already imports `DOCXSchemaValidator`, `withTempDir`, `writeFile`/`path`, `wrapDocument`; if a raw-document helper is needed, write the file directly as below). Use a `<w:p>` whose `w14` prefix is bound to the **2008/9/12** URI:

```ts
it("does not duplicate w14:paraId when the doc binds w14 to the legacy 2008 namespace", async () => {
    await withTempDir(async (dir) => {
        await writeFile(
            path.join(dir, "word", "document.xml"),
            `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
                `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ` +
                `xmlns:w14="http://schemas.microsoft.com/office/word/2008/9/12/wordml">` +
                `<w:body><w:p w14:paraId="57290E37" w14:textId="5B733B31"><w:r><w:t>hi</w:t></w:r></w:p></w:body></w:document>`,
        );
        const v = new DOCXSchemaValidator({ unpackedDir: dir, profile: "strict" });
        await v.repairMissingParaIds();

        const xml = await readFile(path.join(dir, "word", "document.xml"), "utf-8");
        // Must not have stamped a second w14:paraId.
        expect((xml.match(/w14:paraId=/g) ?? []).length).toBe(1);
        // And the result must still parse (duplicate attributes throw "redefined").
        expect(() => parseXml(xml)).not.toThrow();
    });
});
```

Ensure these imports exist at the top of the test file (add any that are missing):

```ts
import { promises as fsp } from "node:fs";
const { readFile } = fsp;
import { parseXml } from "../src/lib/xml-helpers";
```

(If the file already has a `writeFile` helper that creates parent dirs, reuse it; otherwise use `fsp.mkdir(path.dirname(p), {recursive:true})` then `fsp.writeFile`.)

- [ ] **Step 2: Run it, verify failure**

Run: `bunx vitest run tests/validators-docx.test.ts -t "legacy 2008 namespace"`
Expected: FAIL — two `w14:paraId=` matches and/or `parseXml` throws `Attribute w14:paraId redefined`.

- [ ] **Step 3: Apply the fix**

In `repairMissingParaIds`, change the two existence reads to add a qualified-name fallback:

```ts
const paraId = elem.getAttributeNS(W14_NAMESPACE, "paraId") || elem.getAttribute("w14:paraId");
const textId = elem.getAttributeNS(W14_NAMESPACE, "textId") || elem.getAttribute("w14:textId");
```

(`getAttribute("w14:paraId")` matches by serialized qualified name regardless of which URI `w14` is bound to, so an existing paraId is detected and the element falls into the no-op Case 4.)

- [ ] **Step 4: Run it, verify pass**

Run: `bunx vitest run tests/validators-docx.test.ts -t "legacy 2008 namespace"`
Expected: PASS.

- [ ] **Step 5: Full docx validator spec + type-check**

Run: `bunx vitest run tests/validators-docx.test.ts && bunx tsc --noEmit`
Expected: all PASS, tsc exit 0.

- [ ] **Step 6: Commit**

```bash
git add src/scripts/office/validators/docx.ts tests/validators-docx.test.ts
git commit -m "fix: don't duplicate w14:paraId when a doc binds w14 to the legacy 2008 namespace

repairMissingParaIds checked existence via getAttributeNS(W14_NAMESPACE=2010),
missing a w14:paraId bound to the older 2008/9/12 wordml URI and stamping a
duplicate qualified-name attribute (invalid XML, e.g. open-xml-sdk/mcdoc.docx).
Add a qualified-name fallback so an existing paraId/textId is detected.

Co-Authored-By: Claude Opus 4.7 (1M context) <noreply@anthropic.com>"
```

---

# Commit 2 — Harden `diff-docx` on unparseable parts

## Task 2: report per-input failures instead of throwing

**Files:**

- Modify: `scripts/diff-docx.ts` (`runDiffDocx`)
- Test: `tests/diff-docx.cli.test.ts`

- [ ] **Step 1: Write the failing test**

Add to `tests/diff-docx.cli.test.ts`:

```ts
it("reports a non-zero code instead of throwing when an input cannot be read", async () => {
    await withTempDir(async (dir) => {
        const good = path.join(dir, "good");
        await fs.mkdir(path.join(good, "word"), { recursive: true });
        await fs.writeFile(path.join(good, "word", "document.xml"), docXml(`<w:p><w:r><w:t>x</w:t></w:r></w:p>`), "utf-8");
        const bad = path.join(dir, "not-a-real.docx");
        await fs.writeFile(bad, "this is not a zip", "utf-8"); // unpack will fail

        const { code, markdown } = await runDiffDocx([good, bad]);
        expect(code).not.toBe(0);
        expect(markdown.toLowerCase()).toContain("could not read");
    });
});
```

(`docXml`, `fs`, `path`, `withTempDir`, `runDiffDocx` are already imported in this file.)

- [ ] **Step 2: Run it, verify failure**

Run: `bunx vitest run tests/diff-docx.cli.test.ts -t "cannot be read"`
Expected: FAIL — `runDiffDocx` throws (unhandled) instead of returning a code.

- [ ] **Step 3: Apply the fix**

In `scripts/diff-docx.ts`, replace the two `inventoryOf` calls in `runDiffDocx` with guarded collection. Find:

```ts
const diff = diffDocxInventories(await inventoryOf(a, profile), await inventoryOf(b, profile));
```

Replace with:

```ts
let invA: Awaited<ReturnType<typeof inventoryOf>>;
let invB: Awaited<ReturnType<typeof inventoryOf>>;
try {
    invA = await inventoryOf(a, profile);
} catch (err) {
    return { code: 1, markdown: `Error: could not read '${a}': ${err instanceof Error ? err.message : String(err)}` };
}
try {
    invB = await inventoryOf(b, profile);
} catch (err) {
    return { code: 1, markdown: `Error: could not read '${b}': ${err instanceof Error ? err.message : String(err)}` };
}
const diff = diffDocxInventories(invA, invB);
```

- [ ] **Step 4: Run it, verify pass + no regression**

Run: `bunx vitest run tests/diff-docx.cli.test.ts && bunx tsc --noEmit`
Expected: all PASS (the three prior CLI tests still pass), tsc exit 0.

- [ ] **Step 5: Commit**

```bash
git add scripts/diff-docx.ts tests/diff-docx.cli.test.ts
git commit -m "fix: diff-docx reports unreadable inputs instead of crashing

Guard each inventoryOf() in runDiffDocx; an unparseable/unzippable input now
returns a non-zero code with a 'could not read <input>' message rather than
throwing an unhandled error.

Co-Authored-By: Claude Opus 4.7 (1M context) <noreply@anthropic.com>"
```

---

# Commit 3 — `collectVisibleProjection` + `diffVisibleProjections`

Build the module incrementally (one visible dimension per sub-task, each TDD). All sub-tasks edit the same new file `src/scripts/office/validators/docx-visible-projection.ts` and test file `tests/docx-visible-projection.test.ts`; the single commit is at the end (Task 3h).

## Task 3a: module scaffold + paragraph text (final view) + block order

**Files:**

- Create: `src/scripts/office/validators/docx-visible-projection.ts`
- Create: `tests/docx-visible-projection.test.ts`

- [ ] **Step 1: Write the failing test**

Create `tests/docx-visible-projection.test.ts`:

```ts
import { promises as fs } from "node:fs";
import path from "node:path";
import { describe, expect, it } from "vitest";
import { withTempDir } from "../src/lib/run-cli";
import { collectVisibleProjection, diffVisibleProjections } from "../src/scripts/office/validators/docx-visible-projection";

const W_NS = `xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"`;
const W14 = `xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"`;
function doc(body: string): string {
    return `<?xml version="1.0"?><w:document ${W_NS} ${W14}><w:body>${body}</w:body></w:document>`;
}
async function write(dir: string, body: string): Promise<string> {
    await fs.mkdir(path.join(dir, "word"), { recursive: true });
    await fs.writeFile(path.join(dir, "word", "document.xml"), doc(body), "utf-8");
    return dir;
}

describe("collectVisibleProjection", () => {
    it("captures paragraph text in order, final tracked-changes view, tabs and breaks", async () => {
        await withTempDir(async (dir) => {
            await write(
                dir,
                `<w:p><w:r><w:t>One</w:t><w:tab/><w:t>Two</w:t></w:r></w:p>` +
                    `<w:p><w:ins><w:r><w:t>kept</w:t></w:r></w:ins><w:del><w:r><w:delText>gone</w:delText></w:r></w:del></w:p>` +
                    `<w:p><w:r><w:br/><w:t>after-break</w:t></w:r></w:p>`,
            );
            const proj = await collectVisibleProjection(dir, "strict");
            expect(proj.parts["word/document.xml"].blocks.map((b: any) => b.text)).toEqual(["One\tTwo", "kept", "¶after-break"]);
        });
    });
});
```

- [ ] **Step 2: Run it, verify failure**

Run: `bunx vitest run tests/docx-visible-projection.test.ts -t "paragraph text in order"`
Expected: FAIL — module/function does not exist.

- [ ] **Step 3: Implement the scaffold + paragraph text**

Create `src/scripts/office/validators/docx-visible-projection.ts`:

```ts
import { promises as fs } from "node:fs";
import path from "node:path";

import type { Profile } from "../../../lib/types";
import { parseXml } from "../../../lib/xml-helpers";

const WORD_NAMESPACES = [
    "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "http://purl.oclc.org/ooxml/wordprocessingml/main",
] as const;

export interface VisibleRun {
    text: string;
    fmt: string; // canonical sorted visible run-format signature ("" if none)
}
export interface VisibleBlock {
    kind: "paragraph" | "table";
    text: string; // paragraph: final visible text; table: "" (see grid)
    style?: string; // pStyle / tblStyle
    pPr?: string; // canonical visible paragraph-property signature
    runs?: VisibleRun[];
    grid?: { rows: number; cols: number; cells: VisibleBlock[][] }; // tables
    image?: { target: string; cx: number; cy: number };
    hyperlink?: { target: string; text: string };
}
export interface VisiblePart {
    blocks: VisibleBlock[];
}
export interface VisibleProjection {
    parts: Record<string, VisiblePart>; // keyed by part rel path
    pageGeometry: string[]; // ordered section-geometry signatures
    comments: string[]; // ordered comment texts
}

function isWordEl(node: Node | null): node is Element {
    return !!node && node.nodeType === 1 && WORD_NAMESPACES.includes((node as Element).namespaceURI as never);
}
function local(el: Element): string {
    return el.localName ?? el.nodeName.replace(/^.*:/, "");
}
function wAttr(el: Element, name: string): string | null {
    for (const ns of WORD_NAMESPACES) {
        const v = el.getAttributeNS(ns, name);
        if (v) return v;
    }
    return el.getAttribute(`w:${name}`) ?? el.getAttribute(name);
}
function wChild(parent: Element, name: string): Element | null {
    for (let c = parent.firstChild; c; c = c.nextSibling) {
        if (isWordEl(c) && local(c as Element) === name) return c as Element;
    }
    return null;
}
function wChildren(parent: Element, name: string): Element[] {
    const out: Element[] = [];
    for (let c = parent.firstChild; c; c = c.nextSibling) {
        if (isWordEl(c) && local(c as Element) === name) out.push(c as Element);
    }
    return out;
}
function hasWordAncestor(el: Element, name: string): boolean {
    for (let n: Node | null = el.parentNode; n; n = n.parentNode) {
        if (isWordEl(n) && local(n as Element) === name) return true;
    }
    return false;
}

// Final-view visible text of a paragraph: tabs -> \t, breaks -> ¶, deleted text
// (w:delText / inside w:del) dropped, hidden (w:vanish run) dropped.
function paragraphText(p: Element): string {
    let out = "";
    const walk = (node: Node): void => {
        for (let c = node.firstChild; c; c = c.nextSibling) {
            if (c.nodeType !== 1) continue;
            const el = c as Element;
            if (!isWordEl(el)) {
                walk(el);
                continue;
            }
            const ln = local(el);
            if (ln === "del" || ln === "delText" || ln === "delInstrText") continue; // not in final view
            if (ln === "r") {
                const rPr = wChild(el, "rPr");
                if (rPr && wChild(rPr, "vanish")) continue; // hidden run: not rendered
                walk(el);
                continue;
            }
            if (ln === "t") {
                out += el.textContent ?? "";
                continue;
            }
            if (ln === "tab") {
                out += "\t";
                continue;
            }
            if (ln === "br" || ln === "cr") {
                out += "¶";
                continue;
            }
            walk(el);
        }
    };
    walk(p);
    return out;
}

function projectParagraph(p: Element): VisibleBlock {
    return { kind: "paragraph", text: paragraphText(p) };
}

function projectBody(dom: Document): VisibleBlock[] {
    const blocks: VisibleBlock[] = [];
    let body: Element | null = null;
    for (const ns of WORD_NAMESPACES) {
        const list = dom.getElementsByTagNameNS(ns, "body");
        if (list.length > 0) {
            body = list.item(0);
            break;
        }
    }
    if (!body) return blocks;
    for (let c = body.firstChild; c; c = c.nextSibling) {
        if (!isWordEl(c)) continue;
        const el = c as Element;
        const ln = local(el);
        if (ln === "p") blocks.push(projectParagraph(el));
        // tables added in Task 3c
    }
    return blocks;
}

export async function collectVisibleProjection(unpackedDir: string, _profile: Profile = "lenient"): Promise<VisibleProjection> {
    const proj: VisibleProjection = { parts: {}, pageGeometry: [], comments: [] };
    const docPath = path.join(unpackedDir, "word", "document.xml");
    try {
        const dom = parseXml(await fs.readFile(docPath, "utf-8"));
        proj.parts["word/document.xml"] = { blocks: projectBody(dom) };
    } catch {
        // unreadable document.xml: leave parts empty (caller treats as a defect-worthy difference)
    }
    return proj;
}

export interface VisibleProjectionDelta {
    path: string;
    detail: string;
}
export function diffVisibleProjections(_before: VisibleProjection, _after: VisibleProjection): VisibleProjectionDelta[] {
    return []; // implemented in Task 3h
}
```

- [ ] **Step 4: Run it, verify pass**

Run: `bunx vitest run tests/docx-visible-projection.test.ts -t "paragraph text in order"`
Expected: PASS.

## Task 3b: visible run + paragraph formatting

**Files:** modify the new module + test.

- [ ] **Step 1: Write the failing test**

```ts
it("captures visible run formatting and paragraph style, excludes hidden runs", async () => {
    await withTempDir(async (dir) => {
        await write(
            dir,
            `<w:p><w:pPr><w:pStyle w:val="Heading1"/><w:jc w:val="center"/></w:pPr>` +
                `<w:r><w:rPr><w:b/><w:color w:val="FF0000"/></w:rPr><w:t>Bold red</w:t></w:r></w:p>` +
                `<w:p><w:r><w:rPr><w:vanish/></w:rPr><w:t>secret</w:t></w:r><w:r><w:t>shown</w:t></w:r></w:p>`,
        );
        const proj = await collectVisibleProjection(dir, "strict");
        const blocks = proj.parts["word/document.xml"].blocks as any[];
        expect(blocks[0].style).toBe("Heading1");
        expect(blocks[0].pPr).toContain("jc=center");
        expect(blocks[0].runs[0].fmt).toContain("b");
        expect(blocks[0].runs[0].fmt).toContain("color=FF0000");
        expect(blocks[1].text).toBe("shown"); // hidden run excluded
    });
});
```

- [ ] **Step 2: Run it, verify failure**

Run: `bunx vitest run tests/docx-visible-projection.test.ts -t "visible run formatting"`
Expected: FAIL — `style`/`pPr`/`runs` undefined.

- [ ] **Step 3: Implement formatting**

Add helpers and extend `projectParagraph` in the module:

```ts
function runFormat(rPr: Element | null): string {
    if (!rPr) return "";
    const sig: string[] = [];
    for (const toggle of ["b", "i", "strike", "dstrike", "caps", "smallCaps"]) {
        if (wChild(rPr, toggle)) sig.push(toggle);
    }
    for (const [el, key] of [
        ["u", "u"],
        ["color", "color"],
        ["highlight", "highlight"],
        ["sz", "sz"],
        ["vertAlign", "vertAlign"],
        ["rStyle", "rStyle"],
    ] as const) {
        const child = wChild(rPr, el);
        if (child) {
            const cv = wAttr(child, "val");
            if (cv !== null) sig.push(`${key}=${cv}`);
        }
    }
    const fonts = wChild(rPr, "rFonts");
    if (fonts) {
        const a = wAttr(fonts, "ascii");
        if (a) sig.push(`font=${a}`);
    }
    return sig.sort().join(",");
}

function paragraphProps(p: Element): { style?: string; pPr: string } {
    const pPr = wChild(p, "pPr");
    if (!pPr) return { pPr: "" };
    const pStyleEl = wChild(pPr, "pStyle");
    const style = pStyleEl ? (wAttr(pStyleEl, "val") ?? undefined) : undefined;
    const sig: string[] = [];
    for (const [el, key] of [
        ["jc", "jc"],
        ["ind", "ind"],
        ["spacing", "spacing"],
    ] as const) {
        const child = wChild(pPr, el);
        if (child) {
            const v = wAttr(child, "val") ?? wAttr(child, "left") ?? wAttr(child, "line") ?? "y";
            sig.push(`${key}=${v}`);
        }
    }
    const numPr = wChild(pPr, "numPr");
    if (numPr) {
        const numId = wAttr(wChild(numPr, "numId") ?? numPr, "val");
        const ilvl = wAttr(wChild(numPr, "ilvl") ?? numPr, "val");
        sig.push(`num=${numId ?? "?"}:${ilvl ?? "0"}`);
    }
    return { style, pPr: sig.sort().join(",") };
}

function projectRuns(p: Element): VisibleRun[] {
    const runs: VisibleRun[] = [];
    const collect = (node: Node, deleted: boolean): void => {
        for (let c = node.firstChild; c; c = c.nextSibling) {
            if (!isWordEl(c)) continue;
            const el = c as Element;
            const ln = local(el);
            if (ln === "del") {
                continue;
            }
            if (ln === "ins") {
                collect(el, deleted);
                continue;
            }
            if (ln === "r") {
                const rPr = wChild(el, "rPr");
                if (rPr && wChild(rPr, "vanish")) continue;
                const text = paragraphText(el); // reuse: handles t/tab/br within run
                if (text.length > 0) runs.push({ text, fmt: runFormat(rPr) });
            }
        }
    };
    collect(p, false);
    return runs;
}
```

Replace `projectParagraph` with:

```ts
function projectParagraph(p: Element): VisibleBlock {
    const { style, pPr } = paragraphProps(p);
    return { kind: "paragraph", text: paragraphText(p), style, pPr, runs: projectRuns(p) };
}
```

- [ ] **Step 4: Run it, verify pass**

Run: `bunx vitest run tests/docx-visible-projection.test.ts -t "visible run formatting"`
Expected: PASS.

## Task 3c: tables (grid + cell content)

- [ ] **Step 1: Write the failing test**

```ts
it("projects table grid and cell content", async () => {
    await withTempDir(async (dir) => {
        await write(
            dir,
            `<w:tbl><w:tblPr><w:tblStyle w:val="TableGrid"/></w:tblPr><w:tblGrid><w:gridCol/><w:gridCol/></w:tblGrid>` +
                `<w:tr><w:tc><w:p><w:r><w:t>A1</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>B1</w:t></w:r></w:p></w:tc></w:tr></w:tbl>`,
        );
        const proj = await collectVisibleProjection(dir, "strict");
        const tbl = (proj.parts["word/document.xml"].blocks as any[])[0];
        expect(tbl.kind).toBe("table");
        expect(tbl.style).toBe("TableGrid");
        expect(tbl.grid.rows).toBe(1);
        expect(tbl.grid.cols).toBe(2);
        expect(tbl.grid.cells[0][0].text).toBe("A1");
        expect(tbl.grid.cells[0][1].text).toBe("B1");
    });
});
```

- [ ] **Step 2: Run it, verify failure**

Run: `bunx vitest run tests/docx-visible-projection.test.ts -t "table grid and cell"`
Expected: FAIL — tables not projected (block 0 is undefined / not a table).

- [ ] **Step 3: Implement table projection**

Add to the module and wire into `projectBody`:

```ts
function projectTable(tbl: Element): VisibleBlock {
    const tblPr = wChild(tbl, "tblPr");
    const style = tblPr ? (wChild(tblPr, "tblStyle") ? (wAttr(wChild(tblPr, "tblStyle")!, "val") ?? undefined) : undefined) : undefined;
    const rows = wChildren(tbl, "tr");
    const grid = wChild(tbl, "tblGrid");
    const cols = grid ? wChildren(grid, "gridCol").length : rows[0] ? wChildren(rows[0], "tc").length : 0;
    const cells: VisibleBlock[][] = rows.map((tr) =>
        wChildren(tr, "tc").map((tc) => {
            const inner: VisibleBlock[] = [];
            for (let c = tc.firstChild; c; c = c.nextSibling) {
                if (!isWordEl(c)) continue;
                const el = c as Element;
                if (local(el) === "p") inner.push(projectParagraph(el));
                else if (local(el) === "tbl") inner.push(projectTable(el));
            }
            // flatten cell text for convenient assertions
            return {
                kind: "paragraph",
                text: inner.map((b) => b.text).join("\n"),
                runs: inner.flatMap((b) => b.runs ?? []),
            } as VisibleBlock;
        }),
    );
    return { kind: "table", text: "", style, grid: { rows: rows.length, cols, cells } };
}
```

In `projectBody`, inside the loop, after the `p` branch add:

```ts
if (ln === "p") blocks.push(projectParagraph(el));
else if (ln === "tbl") blocks.push(projectTable(el));
```

- [ ] **Step 4: Run it, verify pass**

Run: `bunx vitest run tests/docx-visible-projection.test.ts -t "table grid and cell"`
Expected: PASS.

## Task 3d: images + hyperlinks resolved via relationships

- [ ] **Step 1: Write the failing test**

```ts
it("resolves image and hyperlink relationship targets (not the rId string)", async () => {
    await withTempDir(async (dir) => {
        await fs.mkdir(path.join(dir, "word", "_rels"), { recursive: true });
        await fs.writeFile(
            path.join(dir, "word", "_rels", "document.xml.rels"),
            `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
                `<Relationship Id="rId9" Type="img" Target="media/image1.png"/>` +
                `<Relationship Id="rId5" Type="hyperlink" Target="https://example.com" TargetMode="External"/></Relationships>`,
            "utf-8",
        );
        await write(
            dir,
            `<w:p><w:hyperlink r:id="rId5" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:r><w:t>link</w:t></w:r></w:hyperlink></w:p>` +
                `<w:p><w:r><w:drawing><wp:inline xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing">` +
                `<wp:extent cx="1905000" cy="1270000"/>` +
                `<a:blip xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:embed="rId9"/>` +
                `</wp:inline></w:drawing></w:r></w:p>`,
        );
        const proj = await collectVisibleProjection(dir, "strict");
        const blocks = proj.parts["word/document.xml"].blocks as any[];
        expect(blocks[0].hyperlink).toEqual({ target: "https://example.com", text: "link" });
        expect(blocks[1].image).toEqual({ target: "media/image1.png", cx: 1905000, cy: 1270000 });
    });
});
```

- [ ] **Step 2: Run it, verify failure**

Run: `bunx vitest run tests/docx-visible-projection.test.ts -t "relationship targets"`
Expected: FAIL — `hyperlink`/`image` undefined.

- [ ] **Step 3: Implement rels resolution + image/hyperlink projection**

Add a rels loader and extend collection. The `r:` namespace is `http://schemas.openxmlformats.org/officeDocument/2006/relationships`; drawing namespaces are `wp` (`…/drawingml/2006/wordprocessingDrawing`) and `a` (`…/drawingml/2006/main`).

```ts
const REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships";
const R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const WP_NS = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
const A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main";

async function loadRels(unpackedDir: string, partRelPath: string): Promise<Map<string, string>> {
    // partRelPath e.g. "word/document.xml" -> "word/_rels/document.xml.rels"
    const dir = path.dirname(partRelPath);
    const base = path.basename(partRelPath);
    const relsPath = path.join(unpackedDir, dir, "_rels", `${base}.rels`);
    const map = new Map<string, string>();
    try {
        const dom = parseXml(await fs.readFile(relsPath, "utf-8"));
        const rels = dom.getElementsByTagNameNS(REL_NS, "Relationship");
        for (let i = 0; i < rels.length; i += 1) {
            const r = rels.item(i)!;
            const id = r.getAttribute("Id");
            const target = r.getAttribute("Target");
            if (id && target) map.set(id, target);
        }
    } catch {
        // no rels
    }
    return map;
}

function firstDescendantNS(el: Element, ns: string, name: string): Element | null {
    const list = el.getElementsByTagNameNS(ns, name);
    return list.length > 0 ? list.item(0) : null;
}
```

Then thread `rels` into block projection. Change `projectBody(dom)` to `projectBody(dom, rels)` and pass `rels` down to `projectParagraph`. In `projectParagraph`, before returning, detect hyperlink/image:

```ts
function projectParagraph(p: Element, rels: Map<string, string>): VisibleBlock {
    const { style, pPr } = paragraphProps(p);
    const block: VisibleBlock = { kind: "paragraph", text: paragraphText(p), style, pPr, runs: projectRuns(p) };
    // hyperlink (w:hyperlink with r:id)
    for (let c = p.firstChild; c; c = c.nextSibling) {
        if (isWordEl(c) && local(c as Element) === "hyperlink") {
            const id = (c as Element).getAttributeNS(R_NS, "id");
            block.hyperlink = { target: (id && rels.get(id)) || "", text: paragraphText(c as Element) };
        }
    }
    // image (wp:extent + a:blip r:embed)
    const extent = firstDescendantNS(p, WP_NS, "extent");
    if (extent) {
        const blip = firstDescendantNS(p, A_NS, "blip");
        const embed = blip?.getAttributeNS(R_NS, "embed") ?? "";
        block.image = {
            target: (embed && rels.get(embed)) || "",
            cx: Number.parseInt(extent.getAttribute("cx") ?? "0", 10) || 0,
            cy: Number.parseInt(extent.getAttribute("cy") ?? "0", 10) || 0,
        };
    }
    return block;
}
```

Update all `projectParagraph(...)` / `projectTable(...)` call sites to pass `rels` (table cells call `projectParagraph(el, rels)`; `projectTable(tbl, rels)`), and in `collectVisibleProjection` load rels once:

```ts
const rels = await loadRels(unpackedDir, "word/document.xml");
proj.parts["word/document.xml"] = { blocks: projectBody(dom, rels) };
```

- [ ] **Step 4: Run it, verify pass**

Run: `bunx vitest run tests/docx-visible-projection.test.ts -t "relationship targets"`
Expected: PASS.

## Task 3e: page geometry + comments

- [ ] **Step 1: Write the failing test**

```ts
it("captures page geometry and comment text", async () => {
    await withTempDir(async (dir) => {
        await write(
            dir,
            `<w:p><w:r><w:t>body</w:t></w:r></w:p>` +
                `<w:sectPr><w:pgSz w:w="12240" w:h="15840" w:orient="portrait"/><w:cols w:num="2"/></w:sectPr>`,
        );
        await fs.writeFile(
            path.join(dir, "word", "comments.xml"),
            `<?xml version="1.0"?><w:comments ${W_NS}><w:comment w:id="1"><w:p><w:r><w:t>nice</w:t></w:r></w:p></w:comment></w:comments>`,
            "utf-8",
        );
        const proj = await collectVisibleProjection(dir, "strict");
        expect(proj.pageGeometry.join("|")).toContain("portrait 12240x15840");
        expect(proj.pageGeometry.join("|")).toContain("cols=2");
        expect(proj.comments).toEqual(["nice"]);
    });
});
```

- [ ] **Step 2: Run it, verify failure**

Run: `bunx vitest run tests/docx-visible-projection.test.ts -t "page geometry and comment"`
Expected: FAIL — `pageGeometry`/`comments` empty.

- [ ] **Step 3: Implement page geometry + comments**

Add to `collectVisibleProjection`, after building `parts["word/document.xml"]`:

```ts
// Page geometry (all sectPr in document order).
for (const ns of WORD_NAMESPACES) {
    const sects = dom.getElementsByTagNameNS(ns, "sectPr");
    for (let i = 0; i < sects.length; i += 1) {
        const sect = sects.item(i)!;
        const sig: string[] = [];
        const pgSz = wChild(sect, "pgSz");
        if (pgSz) sig.push(`${wAttr(pgSz, "orient") ?? "portrait"} ${wAttr(pgSz, "w") ?? "?"}x${wAttr(pgSz, "h") ?? "?"}`);
        const pgMar = wChild(sect, "pgMar");
        if (pgMar)
            sig.push(
                `mar ${wAttr(pgMar, "top") ?? "?"},${wAttr(pgMar, "right") ?? "?"},${wAttr(pgMar, "bottom") ?? "?"},${wAttr(pgMar, "left") ?? "?"}`,
            );
        const cols = wChild(sect, "cols");
        sig.push(`cols=${cols ? (wAttr(cols, "num") ?? "1") : "1"}`);
        proj.pageGeometry.push(sig.join(" "));
    }
}
```

And a comments collector (after the try for document.xml):

```ts
try {
    const cdom = parseXml(await fs.readFile(path.join(unpackedDir, "word", "comments.xml"), "utf-8"));
    for (const ns of WORD_NAMESPACES) {
        const cs = cdom.getElementsByTagNameNS(ns, "comment");
        for (let i = 0; i < cs.length; i += 1) {
            const paras = wChildren(cs.item(i)!, "p");
            proj.comments.push(paras.map((p) => paragraphText(p)).join("\n"));
        }
    }
} catch {
    // no comments part
}
```

- [ ] **Step 4: Run it, verify pass**

Run: `bunx vitest run tests/docx-visible-projection.test.ts -t "page geometry and comment"`
Expected: PASS.

## Task 3f: headers/footers content

- [ ] **Step 1: Write the failing test**

```ts
it("projects header and footer part content", async () => {
    await withTempDir(async (dir) => {
        await write(dir, `<w:p><w:r><w:t>body</w:t></w:r></w:p>`);
        await fs.writeFile(
            path.join(dir, "word", "header1.xml"),
            `<?xml version="1.0"?><w:hdr ${W_NS}><w:p><w:r><w:t>HEAD</w:t></w:r></w:p></w:hdr>`,
            "utf-8",
        );
        const proj = await collectVisibleProjection(dir, "strict");
        expect(proj.parts["word/header1.xml"].blocks.map((b: any) => b.text)).toEqual(["HEAD"]);
    });
});
```

- [ ] **Step 2: Run it, verify failure**

Run: `bunx vitest run tests/docx-visible-projection.test.ts -t "header and footer part"`
Expected: FAIL — `parts["word/header1.xml"]` undefined.

- [ ] **Step 3: Implement header/footer projection**

In `collectVisibleProjection`, after the document.xml block, scan the `word/` dir for header*/footer* parts and project their bodies (they contain `w:p`/`w:tbl` directly under `w:hdr`/`w:ftr`):

```ts
try {
    const wordDir = path.join(unpackedDir, "word");
    for (const name of await fs.readdir(wordDir)) {
        if (!/^(header|footer)\d*\.xml$/i.test(name)) continue;
        const rel = `word/${name}`;
        try {
            const hdom = parseXml(await fs.readFile(path.join(wordDir, name), "utf-8"));
            const rels = await loadRels(unpackedDir, rel);
            // root is w:hdr/w:ftr — reuse projectBody by treating root's children
            const blocks: VisibleBlock[] = [];
            const root = hdom.documentElement;
            for (let c = root.firstChild; c; c = c.nextSibling) {
                if (!isWordEl(c)) continue;
                const el = c as Element;
                if (local(el) === "p") blocks.push(projectParagraph(el, rels));
                else if (local(el) === "tbl") blocks.push(projectTable(el, rels));
            }
            proj.parts[rel] = { blocks };
        } catch {
            // skip unreadable header/footer
        }
    }
} catch {
    // no word dir
}
```

- [ ] **Step 4: Run it, verify pass**

Run: `bunx vitest run tests/docx-visible-projection.test.ts -t "header and footer part"`
Expected: PASS.

> **Numbering note:** `numId`/`ilvl` are already captured in each paragraph's `pPr` signature (Task 3b), so a numbering _reference_ change is visible. Resolving `numbering.xml` to concrete number _formats_ is deferred as a documented projection limitation (a format-only change with an unchanged `numId` is rare and low-visibility); if `prove-benign` later surfaces a numbering-format false-negative, extend here.

## Task 3g: deterministic serialization for comparison

- [ ] **Step 1: Write the failing test**

```ts
it("normalizes insignificant whitespace and ignores plumbing so a plumbing-only change is identical", async () => {
    await withTempDir(async (dir) => {
        const a = path.join(dir, "a");
        const b = path.join(dir, "b");
        await write(a, `<w:p><w:r><w:t>Hello</w:t></w:r></w:p>`);
        // b differs ONLY by added w14:paraId + reordered/extra xmlns + indentation
        await fs.mkdir(path.join(b, "word"), { recursive: true });
        await fs.writeFile(
            path.join(b, "word", "document.xml"),
            `<?xml version="1.0"?>\n<w:document ${W_NS} ${W14}>\n  <w:body>\n` +
                `    <w:p w14:paraId="11111111" w14:textId="22222222"><w:r><w:t>Hello</w:t></w:r></w:p>\n` +
                `  </w:body>\n</w:document>`,
            "utf-8",
        );
        const pa = await collectVisibleProjection(a, "strict");
        const pb = await collectVisibleProjection(b, "strict");
        expect(diffVisibleProjections(pa, pb)).toEqual([]);
    });
});
```

- [ ] **Step 2: Run it, verify failure**

Run: `bunx vitest run tests/docx-visible-projection.test.ts -t "plumbing-only change is identical"`
Expected: PASS already if `diffVisibleProjections` returns `[]` unconditionally — so FIRST make it real (Task 3h) and re-run. To keep this red→green honest, implement Task 3h before this assertion is meaningful; treat 3g's test as part of 3h's suite.

## Task 3h: `diffVisibleProjections` + visible-change detection + commit

- [ ] **Step 1: Write the failing test (visible change MUST be detected)**

```ts
it("detects a visible change (dropped bold) as a non-empty delta", async () => {
    await withTempDir(async (dir) => {
        const a = path.join(dir, "a");
        const b = path.join(dir, "b");
        await write(a, `<w:p><w:r><w:rPr><w:b/></w:rPr><w:t>Hi</w:t></w:r></w:p>`);
        await write(b, `<w:p><w:r><w:t>Hi</w:t></w:r></w:p>`); // bold removed
        const delta = diffVisibleProjections(await collectVisibleProjection(a, "strict"), await collectVisibleProjection(b, "strict"));
        expect(delta.length).toBeGreaterThan(0);
        expect(JSON.stringify(delta)).toContain("word/document.xml");
    });
});
```

- [ ] **Step 2: Run it, verify failure**

Run: `bunx vitest run tests/docx-visible-projection.test.ts -t "dropped bold"`
Expected: FAIL — `diffVisibleProjections` returns `[]` (stub).

- [ ] **Step 3: Implement `diffVisibleProjections` by canonical deep compare**

Replace the stub:

```ts
function canonical(value: unknown): string {
    return JSON.stringify(value, (_k, v) => {
        if (v && typeof v === "object" && !Array.isArray(v)) {
            return Object.keys(v as Record<string, unknown>)
                .sort()
                .reduce<Record<string, unknown>>((acc, k) => {
                    acc[k] = (v as Record<string, unknown>)[k];
                    return acc;
                }, {});
        }
        return v;
    });
}

export function diffVisibleProjections(before: VisibleProjection, after: VisibleProjection): VisibleProjectionDelta[] {
    const deltas: VisibleProjectionDelta[] = [];
    const partNames = new Set([...Object.keys(before.parts), ...Object.keys(after.parts)]);
    for (const name of [...partNames].sort()) {
        const a = canonical(before.parts[name]?.blocks ?? null);
        const b = canonical(after.parts[name]?.blocks ?? null);
        if (a !== b) deltas.push({ path: name, detail: `visible content changed (before≠after)` });
    }
    if (canonical(before.pageGeometry) !== canonical(after.pageGeometry)) {
        deltas.push({ path: "<page-geometry>", detail: `${before.pageGeometry.join(" | ")} → ${after.pageGeometry.join(" | ")}` });
    }
    if (canonical(before.comments) !== canonical(after.comments)) {
        deltas.push({ path: "<comments>", detail: `comment text changed` });
    }
    return deltas;
}
```

- [ ] **Step 4: Run the full projection spec**

Run: `bunx vitest run tests/docx-visible-projection.test.ts && bunx tsc --noEmit`
Expected: all PASS (including the 3g plumbing-only test now meaningful), tsc exit 0.

- [ ] **Step 5: Commit (Commit 3)**

```bash
git add src/scripts/office/validators/docx-visible-projection.ts tests/docx-visible-projection.test.ts
git commit -m "feat: collectVisibleProjection + diffVisibleProjections (ordinary-user visible model)

Order-preserving projection of reader-perceptible content (final-view text,
visible run/paragraph formatting, tables, resolved image/hyperlink targets,
page geometry, headers/footers, comments), normalizing away plumbing. Two
projections are equal iff no user-visible change occurred.

Co-Authored-By: Claude Opus 4.7 (1M context) <noreply@anthropic.com>"
```

---

# Commit 4 — `prove-benign.ts` + proof run

## Task 4: prove projection invariance under repair across fixtures

**Files:**

- Create: `scripts/prove-benign.ts`
- Output: `.drift-run/BENIGN_PROOF.md` (committed as the deliverable)

- [ ] **Step 1: Write the script**

Create `scripts/prove-benign.ts` (reuses the extract/repack pattern from `scripts/fixtures-drift-report.ts` and the new projection). It repairs a copy of each `.docx` fixture and asserts the visible projection is invariant:

```ts
/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */
import { promises as fs } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";
import JSZip from "jszip";
import { withTempDir } from "../src/lib/run-cli";
import type { Profile } from "../src/lib/types";
import { DOCXSchemaValidator } from "../src/scripts/office/validators/docx";
import { collectVisibleProjection, diffVisibleProjections } from "../src/scripts/office/validators/docx-visible-projection";

const HERE = path.dirname(fileURLToPath(import.meta.url));
const FIXTURES_ROOT = path.resolve(HERE, "..", "tests", "fixtures");

async function walk(dir: string, out: string[]): Promise<void> {
    for (const e of await fs.readdir(dir, { withFileTypes: true })) {
        const full = path.join(dir, e.name);
        if (e.isDirectory()) await walk(full, out);
        else if (e.isFile() && full.toLowerCase().endsWith(".docx")) out.push(full);
    }
}
async function extract(buf: Buffer, outDir: string): Promise<void> {
    const zip = await JSZip.loadAsync(buf);
    const entries: Array<{ name: string; file: JSZip.JSZipObject }> = [];
    zip.forEach((p, f) => entries.push({ name: p, file: f }));
    for (const { name, file } of entries) {
        const resolved = path.resolve(outDir, name);
        if (!resolved.startsWith(`${outDir}${path.sep}`) && resolved !== outDir) throw new Error(`zip-slip: ${name}`);
        if (file.dir) await fs.mkdir(resolved, { recursive: true });
        else {
            await fs.mkdir(path.dirname(resolved), { recursive: true });
            await fs.writeFile(resolved, await file.async("nodebuffer"));
        }
    }
}

interface Row {
    rel: string;
    verdict: "proven-benign" | "defect" | "errored";
    detail: string;
}

async function main(): Promise<void> {
    const profile: Profile = "strict";
    const out = path.resolve(HERE, "..", ".drift-run");
    const files: string[] = [];
    await walk(FIXTURES_ROOT, files);
    files.sort((a, b) => a.localeCompare(b));
    const rows: Row[] = [];
    for (let i = 0; i < files.length; i += 1) {
        const rel = path.relative(FIXTURES_ROOT, files[i]!).split(path.sep).join("/");
        try {
            const buf = await fs.readFile(files[i]!);
            const row = await withTempDir(async (tmp) => {
                const before = path.join(tmp, "before");
                const after = path.join(tmp, "after");
                await extract(buf, before);
                await extract(buf, after);
                const beforeProj = await collectVisibleProjection(before, profile);
                await new DOCXSchemaValidator({ unpackedDir: after, profile }).repair();
                const afterProj = await collectVisibleProjection(after, profile);
                const delta = diffVisibleProjections(beforeProj, afterProj);
                return delta.length === 0
                    ? ({ rel, verdict: "proven-benign", detail: "visible projection identical" } as Row)
                    : ({ rel, verdict: "defect", detail: delta.map((d) => `${d.path}: ${d.detail}`).join("; ") } as Row);
            });
            rows.push(row);
        } catch (err) {
            rows.push({ rel, verdict: "errored", detail: err instanceof Error ? err.message : String(err) });
        }
        if ((i + 1) % 50 === 0 || i + 1 === files.length) process.stdout.write(`  ${i + 1}/${files.length}\n`);
    }
    const proven = rows.filter((r) => r.verdict === "proven-benign");
    const defects = rows.filter((r) => r.verdict === "defect");
    const errored = rows.filter((r) => r.verdict === "errored");
    const md = [
        `# Benign-drift proof (profile: ${profile})`,
        "",
        `Invariant: **repair must not change a document's visible projection** (the ordinary-Word-user model).`,
        "",
        `- Proven benign (visible projection identical before/after repair): **${proven.length}**`,
        `- **Repair defects (visible change): ${defects.length}**`,
        `- Could not process (encrypted/corrupt): ${errored.length}`,
        "",
        ...(defects.length
            ? [
                  "## Repair defects — visible change detected",
                  "",
                  "| Fixture | visible delta |",
                  "|---|---|",
                  ...defects.map((r) => `| \`${r.rel}\` | ${r.detail} |`),
                  "",
              ]
            : ["_No repair defects: every processable fixture's visible projection was invariant under repair._", ""]),
        ...(errored.length ? ["## Could not process", "", ...errored.map((r) => `- \`${r.rel}\`: ${r.detail}`), ""] : []),
    ].join("\n");
    await fs.mkdir(out, { recursive: true });
    await fs.writeFile(path.join(out, "BENIGN_PROOF.md"), md, "utf-8");
    process.stdout.write(
        `\nproven=${proven.length} defects=${defects.length} errored=${errored.length}\nReport: ${path.join(out, "BENIGN_PROOF.md")}\n`,
    );
}
main().catch((e: unknown) => {
    process.stderr.write(`${e instanceof Error ? (e.stack ?? e.message) : String(e)}\n`);
    process.exit(1);
});
```

- [ ] **Step 2: Smoke-test on a few fixtures**

Run: `bunx tsx scripts/prove-benign.ts 2>&1 | tail -5` is the full run; for a smoke test first, temporarily verify imports compile: `bunx tsc --noEmit`
Expected: tsc exit 0.

- [ ] **Step 3: Full run**

Run: `bunx tsx scripts/prove-benign.ts`
Expected: completes; prints `proven=… defects=… errored=…`. The `mcdoc.docx` defect from Commit 1 should now be **proven-benign** (its projection is invariant after the paraId fix). Any remaining `defects` are genuine repair bugs to triage (follow-up, one TDD fix each).

- [ ] **Step 4: Review defects**

Open `.drift-run/BENIGN_PROOF.md`. For each defect, inspect the visible delta. If a delta is actually a _projection-model gap_ (a plumbing change the projection wrongly treats as visible), fix `docx-visible-projection.ts` (add a focused test first) and re-run. If a delta is a real user-visible repair change, record it as a follow-up repair defect.

- [ ] **Step 5: Commit the proof**

Because `.drift-run/` is gitignored, copy the proof into the docs tree to commit it as the deliverable:

```bash
mkdir -p docs/superpowers/reports
cp .drift-run/BENIGN_PROOF.md docs/superpowers/reports/2026-05-26-benign-drift-proof.md
git add scripts/prove-benign.ts docs/superpowers/reports/2026-05-26-benign-drift-proof.md
git commit -m "feat: prove-benign survey — visible-projection invariance under repair

Repairs a copy of every .docx fixture (strict) and asserts the visible
projection is invariant. Proof report committed. Proven-benign = no user-visible
change; any visible delta is flagged as a repair defect for follow-up.

Co-Authored-By: Claude Opus 4.7 (1M context) <noreply@anthropic.com>"
```

- [ ] **Step 6: Push the branch**

```bash
git push origin feat/inventory-fingerprint
```

---

## Final verification

- [ ] `bunx tsc --noEmit` → exit 0
- [ ] `bunx vitest run tests/validators-docx.test.ts tests/diff-docx.cli.test.ts tests/docx-visible-projection.test.ts` → all PASS
- [ ] `.drift-run/BENIGN_PROOF.md` shows `mcdoc.docx` proven-benign (paraId fix holds) and lists any remaining defects
- [ ] Do NOT run `bun run fmt:fix` (repo-wide). Restrict formatting to changed files; `git checkout -- <file>` any unrelated reformatting.

## Notes / limitations

- Projection is the _final_ tracked-changes view; hidden (`vanish`) text excluded; `xml:space="preserve"` significant; comments included — per approved spec.
- Numbering-format resolution from `numbering.xml` is deferred (numId/ilvl reference is captured); extend only if `prove-benign` surfaces a numbering false-negative.
- Render-diff validation against real Word/LibreOffice (`SOFFICE_AVAILABLE`) is deferred future work to validate the projection model.
