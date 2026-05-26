# DOCX Semantic Fingerprint & Symmetric Inventory Diff — Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Expand the DOCX semantic inventory with four new collector families (atomic marks, table shape, section geometry, image shape) and add a symmetric A-vs-B fingerprint diff with severity tiers, formatters, and a `diff-docx` CLI.

**Architecture:** Aggregate-histogram model. Collectors add counters to the existing `DocxSemanticInventory` (a `Map` keyed by `path\0category\0label\0unit`). A shared `severityClassFor(category)` maps each counter to a class; both the new symmetric `diffDocxInventories` and the existing decrease-only repair gate tier their output through it. Families 3+4 collect only under the `strict` profile.

**Tech Stack:** TypeScript (ESM), `@xmldom/xmldom` via `src/lib/xml-helpers.ts`, `commander` for the CLI, `vitest` for tests, `bun` as package manager.

**Spec:** `docs/superpowers/specs/2026-05-26-docx-inventory-fingerprint-design.md`

**Branch:** `feat/inventory-fingerprint` (already created, stacked on `feat/second-pass-diff-analysis`).

---

## File structure

| File | Responsibility |
|------|----------------|
| `src/scripts/office/validators/docx-diagnostics.ts` | **(modify)** Add `directWordChildren` helper, four collector families, `severityClassFor`, a `profile` arg on `collectDocxSemanticInventory`, and the gate code split in `compareDocxSemanticInventories`. |
| `src/scripts/office/validators/docx-inventory-diff.ts` | **(new)** `DocxInventoryDelta`/`DocxInventoryDiff` types, `diffDocxInventories`, `severityFor`, `inventoryDiffToIssues`, `formatInventoryDiffMarkdown`. |
| `src/scripts/office/validate.ts` | **(modify)** Thread the active `profile` into the two `collectDocxSemanticInventory` calls in the repair path. |
| `scripts/diff-docx.ts` | **(new)** CLI: `diff-docx <a> <b> [--profile …]`. |
| `tests/docx-diagnostics.test.ts` | **(modify)** Tests for new collectors + `severityClassFor` + gate split. |
| `tests/docx-inventory-diff.test.ts` | **(new)** Tests for diff, severity, formatters. |
| `tests/diff-docx.cli.test.ts` | **(new)** CLI smoke tests. |

## Shared reference — helpers that already exist in `docx-diagnostics.ts`

Use these; do **not** reimplement them:
- `addCounter(inventory, pathValue, category, label, unit, count)` — skips `count <= 0`, sums duplicates.
- `directWordChild(parent, local): Element | null` — first direct child in the `w:`/strict namespace with that local name.
- `wordChildAttr(parent, local, attr): string | null` — attribute of a direct word child.
- `enabledChild(parent, local): boolean` — toggle-element semantics.
- `isWordElement(elem)`, `localName(elem)`, `wordNamespace(elem)`.
- `WORD_NAMESPACES` (Set), `WP_NAMESPACE`.
- Test harness in `tests/docx-diagnostics.test.ts`: `withTempDir`, `writeXml(filePath, content)`, `doc(body)` (wraps body in `<w:document><w:body>…`), `W_NS`.

---

# Commit A — Collector families

## Task A1: `directWordChildren` helper

**Files:**
- Modify: `src/scripts/office/validators/docx-diagnostics.ts` (add near `directWordChild`)

- [ ] **Step 1: Add the helper**

Add immediately after the existing `directWordChild` function:

```ts
function directWordChildren(parent: Element, local: string): Element[] {
    const out: Element[] = [];
    for (let child = parent.firstChild; child; child = child.nextSibling) {
        if (child.nodeType !== 1) continue;
        const elem = child as Element;
        if (isWordElement(elem) && localName(elem) === local) out.push(elem);
    }
    return out;
}
```

- [ ] **Step 2: Type-check**

Run: `bunx tsc --noEmit`
Expected: exit 0 (no errors). The function is unused for now; that is fine — it is `function`-scoped and TS does not error on unused module-internal functions under this config. If a lint/unused error appears, proceed to Task A2 which uses it.

---

## Task A2: Family 1 — in-run atomic marks collector

**Files:**
- Modify: `src/scripts/office/validators/docx-diagnostics.ts`
- Test: `tests/docx-diagnostics.test.ts`

- [ ] **Step 1: Write the failing test**

Add inside the top-level `describe("docx diagnostics", …)` block in `tests/docx-diagnostics.test.ts`:

```ts
it("counts in-run atomic marks by type and excludes tab-stop definitions", async () => {
    await withTempDir(async (dir) => {
        const root = path.join(dir, "u");
        await writeXml(
            path.join(root, "word", "document.xml"),
            doc(
                `<w:p><w:pPr><w:tabs><w:tab w:val="left" w:pos="720"/></w:tabs></w:pPr>` +
                    `<w:r><w:br/><w:br w:type="page"/><w:br w:type="column"/><w:tab/><w:sym w:font="Wingdings" w:char="F0E0"/><w:cr/><w:softHyphen/><w:noBreakHyphen/></w:r>` +
                    `<w:r><w:br w:type="separator"/></w:r></w:p>`,
            ),
        );
        const inv = await collectDocxSemanticInventory(root);
        const get = (label: string) =>
            [...inv.counters.values()].find((c) => c.category === "inline mark" && c.label === label)?.count ?? 0;
        expect(get("line break")).toBe(1); // the bare <w:br/>
        expect(get("page break")).toBe(1);
        expect(get("column break")).toBe(1);
        expect(get("tab")).toBe(1); // only the run-level <w:tab/>, NOT the <w:tabs> stop
        expect(get("symbol")).toBe(1);
        expect(get("carriage return")).toBe(1);
        expect(get("soft hyphen")).toBe(1);
        expect(get("non-breaking hyphen")).toBe(1);
        // separator break is excluded entirely
        expect([...inv.counters.values()].some((c) => c.label.includes("separator"))).toBe(false);
    });
});
```

- [ ] **Step 2: Run it, verify failure**

Run: `bunx vitest run tests/docx-diagnostics.test.ts -t "in-run atomic marks"`
Expected: FAIL — all `get(...)` return 0 because no `inline mark` counters exist yet.

- [ ] **Step 3: Implement the collector**

Add this function to `docx-diagnostics.ts`:

```ts
function collectInlineMarks(rel: string, dom: Document, inventory: MutableDocxSemanticInventory): void {
    for (const ns of WORD_NAMESPACES) {
        const brs = dom.getElementsByTagNameNS(ns, "br");
        for (let i = 0; i < brs.length; i += 1) {
            const br = brs.item(i);
            if (!br) continue;
            const type = br.getAttributeNS(ns, "type") ?? br.getAttribute("w:type") ?? br.getAttribute("type") ?? "textWrapping";
            if (type === "separator" || type === "continuationSeparator") continue;
            const label = type === "page" ? "page break" : type === "column" ? "column break" : "line break";
            addCounter(inventory, rel, "inline mark", label, "occurrence(s)", 1);
        }
        // <w:tab/> as a direct child of <w:r> is a tab CHARACTER; <w:tab> inside
        // <w:tabs> (paragraph properties) is a tab-STOP definition — exclude those.
        const tabEls = dom.getElementsByTagNameNS(ns, "tab");
        let tabs = 0;
        for (let i = 0; i < tabEls.length; i += 1) {
            const el = tabEls.item(i);
            if (!el) continue;
            const parent = el.parentNode;
            if (parent && parent.nodeType === 1 && isWordElement(parent as Element) && localName(parent as Element) === "r") tabs += 1;
        }
        if (tabs > 0) addCounter(inventory, rel, "inline mark", "tab", "occurrence(s)", tabs);
        const simple: ReadonlyArray<readonly [string, string]> = [
            ["sym", "symbol"],
            ["cr", "carriage return"],
            ["softHyphen", "soft hyphen"],
            ["noBreakHyphen", "non-breaking hyphen"],
        ];
        for (const [local, label] of simple) {
            const count = dom.getElementsByTagNameNS(ns, local).length;
            if (count > 0) addCounter(inventory, rel, "inline mark", label, "occurrence(s)", count);
        }
    }
}
```

- [ ] **Step 4: Wire it into the dispatcher**

In `collectXmlPart`, add the call after `collectFormatting(rel, dom, inventory);`:

```ts
    collectFormatting(rel, dom, inventory);
    collectInlineMarks(rel, dom, inventory);
```

- [ ] **Step 5: Run it, verify pass**

Run: `bunx vitest run tests/docx-diagnostics.test.ts -t "in-run atomic marks"`
Expected: PASS.

- [ ] **Step 6: Run the whole diagnostics spec for no regression**

Run: `bunx vitest run tests/docx-diagnostics.test.ts`
Expected: all PASS.

---

## Task A3: Family 2 — table shape collector

**Files:**
- Modify: `src/scripts/office/validators/docx-diagnostics.ts`
- Test: `tests/docx-diagnostics.test.ts`

- [ ] **Step 1: Write the failing test**

```ts
it("captures table shape as rows×cols with a tblGrid-absent fallback", async () => {
    await withTempDir(async (dir) => {
        const root = path.join(dir, "u");
        await writeXml(
            path.join(root, "word", "document.xml"),
            doc(
                // table A: explicit 1-row, 3-col grid, one cell spans 2
                `<w:tbl><w:tblGrid><w:gridCol/><w:gridCol/><w:gridCol/></w:tblGrid>` +
                    `<w:tr><w:tc><w:tcPr><w:gridSpan w:val="2"/></w:tcPr></w:tc><w:tc/></w:tr></w:tbl>` +
                    // table B: NO tblGrid, 2 rows; first row has 2 cells, one with gridSpan=2 → 3 cols
                    `<w:tbl><w:tr><w:tc><w:tcPr><w:gridSpan w:val="2"/></w:tcPr></w:tc><w:tc/></w:tr><w:tr><w:tc/><w:tc/><w:tc/></w:tr></w:tbl>`,
            ),
        );
        const inv = await collectDocxSemanticInventory(root);
        const shapes = [...inv.counters.values()].filter((c) => c.category === "table shape");
        const shape = (label: string) => shapes.find((c) => c.label === label)?.count ?? 0;
        expect(shape("table 1×3")).toBe(1); // table A
        expect(shape("table 2×3")).toBe(1); // table B via gridSpan fallback
        expect(shape("merged cell gridSpan=2")).toBe(2); // one in each table
    });
});
```

- [ ] **Step 2: Run it, verify failure**

Run: `bunx vitest run tests/docx-diagnostics.test.ts -t "table shape as rows"`
Expected: FAIL — no `table shape` counters.

- [ ] **Step 3: Implement the collector**

Add this function. Note the merged-cell scan is a single loop **after** the table loop (one `getElementsByTagNameNS("tc")` per part), so cells are not double-counted:

```ts
function collectTableShape(rel: string, dom: Document, inventory: MutableDocxSemanticInventory): void {
    for (const ns of WORD_NAMESPACES) {
        const tables = dom.getElementsByTagNameNS(ns, "tbl");
        for (let i = 0; i < tables.length; i += 1) {
            const tbl = tables.item(i);
            if (!tbl) continue;
            const rows = directWordChildren(tbl, "tr");
            const rowCount = rows.length;
            const grid = directWordChild(tbl, "tblGrid");
            let cols = grid ? directWordChildren(grid, "gridCol").length : 0;
            if (cols === 0 && rows[0]) {
                // Fallback for tables with no/empty tblGrid: sum gridSpan across the first row.
                let sum = 0;
                for (const tc of directWordChildren(rows[0], "tc")) {
                    const tcPr = directWordChild(tc, "tcPr");
                    const span = tcPr ? wordChildAttr(tcPr, "gridSpan", "val") : null;
                    const n = span ? Number.parseInt(span, 10) : 1;
                    sum += Number.isFinite(n) && n > 0 ? n : 1;
                }
                cols = sum;
            }
            addCounter(inventory, rel, "table shape", `table ${rowCount}×${cols}`, "table(s)", 1);
        }
        const cells = dom.getElementsByTagNameNS(ns, "tc");
        for (let i = 0; i < cells.length; i += 1) {
            const tc = cells.item(i);
            if (!tc) continue;
            const tcPr = directWordChild(tc, "tcPr");
            if (!tcPr) continue;
            const span = wordChildAttr(tcPr, "gridSpan", "val");
            const spanN = span ? Number.parseInt(span, 10) : 1;
            if (Number.isFinite(spanN) && spanN > 1) {
                addCounter(inventory, rel, "table shape", `merged cell gridSpan=${spanN}`, "cell(s)", 1);
            }
            if (directWordChild(tcPr, "vMerge")) {
                addCounter(inventory, rel, "table shape", "merged cell vMerge", "cell(s)", 1);
            }
        }
    }
}
```

- [ ] **Step 4: Wire it into the dispatcher**

In `collectXmlPart`, after `collectInlineMarks(rel, dom, inventory);`:

```ts
    collectInlineMarks(rel, dom, inventory);
    collectTableShape(rel, dom, inventory);
```

- [ ] **Step 5: Run it, verify pass**

Run: `bunx vitest run tests/docx-diagnostics.test.ts -t "table shape as rows"`
Expected: PASS.

---

## Task A4: `profile` argument + Families 3 & 4 (strict-only)

**Files:**
- Modify: `src/scripts/office/validators/docx-diagnostics.ts`
- Test: `tests/docx-diagnostics.test.ts`

- [ ] **Step 1: Write the failing test**

Add the import of `Profile` is not needed in the test (it passes a string literal). Add:

```ts
it("collects section geometry and image shape only under the strict profile", async () => {
    await withTempDir(async (dir) => {
        const root = path.join(dir, "u");
        await writeXml(
            path.join(root, "word", "document.xml"),
            doc(
                `<w:p><w:r><w:drawing><wp:inline xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing">` +
                    `<wp:extent cx="1905000" cy="1270000"/></wp:inline></w:drawing></w:r></w:p>` +
                    `<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/><w:cols w:num="2"/></w:sectPr>`,
            ),
        );
        const lenient = await collectDocxSemanticInventory(root); // default lenient
        const strict = await collectDocxSemanticInventory(root, "strict");
        const has = (inv: Awaited<ReturnType<typeof collectDocxSemanticInventory>>, category: string) =>
            [...inv.counters.values()].some((c) => c.category === category);

        expect(has(lenient, "section geometry")).toBe(false);
        expect(has(lenient, "image shape")).toBe(false);
        expect(has(strict, "section geometry")).toBe(true);
        expect(has(strict, "image shape")).toBe(true);

        const strictVals = [...strict.counters.values()];
        expect(strictVals.find((c) => c.category === "section geometry" && c.label === "section portrait 12240×15840")?.count).toBe(1);
        expect(strictVals.find((c) => c.category === "section geometry" && c.label === "section columns=2")?.count).toBe(1);
        expect(strictVals.find((c) => c.category === "image shape" && c.label === "image ~1905000×1270000 inline")?.count).toBe(1);

        const wordValid = await collectDocxSemanticInventory(root, "word-valid");
        expect(has(wordValid, "section geometry")).toBe(false); // word-valid behaves like lenient
    });
});
```

- [ ] **Step 2: Run it, verify failure**

Run: `bunx vitest run tests/docx-diagnostics.test.ts -t "section geometry and image shape only"`
Expected: FAIL — `collectDocxSemanticInventory` takes one arg / no such categories.

- [ ] **Step 3: Add the `Profile` import**

At the top of `docx-diagnostics.ts`, extend the type import:

```ts
import type { Profile, ValidationIssue } from "../../../lib/types";
```

- [ ] **Step 4: Add the two strict-only collectors**

```ts
function collectSectionGeometry(rel: string, dom: Document, inventory: MutableDocxSemanticInventory): void {
    for (const ns of WORD_NAMESPACES) {
        const sects = dom.getElementsByTagNameNS(ns, "sectPr");
        for (let i = 0; i < sects.length; i += 1) {
            const sect = sects.item(i);
            if (!sect) continue;
            const pgSz = directWordChild(sect, "pgSz");
            if (pgSz) {
                const w = wordChildAttrSelf(pgSz, "w") ?? "?";
                const h = wordChildAttrSelf(pgSz, "h") ?? "?";
                const orient = wordChildAttrSelf(pgSz, "orient") ?? "portrait";
                addCounter(inventory, rel, "section geometry", `section ${orient} ${w}×${h}`, "section(s)", 1);
            }
            const pgMar = directWordChild(sect, "pgMar");
            if (pgMar) {
                const t = wordChildAttrSelf(pgMar, "top") ?? "?";
                const r = wordChildAttrSelf(pgMar, "right") ?? "?";
                const b = wordChildAttrSelf(pgMar, "bottom") ?? "?";
                const l = wordChildAttrSelf(pgMar, "left") ?? "?";
                addCounter(inventory, rel, "section geometry", `section margins T${t} R${r} B${b} L${l}`, "section(s)", 1);
            }
            const cols = directWordChild(sect, "cols");
            const num = cols ? wordChildAttrSelf(cols, "num") ?? "1" : "1";
            addCounter(inventory, rel, "section geometry", `section columns=${num}`, "section(s)", 1);
        }
    }
}

function collectImageShape(rel: string, dom: Document, inventory: MutableDocxSemanticInventory): void {
    const extents = dom.getElementsByTagNameNS(WP_NAMESPACE, "extent");
    for (let i = 0; i < extents.length; i += 1) {
        const ext = extents.item(i);
        if (!ext) continue;
        const parent = ext.parentNode;
        if (!parent || parent.nodeType !== 1) continue;
        const parentLocal = localName(parent as Element);
        const wrap = parentLocal === "inline" ? "inline" : parentLocal === "anchor" ? "anchor" : null;
        if (!wrap) continue; // ignore a:extent / other extents not under wp:inline|wp:anchor
        const cx = roundEmu(ext.getAttribute("cx"));
        const cy = roundEmu(ext.getAttribute("cy"));
        addCounter(inventory, rel, "image shape", `image ~${cx}×${cy} ${wrap}`, "image(s)", 1);
    }
}

function roundEmu(raw: string | null): number {
    const n = raw ? Number.parseInt(raw, 10) : 0;
    if (!Number.isFinite(n)) return 0;
    return Math.round(n / 1000) * 1000; // ≈0.1mm bucket absorbs tool re-rounding
}
```

- [ ] **Step 5: Add the `wordChildAttrSelf` helper**

`wordChildAttr` reads an attribute off a *child*; here we need an attribute off the element itself. Add:

```ts
function wordChildAttrSelf(elem: Element, attr: string): string | null {
    return elem.getAttributeNS(wordNamespace(elem), attr) ?? elem.getAttribute(`w:${attr}`) ?? elem.getAttribute(attr);
}
```

- [ ] **Step 6: Thread `profile` through `collectXmlPart` and `collectDocxSemanticInventory`**

Change the signature and dispatch. Replace the existing `collectDocxSemanticInventory` and `collectXmlPart`:

```ts
export async function collectDocxSemanticInventory(unpackedDir: string, profile: Profile = "lenient"): Promise<DocxSemanticInventory> {
    const inventory: MutableDocxSemanticInventory = { counters: new Map() };
    const files = await walkFiles(unpackedDir);
    for (const file of files) {
        const rel = path.relative(unpackedDir, file).split(path.sep).join("/");
        if (!rel.endsWith(".xml") && !rel.endsWith(".rels")) {
            await collectPackageAsset(file, rel, inventory);
            continue;
        }
        let dom: Document;
        try {
            dom = parseXml(await fs.readFile(file, "utf-8"));
        } catch {
            continue;
        }
        collectXmlPart(rel, dom, inventory, profile);
    }
    return { counters: new Map([...inventory.counters].sort(([a], [b]) => a.localeCompare(b))) };
}
```

```ts
function collectXmlPart(rel: string, dom: Document, inventory: MutableDocxSemanticInventory, profile: Profile): void {
    if (rel.endsWith(".rels")) {
        collectRelationships(rel, dom, inventory);
        return;
    }
    if (rel === "[Content_Types].xml") collectContentTypes(rel, dom, inventory);
    collectDocumentStructure(rel, dom, inventory);
    collectText(rel, dom, inventory);
    collectFormatting(rel, dom, inventory);
    collectInlineMarks(rel, dom, inventory);
    collectTableShape(rel, dom, inventory);
    collectStyles(rel, dom, inventory);
    collectComments(rel, dom, inventory);
    collectTrackedChanges(rel, dom, inventory);
    if (profile === "strict") {
        collectSectionGeometry(rel, dom, inventory);
        collectImageShape(rel, dom, inventory);
    }
}
```

- [ ] **Step 7: Run it, verify pass**

Run: `bunx vitest run tests/docx-diagnostics.test.ts -t "section geometry and image shape only"`
Expected: PASS.

- [ ] **Step 8: Full diagnostics spec + type-check**

Run: `bunx vitest run tests/docx-diagnostics.test.ts && bunx tsc --noEmit`
Expected: all PASS, tsc exit 0.

---

## Task A5: Thread `profile` into the repair path in `validate.ts`

**Files:**
- Modify: `src/scripts/office/validate.ts` (the `if (opts.autoRepair)` block, ~lines 219 & 227)

- [ ] **Step 1: Update both collect calls**

`runValidation` already has `profile` in scope (`const profile: Profile = opts.profile ?? DEFAULT_PROFILE;` near the top of the file). In the repair block, change:

```ts
        const beforeInventory = (opts.suffix === ".docx" || opts.suffix === ".docm") ? await collectDocxSemanticInventory(unpackedDir) : null;
```
to
```ts
        const beforeInventory = (opts.suffix === ".docx" || opts.suffix === ".docm") ? await collectDocxSemanticInventory(unpackedDir, profile) : null;
```

and
```ts
                ? compareDocxSemanticInventories(beforeInventory, await collectDocxSemanticInventory(unpackedDir))
```
to
```ts
                ? compareDocxSemanticInventories(beforeInventory, await collectDocxSemanticInventory(unpackedDir, profile))
```

> If the symbol in scope is named differently than `profile`, use `grep -n "profile" src/scripts/office/validate.ts` to confirm; the spec-confirmed line is `const profile: Profile = opts.profile ?? DEFAULT_PROFILE;`.

- [ ] **Step 2: Type-check**

Run: `bunx tsc --noEmit`
Expected: exit 0.

- [ ] **Step 3: Commit A**

```bash
git add src/scripts/office/validators/docx-diagnostics.ts src/scripts/office/validate.ts tests/docx-diagnostics.test.ts
git commit -m "feat: add atomic-mark, table-shape, section, and image collectors to docx inventory

New families: in-run atomic marks (always), table shape with tblGrid-absent
fallback (always), section geometry and image shape (strict profile only).
collectDocxSemanticInventory gains an optional profile arg threaded from
validate.ts's repair path.

Co-Authored-By: Claude Opus 4.7 (1M context) <noreply@anthropic.com>"
```

---

# Commit B — Severity, symmetric diff, formatters, gate split

## Task B1: `severityClassFor` (shared classifier)

**Files:**
- Modify: `src/scripts/office/validators/docx-diagnostics.ts`
- Test: `tests/docx-diagnostics.test.ts`

- [ ] **Step 1: Write the failing test**

Add to the imports in `tests/docx-diagnostics.test.ts`:

```ts
import {
    buildRepairPlanIssues,
    collectDocxSemanticInventory,
    compareDocxSemanticInventories,
    severityClassFor,
} from "../src/scripts/office/validators/docx-diagnostics";
```

Add the test:

```ts
it("classifies categories into severity classes", () => {
    expect(severityClassFor("text")).toBe("content");
    expect(severityClassFor("document structure")).toBe("content");
    expect(severityClassFor("tracked change")).toBe("content");
    expect(severityClassFor("content type")).toBe("content");
    expect(severityClassFor("numbering")).toBe("content");
    expect(severityClassFor("table shape")).toBe("table-shape");
    expect(severityClassFor("section geometry")).toBe("section-geometry");
    expect(severityClassFor("image shape")).toBe("image-shape");
    expect(severityClassFor("formatting")).toBe("formatting");
    expect(severityClassFor("style reference")).toBe("formatting");
    expect(severityClassFor("inline mark")).toBe("atomic-marks");
    expect(severityClassFor("package asset")).toBe("bookkeeping");
    expect(severityClassFor("relationship")).toBe("bookkeeping");
    expect(severityClassFor("comment thread")).toBe("bookkeeping");
    expect(severityClassFor("comment")).toBe("content");
});
```

- [ ] **Step 2: Run it, verify failure**

Run: `bunx vitest run tests/docx-diagnostics.test.ts -t "classifies categories"`
Expected: FAIL — `severityClassFor` not exported.

- [ ] **Step 3: Implement and export the classifier**

Add to `docx-diagnostics.ts`:

```ts
export type InventorySeverityClass =
    | "content"
    | "table-shape"
    | "section-geometry"
    | "image-shape"
    | "formatting"
    | "atomic-marks"
    | "bookkeeping";

const SEVERITY_CLASS_BY_CATEGORY: Record<string, InventorySeverityClass> = {
    text: "content",
    "document structure": "content",
    comment: "content",
    "comment marker": "content",
    "tracked change": "content",
    "content type": "content",
    numbering: "content",
    formatting: "formatting",
    "style reference": "formatting",
    "style definition": "formatting",
    "style formatting": "formatting",
    "table shape": "table-shape",
    "section geometry": "section-geometry",
    "image shape": "image-shape",
    "inline mark": "atomic-marks",
    "package asset": "bookkeeping",
    relationship: "bookkeeping",
    "comment thread": "bookkeeping",
};

export function severityClassFor(category: string): InventorySeverityClass {
    return SEVERITY_CLASS_BY_CATEGORY[category] ?? "content";
}
```

> Default to `"content"` for unknown categories so a future collector that forgets to register defaults to the safest (error-on-loss) tier.

- [ ] **Step 4: Run it, verify pass**

Run: `bunx vitest run tests/docx-diagnostics.test.ts -t "classifies categories"`
Expected: PASS.

---

## Task B2: Split the repair gate codes by class

**Files:**
- Modify: `src/scripts/office/validators/docx-diagnostics.ts` (`compareDocxSemanticInventories`)
- Test: `tests/docx-diagnostics.test.ts`

- [ ] **Step 1: Write the failing test**

```ts
it("tiers repair loss: content → error, other classes → fidelity warning", async () => {
    await withTempDir(async (dir) => {
        const before = path.join(dir, "before");
        const after = path.join(dir, "after");
        await writeXml(
            path.join(before, "word", "document.xml"),
            doc(`<w:p><w:r><w:t>hello</w:t><w:br/></w:r></w:p><w:p><w:r><w:t>world</w:t></w:r></w:p>`),
        );
        // after: lost one paragraph (content) AND the line break (atomic mark)
        await writeXml(path.join(after, "word", "document.xml"), doc(`<w:p><w:r><w:t>hello</w:t></w:r></w:p>`));

        const issues = compareDocxSemanticInventories(
            await collectDocxSemanticInventory(before),
            await collectDocxSemanticInventory(after),
        );
        const byCode = (code: string) => issues.filter((i) => i.code === code);
        expect(byCode("repair-content-loss").some((i) => i.severity === "error")).toBe(true);
        expect(byCode("repair-fidelity-loss").some((i) => i.severity === "warning")).toBe(true);
        // the line-break loss must NOT be an error
        expect(byCode("repair-content-loss").some((i) => i.message.includes("line break"))).toBe(false);
    });
});
```

> The project's `Severity` type is `"error" | "warning" | "info"` (`src/lib/types.ts:31`). Use the literal `"warning"`.

- [ ] **Step 2: Run it, verify failure**

Run: `bunx vitest run tests/docx-diagnostics.test.ts -t "tiers repair loss"`
Expected: FAIL — today every loss is `repair-content-loss` error; no `repair-fidelity-loss`.

- [ ] **Step 3: Update `compareDocxSemanticInventories`**

Replace the loop body so each decrease is dispatched by class:

```ts
export function compareDocxSemanticInventories(before: DocxSemanticInventory, after: DocxSemanticInventory): ValidationIssue[] {
    const issues: ValidationIssue[] = [];
    const beforeCounters = [...before.counters.values()].sort(compareCounters);
    for (const counter of beforeCounters) {
        const afterCount = after.counters.get(counterKey(counter.path, counter.category, counter.label, counter.unit))?.count ?? 0;
        if (counter.count <= afterCount) continue;
        const lost = counter.count - afterCount;
        const isContent = severityClassFor(counter.category) === "content";
        issues.push({
            severity: isContent ? "error" : "warning",
            path: counter.path,
            code: isContent ? "repair-content-loss" : "repair-fidelity-loss",
            message:
                `Repair lost ${counter.category} '${counter.label}': ` + `${counter.count} → ${afterCount} (-${lost} ${counter.unit}).`,
        });
    }
    if (issues.length > 0) return issues;
    return [
        {
            severity: "info",
            code: "repair-content-preserved",
            message: `Repair semantic inventory preserved: no tracked content counters decreased across ${before.counters.size} counter(s).`,
        },
    ];
}
```

- [ ] **Step 4: Run it, verify pass + no regression**

Run: `bunx vitest run tests/docx-diagnostics.test.ts`
Expected: all PASS (including the pre-existing "formatting coverage loss" test — note that test expects `repair-content-loss`; formatting is **not** content, so it will now be `repair-fidelity-loss` warning. **Update that pre-existing test** to expect `code: "repair-fidelity-loss"` and `severity: "warning"`; its message assertion stays the same).

- [ ] **Step 5: Fix the pre-existing formatting test**

In the first test ("reports formatting coverage loss…"), change:
```ts
            expect(issues[0]).toMatchObject({
                severity: "error",
                code: "repair-content-loss",
                path: "word/document.xml",
            });
```
to
```ts
            expect(issues[0]).toMatchObject({
                severity: "warning",
                code: "repair-fidelity-loss",
                path: "word/document.xml",
            });
```

Run: `bunx vitest run tests/docx-diagnostics.test.ts`
Expected: all PASS.

---

## Task B3: New module — `diffDocxInventories`

**Files:**
- Create: `src/scripts/office/validators/docx-inventory-diff.ts`
- Test: `tests/docx-inventory-diff.test.ts`

- [ ] **Step 1: Write the failing test**

Create `tests/docx-inventory-diff.test.ts`:

```ts
import { describe, expect, it } from "vitest";
import type { DocxSemanticInventory } from "../src/scripts/office/validators/docx-diagnostics";
import { diffDocxInventories } from "../src/scripts/office/validators/docx-inventory-diff";

function inv(entries: Array<{ path: string; category: string; label: string; unit: string; count: number }>): DocxSemanticInventory {
    const counters = new Map<string, { path: string; category: string; label: string; unit: string; count: number }>();
    for (const e of entries) counters.set(`${e.path}\u0000${e.category}\u0000${e.label}\u0000${e.unit}`, e);
    return { counters };
}

describe("diffDocxInventories", () => {
    it("splits added / removed / changed / unchanged by key and direction", () => {
        const before = inv([
            { path: "word/document.xml", category: "document structure", label: "paragraph", unit: "element(s)", count: 10 },
            { path: "word/document.xml", category: "table shape", label: "table 3×4", unit: "table(s)", count: 1 },
            { path: "word/document.xml", category: "inline mark", label: "tab", unit: "occurrence(s)", count: 5 },
        ]);
        const after = inv([
            { path: "word/document.xml", category: "document structure", label: "paragraph", unit: "element(s)", count: 8 }, // changed ↓
            { path: "word/document.xml", category: "table shape", label: "table 3×2", unit: "table(s)", count: 1 }, // reshape: 3×4 removed, 3×2 added
            { path: "word/document.xml", category: "inline mark", label: "tab", unit: "occurrence(s)", count: 5 }, // unchanged
        ]);
        const d = diffDocxInventories(before, after);
        expect(d.added.map((x) => x.label)).toEqual(["table 3×2"]);
        expect(d.removed.map((x) => x.label)).toEqual(["table 3×4"]);
        expect(d.changed.map((x) => x.label)).toEqual(["paragraph"]);
        expect(d.changed[0]).toMatchObject({ before: 10, after: 8 });
        expect(d.unchangedCount).toBe(1);
    });

    it("treats identical inventories as fully unchanged", () => {
        const a = inv([{ path: "p", category: "text", label: "visible text", unit: "character(s)", count: 3 }]);
        const d = diffDocxInventories(a, a);
        expect(d.added).toHaveLength(0);
        expect(d.removed).toHaveLength(0);
        expect(d.changed).toHaveLength(0);
        expect(d.unchangedCount).toBe(1);
    });
});
```

- [ ] **Step 2: Run it, verify failure**

Run: `bunx vitest run tests/docx-inventory-diff.test.ts -t "splits added"`
Expected: FAIL — module/function does not exist.

- [ ] **Step 3: Create the module with types + diff**

Create `src/scripts/office/validators/docx-inventory-diff.ts`:

```ts
import type { DocxSemanticCounter, DocxSemanticInventory } from "./docx-diagnostics";

export interface DocxInventoryDelta {
    key: string;
    path: string;
    category: string;
    label: string;
    unit: string;
    before: number;
    after: number;
}

export interface DocxInventoryDiff {
    added: DocxInventoryDelta[];
    removed: DocxInventoryDelta[];
    changed: DocxInventoryDelta[];
    unchangedCount: number;
}

function keyOf(c: DocxSemanticCounter): string {
    return `${c.path}\u0000${c.category}\u0000${c.label}\u0000${c.unit}`;
}

function sortDeltas(a: DocxInventoryDelta, b: DocxInventoryDelta): number {
    return a.path.localeCompare(b.path) || a.category.localeCompare(b.category) || a.label.localeCompare(b.label);
}

export function diffDocxInventories(before: DocxSemanticInventory, after: DocxSemanticInventory): DocxInventoryDiff {
    const added: DocxInventoryDelta[] = [];
    const removed: DocxInventoryDelta[] = [];
    const changed: DocxInventoryDelta[] = [];
    let unchangedCount = 0;

    for (const b of before.counters.values()) {
        const key = keyOf(b);
        const a = after.counters.get(key);
        const afterCount = a?.count ?? 0;
        const delta: DocxInventoryDelta = {
            key,
            path: b.path,
            category: b.category,
            label: b.label,
            unit: b.unit,
            before: b.count,
            after: afterCount,
        };
        if (afterCount === 0) removed.push(delta);
        else if (afterCount !== b.count) changed.push(delta);
        else unchangedCount += 1;
    }
    for (const a of after.counters.values()) {
        const key = keyOf(a);
        if (before.counters.has(key)) continue;
        added.push({ key, path: a.path, category: a.category, label: a.label, unit: a.unit, before: 0, after: a.count });
    }
    added.sort(sortDeltas);
    removed.sort(sortDeltas);
    changed.sort(sortDeltas);
    return { added, removed, changed, unchangedCount };
}
```

- [ ] **Step 4: Run it, verify pass**

Run: `bunx vitest run tests/docx-inventory-diff.test.ts`
Expected: PASS.

---

## Task B4: `severityFor` + `inventoryDiffToIssues`

**Files:**
- Modify: `src/scripts/office/validators/docx-inventory-diff.ts`
- Test: `tests/docx-inventory-diff.test.ts`

- [ ] **Step 1: Write the failing test**

```ts
import { diffDocxInventories, inventoryDiffToIssues } from "../src/scripts/office/validators/docx-inventory-diff";

it("tiers diff deltas into severity-graded issues", () => {
    const before = inv([
        { path: "word/document.xml", category: "document structure", label: "paragraph", unit: "element(s)", count: 10 },
        { path: "word/document.xml", category: "formatting", label: "bold", unit: "formatted character(s)", count: 4 },
        { path: "word/document.xml", category: "table shape", label: "table 3×2", unit: "table(s)", count: 1 },
    ]);
    const after = inv([
        { path: "word/document.xml", category: "document structure", label: "paragraph", unit: "element(s)", count: 8 }, // content ↓ → error
        { path: "word/document.xml", category: "formatting", label: "bold", unit: "formatted character(s)", count: 1 }, // formatting ↓ → warn
        { path: "word/document.xml", category: "table shape", label: "table 3×4", unit: "table(s)", count: 1 }, // shape change → warn (both)
    ]);
    const issues = inventoryDiffToIssues(diffDocxInventories(before, after));
    const find = (code: string) => issues.find((i) => i.code === code);
    expect(find("inventory-content-loss")?.severity).toBe("error");
    expect(find("inventory-formatting-drift")?.severity).toBe("warning");
    expect(issues.filter((i) => i.code === "inventory-shape-change").every((i) => i.severity === "warning")).toBe(true);
    // a content GAIN is a warning, never an error
    const grew = inventoryDiffToIssues(diffDocxInventories(after, before));
    expect(grew.find((i) => i.code === "inventory-content-added")?.severity).toBe("warning");
    expect(grew.some((i) => i.severity === "error")).toBe(false); // table 3×4→3×2 + paragraph growth must not error
});
```

- [ ] **Step 2: Run it, verify failure**

Run: `bunx vitest run tests/docx-inventory-diff.test.ts -t "tiers diff deltas"`
Expected: FAIL — `inventoryDiffToIssues` not defined.

- [ ] **Step 3: Implement `severityFor` + `inventoryDiffToIssues`**

Append to `docx-inventory-diff.ts`:

```ts
import { severityClassFor } from "./docx-diagnostics";
import type { InventorySeverityClass } from "./docx-diagnostics";
import type { ValidationIssue } from "../../../lib/types";

type Direction = "decrease" | "increase";

interface TierResult {
    severity: ValidationIssue["severity"];
    code: string;
}

export function severityFor(category: string, direction: Direction): TierResult {
    const cls: InventorySeverityClass = severityClassFor(category);
    switch (cls) {
        case "content":
            return direction === "decrease"
                ? { severity: "error", code: "inventory-content-loss" }
                : { severity: "warning", code: "inventory-content-added" };
        case "table-shape":
        case "section-geometry":
        case "image-shape":
            return { severity: "warning", code: "inventory-shape-change" };
        case "formatting":
            return direction === "decrease"
                ? { severity: "warning", code: "inventory-formatting-drift" }
                : { severity: "info", code: "inventory-formatting-drift" };
        case "atomic-marks":
            return direction === "decrease"
                ? { severity: "warning", code: "inventory-mark-drift" }
                : { severity: "info", code: "inventory-mark-drift" };
        case "bookkeeping":
            return direction === "decrease"
                ? { severity: "warning", code: "inventory-bookkeeping-drift" }
                : { severity: "info", code: "inventory-bookkeeping-drift" };
    }
}

function directionOf(d: DocxInventoryDelta): Direction {
    return d.after < d.before ? "decrease" : "increase";
}

export function inventoryDiffToIssues(diff: DocxInventoryDiff): ValidationIssue[] {
    const issues: ValidationIssue[] = [];
    const emit = (d: DocxInventoryDelta, direction: Direction): void => {
        const tier = severityFor(d.category, direction);
        issues.push({
            severity: tier.severity,
            code: tier.code,
            path: d.path,
            message: `${d.category} '${d.label}': ${d.before} → ${d.after} (${d.unit}).`,
        });
    };
    for (const d of diff.removed) emit(d, "decrease");
    for (const d of diff.added) emit(d, "increase");
    for (const d of diff.changed) emit(d, directionOf(d));
    return issues;
}
```

- [ ] **Step 4: Run it, verify pass**

Run: `bunx vitest run tests/docx-inventory-diff.test.ts -t "tiers diff deltas"`
Expected: PASS.

---

## Task B5: `formatInventoryDiffMarkdown`

**Files:**
- Modify: `src/scripts/office/validators/docx-inventory-diff.ts`
- Test: `tests/docx-inventory-diff.test.ts`

- [ ] **Step 1: Write the failing test**

```ts
import { formatInventoryDiffMarkdown } from "../src/scripts/office/validators/docx-inventory-diff";

it("renders a markdown report with sections, severity prefixes, and a summary", () => {
    const before = inv([
        { path: "word/document.xml", category: "document structure", label: "paragraph", unit: "element(s)", count: 10 },
        { path: "word/document.xml", category: "table shape", label: "table 3×4", unit: "table(s)", count: 1 },
    ]);
    const after = inv([
        { path: "word/document.xml", category: "document structure", label: "paragraph", unit: "element(s)", count: 8 },
        { path: "word/document.xml", category: "table shape", label: "table 3×2", unit: "table(s)", count: 1 },
    ]);
    const md = formatInventoryDiffMarkdown(diffDocxInventories(before, after));
    expect(md).toContain("## Removed");
    expect(md).toContain("## Added");
    expect(md).toContain("## Changed");
    expect(md).toContain("## Reshaped"); // shape loss+gain paired for readability
    expect(md).toMatch(/table 3×4.*→.*table 3×2/);
    expect(md).toMatch(/1 added, 1 removed, 1 changed, 0 unchanged/);
    expect(md).toMatch(/🔴|🟠|⚪/); // severity prefixes present
});
```

- [ ] **Step 2: Run it, verify failure**

Run: `bunx vitest run tests/docx-inventory-diff.test.ts -t "renders a markdown report"`
Expected: FAIL — `formatInventoryDiffMarkdown` not defined.

- [ ] **Step 3: Implement the formatter**

Append to `docx-inventory-diff.ts`:

```ts
const SEV_PREFIX: Record<ValidationIssue["severity"], string> = {
    error: "🔴",
    warning: "🟠",
    info: "⚪",
};

function renderSection(title: string, deltas: DocxInventoryDelta[]): string[] {
    if (deltas.length === 0) return [];
    const lines = [`## ${title}`, ""];
    let lastPath = "";
    for (const d of deltas) {
        if (d.path !== lastPath) {
            lines.push(`### ${d.path}`);
            lastPath = d.path;
        }
        // directionOf classifies removed (→decrease), added (→increase), and
        // changed deltas alike, so the per-line severity is always correct.
        const tier = severityFor(d.category, directionOf(d));
        lines.push(`- ${SEV_PREFIX[tier.severity]} ${d.category} \`${d.label}\`: ${d.before} → ${d.after} (${d.unit})`);
    }
    lines.push("");
    return lines;
}

const SHAPE_CATEGORIES = new Set(["table shape", "section geometry", "image shape"]);

function renderReshaped(diff: DocxInventoryDiff): string[] {
    // Display-only pairing: within each (path, category) of a shape class, zip
    // removed shape-keys with added shape-keys. No identity guarantee.
    const groups = new Map<string, { removed: DocxInventoryDelta[]; added: DocxInventoryDelta[] }>();
    const push = (d: DocxInventoryDelta, side: "removed" | "added"): void => {
        if (!SHAPE_CATEGORIES.has(d.category)) return;
        const k = `${d.path}\u0000${d.category}`;
        const g = groups.get(k) ?? { removed: [], added: [] };
        g[side].push(d);
        groups.set(k, g);
    };
    for (const d of diff.removed) push(d, "removed");
    for (const d of diff.added) push(d, "added");
    const lines: string[] = [];
    for (const [k, g] of [...groups.entries()].sort(([a], [b]) => a.localeCompare(b))) {
        if (g.removed.length === 0 || g.added.length === 0) continue;
        const [pathValue, category] = k.split("\u0000");
        if (lines.length === 0) lines.push("## Reshaped", "");
        lines.push(`### ${pathValue} — ${category}`);
        const n = Math.max(g.removed.length, g.added.length);
        for (let i = 0; i < n; i += 1) {
            lines.push(`- 🟠 \`${g.removed[i]?.label ?? "—"}\` → \`${g.added[i]?.label ?? "—"}\``);
        }
    }
    if (lines.length > 0) lines.push("");
    return lines;
}

export function formatInventoryDiffMarkdown(diff: DocxInventoryDiff): string {
    const issues = inventoryDiffToIssues(diff);
    const errorCount = issues.filter((i) => i.severity === "error").length;
    const warnCount = issues.filter((i) => i.severity === "warning").length;
    const lines: string[] = ["# DOCX inventory diff", ""];
    lines.push(
        `**Summary:** ${diff.added.length} added, ${diff.removed.length} removed, ${diff.changed.length} changed, ${diff.unchangedCount} unchanged; ${errorCount} error / ${warnCount} warn`,
        "",
    );
    lines.push(...renderReshaped(diff));
    lines.push(...renderSection("Removed", diff.removed));
    lines.push(...renderSection("Added", diff.added));
    lines.push(...renderSection("Changed", diff.changed));
    return lines.join("\n");
}
```

> Shape entries also appear in the plain Removed/Added sections; the Reshaped section is an additional readability summary, not a replacement.

- [ ] **Step 4: Run it, verify pass + full diff spec**

Run: `bunx vitest run tests/docx-inventory-diff.test.ts`
Expected: all PASS.

- [ ] **Step 6: Type-check + commit B**

```bash
bunx tsc --noEmit
git add src/scripts/office/validators/docx-diagnostics.ts src/scripts/office/validators/docx-inventory-diff.ts tests/docx-diagnostics.test.ts tests/docx-inventory-diff.test.ts
git commit -m "feat: symmetric docx inventory diff with severity tiers and formatters

Adds severityClassFor (shared), diffDocxInventories (added/removed/changed),
severityFor, inventoryDiffToIssues, and formatInventoryDiffMarkdown. Splits the
repair gate into repair-content-loss (error) and repair-fidelity-loss (warn).

Co-Authored-By: Claude Opus 4.7 (1M context) <noreply@anthropic.com>"
```

---

# Commit C — `diff-docx` CLI

## Task C1: CLI

**Files:**
- Create: `scripts/diff-docx.ts`
- Test: `tests/diff-docx.cli.test.ts`

- [ ] **Step 1: Write the failing test**

Create `tests/diff-docx.cli.test.ts`:

```ts
import { promises as fs } from "node:fs";
import path from "node:path";
import { describe, expect, it } from "vitest";
import { withTempDir } from "../src/lib/run-cli";
import { runDiffDocx } from "../scripts/diff-docx";

const W_NS = `xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"`;
const docXml = (body: string) => `<?xml version="1.0"?><w:document ${W_NS}><w:body>${body}</w:body></w:document>`;

async function writeUnpacked(dir: string, body: string): Promise<void> {
    await fs.mkdir(path.join(dir, "word"), { recursive: true });
    await fs.writeFile(path.join(dir, "word", "document.xml"), docXml(body), "utf-8");
}

describe("diff-docx CLI", () => {
    it("exits 0 when only warn/info diffs (table reshape)", async () => {
        await withTempDir(async (dir) => {
            const a = path.join(dir, "a");
            const b = path.join(dir, "b");
            await writeUnpacked(a, `<w:tbl><w:tblGrid><w:gridCol/><w:gridCol/></w:tblGrid><w:tr><w:tc/><w:tc/></w:tr></w:tbl>`);
            await writeUnpacked(b, `<w:tbl><w:tblGrid><w:gridCol/><w:gridCol/><w:gridCol/></w:tblGrid><w:tr><w:tc/><w:tc/><w:tc/></w:tr></w:tbl>`);
            const { code, markdown } = await runDiffDocx([a, b]);
            expect(code).toBe(0);
            expect(markdown).toContain("## Added");
        });
    });

    it("exits non-zero when a Content element is lost (paragraph removed)", async () => {
        await withTempDir(async (dir) => {
            const a = path.join(dir, "a");
            const b = path.join(dir, "b");
            await writeUnpacked(a, `<w:p><w:r><w:t>one</w:t></w:r></w:p><w:p><w:r><w:t>two</w:t></w:r></w:p>`);
            await writeUnpacked(b, `<w:p><w:r><w:t>one</w:t></w:r></w:p>`);
            const { code } = await runDiffDocx([a, b]);
            expect(code).toBe(1);
        });
    });
});
```

- [ ] **Step 2: Run it, verify failure**

Run: `bunx vitest run tests/diff-docx.cli.test.ts`
Expected: FAIL — `scripts/diff-docx.ts` / `runDiffDocx` does not exist.

- [ ] **Step 3: Implement the CLI module**

Create `scripts/diff-docx.ts`. It accepts either unpacked directories or packed files; if a path is a file it unpacks into a temp dir first.

```ts
import { promises as fs } from "node:fs";
import path from "node:path";

import { Command } from "commander";
import { runCli, withTempDir } from "../src/lib/run-cli";
import { unpack } from "../src/scripts/office/unpack";
import { collectDocxSemanticInventory } from "../src/scripts/office/validators/docx-diagnostics";
import { diffDocxInventories, formatInventoryDiffMarkdown, inventoryDiffToIssues } from "../src/scripts/office/validators/docx-inventory-diff";
import type { Profile } from "../src/lib/types";

async function isDirectory(p: string): Promise<boolean> {
    try {
        return (await fs.stat(p)).isDirectory();
    } catch {
        return false;
    }
}

async function inventoryOf(input: string, profile: Profile): Promise<ReturnType<typeof collectDocxSemanticInventory>> {
    if (await isDirectory(input)) return collectDocxSemanticInventory(input, profile);
    // packed file → unpack into a temp dir, then collect
    return withTempDir(async (tmp) => {
        const out = path.join(tmp, "unpacked");
        const res = await unpack(input, out);
        if (!res.ok) throw new Error(res.message);
        return collectDocxSemanticInventory(out, profile);
    });
}

export async function runDiffDocx(args: readonly string[]): Promise<{ code: number; markdown: string }> {
    const cmd = new Command();
    cmd.name("diff-docx")
        .description("Symmetric semantic fingerprint diff between two DOCX inputs (packed file or unpacked dir).")
        .argument("<a>", "first document (file or unpacked dir)")
        .argument("<b>", "second document (file or unpacked dir)")
        .option("--profile <profile>", "lenient | strict | word-valid", "lenient")
        .allowExcessArguments(false);
    cmd.exitOverride();
    cmd.parse(args as string[], { from: "user" });
    const opts = cmd.opts<{ profile: string }>();
    const profile = (["lenient", "strict", "word-valid"].includes(opts.profile) ? opts.profile : "lenient") as Profile;
    const [a, b] = cmd.args;

    const diff = diffDocxInventories(await inventoryOf(a, profile), await inventoryOf(b, profile));
    const markdown = formatInventoryDiffMarkdown(diff);
    const hasError = inventoryDiffToIssues(diff).some((i) => i.severity === "error");
    return { code: hasError ? 1 : 0, markdown };
}

runCli(import.meta.url, async () => {
    const { code, markdown } = await runDiffDocx(process.argv.slice(2));
    process.stdout.write(`${markdown}\n`);
    return code;
});
```

- [ ] **Step 4: Run it, verify pass**

Run: `bunx vitest run tests/diff-docx.cli.test.ts`
Expected: PASS.

- [ ] **Step 5: Manual smoke (optional but recommended)**

Run: `bunx tsx scripts/diff-docx.ts tests/fixtures/word-strict/unpacked-working tests/fixtures/word-strict/unpacked-broken --profile strict | head -40`
Expected: a markdown report; exit code reflects whether content was lost.

- [ ] **Step 6: Type-check + commit C**

```bash
bunx tsc --noEmit
git add scripts/diff-docx.ts tests/diff-docx.cli.test.ts
git commit -m "feat: add diff-docx CLI for symmetric DOCX inventory fingerprint

diff-docx <a> <b> [--profile] unpacks packed inputs, collects both inventories,
prints the markdown report, and exits non-zero on any content-loss (error-tier)
difference.

Co-Authored-By: Claude Opus 4.7 (1M context) <noreply@anthropic.com>"
```

---

# Commit D — Manifest regeneration & review

## Task D1: Regenerate and review the fixtures manifest

The new collectors feed the repair gate (Commit A/B). Repairs that drop **content** now error (`repair-content-loss`); repairs that drop marks/formatting/bookkeeping warn (`repair-fidelity-loss`). `fixtures-all` compares validator output to `tests/fixtures-all.manifest.json`, so it must be regenerated.

**Files:**
- Modify: `tests/fixtures-all.manifest.json` (generated)

- [ ] **Step 1: See which fixtures move (validator-side only)**

Run: `bunx vitest run tests/fixtures-all-lenient.test.ts tests/fixtures-all-strict.test.ts 2>&1 | tail -40`
Expected: a list of fixtures whose `errorCodes` / codes changed. Capture it for the commit message.

- [ ] **Step 2: Regenerate the manifest**

> **REQUIRES Word/LibreOffice for the word-probe fields.** On a machine with LibreOffice:
```bash
SOFFICE_AVAILABLE=1 bun run test:fixtures:word   # refresh probe JSONL
bunx tsx scripts/update-manifest.ts              # rebuild manifest from validator + probe
```
> On a machine **without** LibreOffice, do NOT run `update-manifest.ts` (it would blank the word-probe fields). Instead, hand-patch only the changed `errorCodes`/codes arrays for the fixtures identified in Step 1, leaving `word`/`aligned` fields untouched.

- [ ] **Step 3: Review the diff**

Run: `git diff tests/fixtures-all.manifest.json | head -200`
Expected: error-code changes confined to fixtures where a repair drops genuine content; the rest are `repair-fidelity-loss` warnings (which appear in the issues list, not `errorCodes`). Confirm no fixture you expect to be *valid* gained an `xsd-error`/content-loss it shouldn't have. If one did, that is a real collector bug — stop and fix it before committing.

- [ ] **Step 4: Run the fixtures suites green**

Run: `bunx vitest run tests/fixtures-all-lenient.test.ts tests/fixtures-all-strict.test.ts`
Expected: all PASS.

- [ ] **Step 5: Commit D**

```bash
git add tests/fixtures-all.manifest.json
git commit -m "test: regenerate fixtures manifest for inventory-fingerprint gate split

New inventory collectors feed the repair gate. Content losses are
repair-content-loss (error); mark/formatting/bookkeeping losses are
repair-fidelity-loss (warn). Manifest regenerated; error-code changes are
confined to fixtures where repair drops genuine content. word-probe fields
preserved from the LibreOffice probe.

Co-Authored-By: Claude Opus 4.7 (1M context) <noreply@anthropic.com>"
```

- [ ] **Step 6: Push the branch**

```bash
git push -u origin feat/inventory-fingerprint
```

---

## Final verification (run after all commits)

- [ ] `bunx tsc --noEmit` → exit 0
- [ ] `bunx vitest run tests/docx-diagnostics.test.ts tests/docx-inventory-diff.test.ts tests/diff-docx.cli.test.ts` → all PASS
- [ ] `bunx vitest run tests/fixtures-all-lenient.test.ts tests/fixtures-all-strict.test.ts` → all PASS
- [ ] `bun run fmt:fix` then `git diff --stat` → confirm no unrelated files were reformatted (revert any that were, per the repo's focused-commit convention)

## Notes / known limitations (from the spec)
- Aggregate histogram cannot distinguish reshape from delete+add, nor reshape-up from reshape-down, when multiple same-shape elements exist. Genuine content loss is still caught via the Content-class element counts.
- Families 3 & 4 collect only under `strict`; `word-valid` behaves like `lenient`.
- `diff-docx` exits non-zero iff a Content-class decrease exists.
