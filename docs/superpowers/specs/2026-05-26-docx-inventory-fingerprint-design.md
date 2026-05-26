# DOCX semantic fingerprint & symmetric inventory diff

**Date:** 2026-05-26
**Branch:** `feat/inventory-fingerprint` (stacked on `feat/second-pass-diff-analysis` / PR #5)
**Status:** Design approved — pending implementation plan

## Problem

Two `.docx` files can differ in ways that do **not** change the rendered
document (whitespace, element reordering, format-equivalent rewrites), and in
ways that **do** (a table shrinks, a page break disappears). The existing
`compareDocxSemanticInventories` in `docx-diagnostics.ts` is **one-directional
and decrease-only** — it answers "did *my repair step* destroy content?" and
emits `repair-content-loss` errors. It cannot describe how two arbitrary
documents differ, and its collectors miss several element classes entirely
(in-run atomic marks; per-table shape; section geometry; image extents).

We want a richer, **symmetric** fingerprint that enumerates more "stuff we
have" — either as a plain quantity (e.g. line breaks) or with shape (e.g. a
table's rows×cols) — so benign, non-rendering differences are explicitly
accounted for instead of being invisible or mistaken for loss.

## Goals

- A symmetric A-vs-B comparison reporting **additions, removals, and count
  deltas** between two document inventories.
- Expanded collector coverage:
  - **Family 1 — in-run atomic marks** (quantity). Always collected.
  - **Family 2 — table shape** (count + shape). Always collected.
  - **Family 3 — section/page geometry** (shape). Strict profile only.
  - **Family 4 — image/drawing shape** (shape, tolerance-bucketed). Strict profile only.
- Structured diff output + two formatters (`ValidationIssue[]`, markdown) + a
  thin CLI.
- The new collectors feed **both** the new fingerprint **and** the existing
  repair content-loss gate (deliberate — see "Repair gate" below).

## Non-goals (YAGNI)

- No per-instance element identity or move/reorder detection. Comparison is an
  **aggregate histogram**: shape is encoded in the counter key, so a reshape
  reads as remove(old-shape) + add(new-shape).
- No change to the repair pipeline's **decrease-only direction** or its call
  site. `compareDocxSemanticInventories` still fires only on decreases. Its
  output codes are refined by severity tier (see "Repair gate"), but it never
  becomes symmetric and is never replaced by the fingerprint.
- No auto-wiring of the fingerprint into `validate.ts` output. It is a
  standalone library export + CLI.

## Architecture

Aggregate-histogram model, reusing the existing `DocxSemanticCounter` /
`DocxSemanticInventory` shape (a `Map` keyed by `path + category + label + unit`).

### Module layout

| File | Change | Responsibility |
|------|--------|----------------|
| `src/scripts/office/validators/docx-diagnostics.ts` | extend | Add the four new collector families. Keep `compareDocxSemanticInventories` and `buildRepairPlanIssues` byte-for-byte unchanged. |
| `src/scripts/office/validators/docx-inventory-diff.ts` | new | `diffDocxInventories(before, after)`, the `DocxInventoryDiff` type, and the two formatters. |
| `scripts/diff-docx.ts` | new | CLI: `diff-docx <a> <b> [--profile strict\|lenient]`. Unpacks packed inputs, collects both inventories, prints the markdown report. |

### Profile threading

`collectDocxSemanticInventory(unpackedDir, profile: Profile = "lenient")` gains
an optional `profile` argument (`Profile = "lenient" | "strict" | "word-valid"`,
from `src/lib/types.ts`). Families 3 and 4 are collected **only when
`profile === "strict"`**. Lenient and word-valid omit them entirely so they
cannot generate noise.

`validate.ts` passes its active `profile` into both `collectDocxSemanticInventory`
calls in the repair path so strict repairs gate on the strict-only families too.

## New coverage detail

### Family 1 — in-run atomic marks (always; quantity)
Counters under category `"inline mark"`, unit `"occurrence(s)"`:
- `w:br` split by `w:type` attribute → `line break` (no type / `textWrapping`),
  `page break` (`page`), `column break` (`column`).
- `w:tab` → `tab`; `w:sym` → `symbol`; `w:cr` → `carriage return`;
  `w:softHyphen` → `soft hyphen`; `w:noBreakHyphen` → `non-breaking hyphen`.

Example: `"inline mark | page break | occurrence(s) | 3"`.

### Family 2 — table shape (always; count + shape)
Per `w:tbl`, category `"table shape"`:
- Columns = count of `w:gridCol` inside the table's `w:tblGrid`; rows = count of
  direct `w:tr` children → label `table {rows}×{cols}`, unit `table(s)`.
- Merged cells: `w:gridSpan w:val=N` → `merged cell gridSpan={N}`; `w:vMerge`
  (with `w:val="restart"` or continuation) → `merged cell vMerge`, unit
  `cell(s)`.

Examples: `"table shape | table 3×4 | table(s) | 2"`,
`"table shape | merged cell gridSpan=2 | cell(s) | 5"`.

### Family 3 — section/page geometry (strict only; shape)
Per `w:sectPr` (body-level and paragraph-level), category `"section geometry"`:
- `w:pgSz` → `section {orient} {w}×{h}` (orient defaults to `portrait`).
- `w:pgMar` → `section margins T{top} R{right} B{bottom} L{left}`.
- `w:cols w:num=N` → `section columns={N}` (default 1).

### Family 4 — image/drawing shape (strict only; shape, tolerance-bucketed)
Per `wp:extent`, category `"image shape"`:
- `cx`/`cy` **rounded to the nearest 1000 EMU** to absorb tool re-rounding
  noise; wrap = `inline` (`wp:inline` ancestor) or `anchor` (`wp:anchor`).
- Label `image ~{cx}×{cy} {wrap}`, unit `image(s)`.

## Diff data model & semantics

```ts
export interface DocxInventoryDelta {
  key: string;          // path\0category\0label\0unit
  path?: string;
  category: string;
  label: string;
  unit: string;
  before: number;       // 0 for added
  after: number;        // 0 for removed
}

export interface DocxInventoryDiff {
  added: DocxInventoryDelta[];     // key only in B (before 0)
  removed: DocxInventoryDelta[];   // key only in A (after 0)
  changed: DocxInventoryDelta[];   // same key, before !== after
  unchangedCount: number;          // keys identical in both
}

export function diffDocxInventories(
  before: DocxSemanticInventory,
  after: DocxSemanticInventory,
): DocxInventoryDiff;
```

- **Symmetric**: a key in only B → `added`; only A → `removed`; in both with a
  different count → `changed`; identical → counted in `unchangedCount`.
- A **table reshape** `3×4 → 3×2` is `removed("table 3×4") + added("table 3×2")`
  (different keys). The structured `changed[]` is reserved strictly for same-key
  count deltas — no fragile 1:1 instance pairing (ambiguous with multiple
  tables). Readability of reshapes is a **formatter** concern, not a data-model
  concern.
- Deterministic ordering: all arrays sorted by `(path, category, label)`.

## Severity policy

Severity is a function of the difference's **category-class** and its
**direction** (loss / reshape / gain). Principle: **error = the rendered
document changed or content was destroyed; warn = appearance/layout/metadata
fidelity shifted but content survived; info = a pure, non-rendering addition.**

| Category-class | Members | Loss (removed / count↓) | Reshape (shape key changed) | Gain (added / count↑) |
|---|---|---|---|---|
| **Content** | text chars, paragraph, run-with-text, table/row/cell *count*, footnote, endnote, comment entry, comment marker, bookmark, numbering ref, math, drawing/picture, tracked change (ins/del), referenced part/relationship, content-type override | 🔴 error | — | 🟠 warn |
| **Table shape** | `table R×C` | 🔴 error | 🔴 error | 🟠 warn |
| **Section geometry** *(strict only)* | page size, orientation, margins, columns | 🔴 error | 🔴 error | 🔴 error |
| **Image shape** *(strict only)* | extent (bucketed), wrap | 🔴 error | 🔴 error | 🟠 warn |
| **Formatting** | bold/italic/underline/strike/caps/hidden/color/highlight/size/vertAlign, style ref, style def, style formatting | 🟠 warn | 🟠 warn | ⚪ info |
| **Atomic marks** | line/page/column break, tab, symbol, cr, soft/no-break hyphen | 🟠 warn | — | ⚪ info |
| **Bookkeeping** | package-asset bytes & part-exists, relationship target-mode/type, comment-thread aux (Extended/Ids/Extensible) | 🟠 warn | 🟠 warn | ⚪ info |

A counter's category-class is assigned by a single `severityClassFor(counter)`
helper (keyed off `category`) shared by both consumers, so the policy lives in
one place.

## Formatters

- `inventoryDiffToIssues(diff): ValidationIssue[]` — **severity-graded** per the
  matrix above. Codes: `inventory-loss` (error), `inventory-shape-change`
  (error or warn per table), `inventory-formatting-drift` (warn),
  `inventory-mark-drift` (warn), `inventory-bookkeeping-drift` (warn),
  `inventory-added` (warn for content, info otherwise). Descriptive — these are
  emitted by the standalone fingerprint, not the validator pass/fail path.
- `formatInventoryDiffMarkdown(diff): string` — sections Added / Removed /
  Changed, grouped by part `path`, and within each part grouped by `category` so
  a reshape's remove+add render adjacently under the same category heading. Each
  line is prefixed with its severity (🔴/🟠/⚪). Includes a one-line summary
  (`N added, M removed, K changed, U unchanged; E error / W warn`).
- **CLI exit code:** `diff-docx` exits non-zero iff the diff contains any
  **error-tier** difference — usable in CI as "fail if the repaired doc lost
  content vs the original."

## Repair gate interaction (decided: feed the gate)

Because the new collectors live in the shared `collectDocxSemanticInventory`,
the repair pipeline's before/after snapshots now include the richer counters.
`compareDocxSemanticInventories` stays decrease-only but splits its loss code by
severity tier (via the shared `severityClassFor`):

- **Content / Table-shrink / strict Section+Image loss** → `repair-content-loss`
  (**error**) — unchanged behavior for genuine content.
- **Formatting / Atomic-mark / Bookkeeping loss** → `repair-fidelity-loss`
  (**warn**) — new, non-blocking.

**Consequence for the manifest:** because the *new* collector families
(atomic marks, formatting nuance, bookkeeping) map to **warn**, they add
`repair-fidelity-loss` (a warning), **not** new error codes. So `fixtures-all`
entries only flip to *failing* when a repair drops genuine **content** (text,
table row/col, footnote, …) — which is exactly what should fail. Expect
warning-code additions across many entries and error-code changes on only a few.
Regenerate validator-side codes with `bunx tsx scripts/update-manifest.ts`;
preserve the LibreOffice **word-probe** fields (regenerate on a Word-equipped
machine via `SOFFICE_AVAILABLE=1 bun run test:fixtures:word`, or patch only the
changed code arrays to avoid clobbering word data).

**Byte-noise guard:** package-asset *bytes* appear in the fingerprint (warn),
but repair only mutates unpacked XML and never rewrites media bytes, so the
decrease-only gate will not emit spurious `repair-fidelity-loss` on byte counts.
If that ever proves false, exclude the `part bytes` counter from the gate while
keeping it in the fingerprint.

## Testing (TDD red-green)

One vitest spec per concern, fixtures as inline `wrapDocument(...)` strings like
`tests/validators-docx.test.ts`:

1. **Collectors** (`tests/docx-diagnostics.test.ts`, extend): each family.
   Families 3+4 assert **absent under lenient, present under strict**.
2. **Diff** (`tests/docx-inventory-diff.test.ts`, new): added / removed /
   changed / unchanged; reshape produces remove+add; empty-vs-empty;
   identical-vs-identical → all unchanged.
3. **Severity policy** (`tests/inventory-severity.test.ts` or within the diff
   spec): `severityClassFor` maps each category to the right tier; a content
   loss → error, a formatting loss → warn, a bookkeeping loss → warn, a content
   gain → warn, a mark gain → info.
4. **Formatters** (same file): `inventoryDiffToIssues` emits the correct
   code+severity per tier (e.g. content loss → `inventory-loss` error;
   formatting drift → warn; bookkeeping → warn); markdown grouping, severity
   prefixes, and summary line.
5. **CLI** (`tests/diff-docx.cli.test.ts`, new): exit **0** when only
   warn/info-tier diffs; exit **non-zero** when an error-tier diff exists
   (e.g. a removed paragraph); expected markdown sections present.
6. **Repair-gate regression**: a fixture where repair drops a **paragraph**
   still yields `repair-content-loss` (**error**); a fixture where repair drops
   a **line break** now yields `repair-fidelity-loss` (**warn**) and does **not**
   fail validation — proving the gate sees the new counters and tiers them
   correctly.

## Stacked-PR / delivery

- Branch `feat/inventory-fingerprint` off `feat/second-pass-diff-analysis`.
- PR targets the PR #5 branch (not `main`) so it stacks; nothing merged.
- Commit order: (a) collectors + their tests; (b) diff + formatters + tests;
  (c) CLI + test; (d) manifest regeneration/review as its own commit with the
  fixture-delta explained.
