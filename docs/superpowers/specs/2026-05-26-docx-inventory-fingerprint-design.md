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
- No new severity model for the repair pipeline. `compareDocxSemanticInventories`
  keeps its exact decrease-only logic and `repair-content-loss` **error** code.
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

## Formatters

- `inventoryDiffToIssues(diff): ValidationIssue[]` — **info severity**, codes
  `inventory-added` / `inventory-removed` / `inventory-changed`. Descriptive
  only; never fails validation.
- `formatInventoryDiffMarkdown(diff): string` — sections Added / Removed /
  Changed, grouped by part `path`, and within each part grouped by `category` so
  a reshape's remove+add render adjacently under the same category heading.
  Includes a one-line summary (`N added, M removed, K changed, U unchanged`).

## Repair gate interaction (decided: feed the gate)

Because the new collectors live in the shared `collectDocxSemanticInventory`,
the repair pipeline's before/after snapshots now include the richer counters.
`compareDocxSemanticInventories` stays decrease-only, so the **only** new effect
is: a repair step that drops a newly-counted element (a line break, a table
row/col, a merged cell, and under strict a section/image shape) is now caught as
`repair-content-loss`. This is strictly more protection.

**Consequence:** some `fixtures-all` manifest entries will gain
`repair-content-loss` (and `repair-plan*`) codes and must be regenerated /
reviewed. The manifest's validator-side `errorCodes` can be regenerated with
`bunx tsx scripts/update-manifest.ts`; the LibreOffice **word-probe** fields
must be preserved (regenerate on a Word-equipped machine via
`SOFFICE_AVAILABLE=1 bun run test:fixtures:word`, per the existing workflow, or
patch only the affected `errorCodes` to avoid clobbering word data).

## Testing (TDD red-green)

One vitest spec per concern, fixtures as inline `wrapDocument(...)` strings like
`tests/validators-docx.test.ts`:

1. **Collectors** (`tests/docx-diagnostics.test.ts`, extend): each family.
   Families 3+4 assert **absent under lenient, present under strict**.
2. **Diff** (`tests/docx-inventory-diff.test.ts`, new): added / removed /
   changed / unchanged; reshape produces remove+add; empty-vs-empty;
   identical-vs-identical → all unchanged.
3. **Formatters** (same file): `inventoryDiffToIssues` severities + codes;
   markdown grouping and summary line.
4. **CLI** (`tests/diff-docx.cli.test.ts`, new): smoke test over two tiny
   unpacked fixtures → exit 0, expected markdown sections.
5. **Repair-gate regression**: a fixture where repair drops a line break now
   yields `repair-content-loss` (proves the gate sees the new counters).

## Stacked-PR / delivery

- Branch `feat/inventory-fingerprint` off `feat/second-pass-diff-analysis`.
- PR targets the PR #5 branch (not `main`) so it stacks; nothing merged.
- Commit order: (a) collectors + their tests; (b) diff + formatters + tests;
  (c) CLI + test; (d) manifest regeneration/review as its own commit with the
  fixture-delta explained.
