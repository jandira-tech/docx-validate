# DOCX semantic fingerprint & symmetric inventory diff

**Date:** 2026-05-26
**Branch:** `feat/inventory-fingerprint` (stacked on `feat/second-pass-diff-analysis` / PR #5)
**Status:** Design approved — pending implementation plan

## Problem

Two `.docx` files can differ in ways that do **not** change the rendered
document (whitespace, element reordering, format-equivalent rewrites), and in
ways that **do** (a table shrinks, a page break disappears). The existing
`compareDocxSemanticInventories` in `docx-diagnostics.ts` is **one-directional
and decrease-only** — it answers "did _my repair step_ destroy content?" and
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
- **No disambiguation between a reshape and a delete+add**, nor between a
  reshape-down and reshape-up, when multiple same-shape elements exist. With
  three `3×4` tables, resizing one to `3×2` and deleting one then adding a `3×2`
  produce the same histogram delta. This is an accepted limitation of the
  identity-free model; genuine content loss is still caught via the Content-class
  counts (see Severity policy), which move unambiguously with element count.
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

| File                                                   | Change | Responsibility                                                                                                                                                                                                                                                                                                                                                                             |
| ------------------------------------------------------ | ------ | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ |
| `src/scripts/office/validators/docx-diagnostics.ts`    | extend | Add the four new collector families and a `profile` arg on `collectDocxSemanticInventory`. Add `severityClassFor`. `compareDocxSemanticInventories` keeps its decrease-only logic but gains a `severityClassFor` dispatch that emits two codes (`repair-content-loss` / `repair-fidelity-loss`) instead of one. `buildRepairPlanIssues` is unchanged (it already handles any code string). |
| `src/scripts/office/validators/docx-inventory-diff.ts` | new    | `diffDocxInventories(before, after)`, the `DocxInventoryDiff` type, `severityFor`, and the two formatters.                                                                                                                                                                                                                                                                                 |
| `scripts/diff-docx.ts`                                 | new    | CLI: `diff-docx <a> <b> [--profile lenient\|strict\|word-valid]`. Unpacks packed inputs, collects both inventories, prints the markdown report.                                                                                                                                                                                                                                            |
| `src/scripts/office/validate.ts`                       | edit   | Thread the active `profile` into both `collectDocxSemanticInventory` calls in the repair path (currently called with no profile).                                                                                                                                                                                                                                                          |

### Profile threading

`collectDocxSemanticInventory(unpackedDir, profile: Profile = "lenient")` gains
an optional `profile` argument (`Profile = "lenient" | "strict" | "word-valid"`,
from `src/lib/types.ts`). Families 3 and 4 are collected **only when
`profile === "strict"`**. Lenient and word-valid omit them entirely so they
cannot generate noise.

`validate.ts` passes its active `profile` into both `collectDocxSemanticInventory`
calls in the repair path so strict repairs gate on the strict-only families too.

`word-valid` behaves identically to `lenient` for inventory collection (families
3+4 omitted) — it is a Word-openability profile, not a spec-purist one. The
`diff-docx` CLI accepts all three profile values for symmetry with the validator,
but only `strict` changes inventory output.

## New coverage detail

### Family 1 — in-run atomic marks (always; quantity)

Counters under category `"inline mark"`, unit `"occurrence(s)"`:

- `w:br` split by `w:type` attribute → `line break` (no type / `textWrapping`;
  per ISO 29500-4 §17.3.3.1 the default is `textWrapping`), `page break`
  (`page`), `column break` (`column`). `w:br` with `w:type='separator'` or
  `'continuationSeparator'` (footnote/endnote separators) is **excluded** to
  avoid noise.
- `w:tab` → `tab`; `w:sym` → `symbol`; `w:cr` → `carriage return`;
  `w:softHyphen` → `soft hyphen`; `w:noBreakHyphen` → `non-breaking hyphen`.

Example: `"inline mark | page break | occurrence(s) | 3"`.

### Family 2 — table shape (always; count + shape)

Per `w:tbl`, category `"table shape"`:

- Rows = count of direct `w:tr` children. Columns = count of `w:gridCol` in the
  table's `w:tblGrid`. **Fallback when `w:tblGrid` is absent or empty** (common
  in tool-generated/malformed tables): columns = `max(w:gridSpan/@w:val)` summed
  per `w:tc` across the first `w:tr`, defaulting to `count(w:tc)` in the first
  row. This prevents silently reporting `0` columns. Label `table {rows}×{cols}`,
  unit `table(s)`.
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

Per `wp:extent` (namespace
`http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing`),
category `"image shape"`:

- `cx`/`cy` are **attributes** (EMU), **rounded to the nearest 1000 EMU**
  (≈0.1 mm) to absorb the re-rounding differences common between Word,
  LibreOffice, and Google Docs.
- Wrap = `inline` when `wp:extent`'s **parent** is `wp:inline`, `anchor` when its
  parent is `wp:anchor`. `wp:extent` is always a _direct child_ of those two
  elements — match on parent, not ancestor.
- Collect **only `wp:extent`** — explicitly **not** `wp:effectExtent` (a sibling
  that also carries `cx`/`cy`) and **not** `a:extent` (DrawingML-main, a child of
  `a:xfrm` for group-shape transforms).
- Label `image ~{cx}×{cy} {wrap}`, unit `image(s)`.

## Diff data model & semantics

```ts
export interface DocxInventoryDelta {
    key: string; // path\0category\0label\0unit
    path: string; // always present (every counter carries its part path)
    category: string;
    label: string;
    unit: string;
    before: number; // 0 for added
    after: number; // 0 for removed
}

export interface DocxInventoryDiff {
    added: DocxInventoryDelta[]; // key only in B (before 0)
    removed: DocxInventoryDelta[]; // key only in A (after 0)
    changed: DocxInventoryDelta[]; // same key, before !== after
    unchangedCount: number; // keys identical in both
}

export function diffDocxInventories(before: DocxSemanticInventory, after: DocxSemanticInventory): DocxInventoryDiff;
```

- **Symmetric**: a key in only B → `added`; only A → `removed`; in both with a
  different count → `changed`; identical → counted in `unchangedCount`.
- A **shape change** such as `table 3×4 → 3×2` is `removed("table 3×4") +
added("table 3×2")` (different keys). The structured `changed[]` is reserved
  strictly for same-key count deltas — no fragile 1:1 instance pairing (ambiguous
  with multiple same-shape elements, per Non-goals). Pairing a removed shape-key
  with an added shape-key for human readability is a **formatter** concern (see
  the "Reshaped" subsection below), not a data-model guarantee.
- Deterministic ordering: all arrays sorted by `(path, category, label)`.

## Severity policy

Every diff entry has a **direction**: a `removed` delta, or a `changed` delta
with `after < before`, is a **decrease**; an `added` delta, or a `changed` delta
with `after > before`, is an **increase**. The histogram never emits a standalone
"reshape" event — a shape change (`table 3×4 → 3×2`) decomposes into a _decrease_
of `table 3×4` plus an _increase_ of `table 3×2`. There are therefore only two
directions to tier.

Severity = `severityFor(categoryClass, direction)`, where
`severityClassFor(counter)` maps a counter's `category` to one class. Both
helpers are shared by the fingerprint and the repair gate so the policy lives in
one place.

**Principle:** 🔴 error = a **content-bearing element was destroyed** (a decrease
in the Content class); 🟠 warn = layout / formatting / metadata fidelity shifted,
or content was added; ⚪ info = a pure non-rendering addition.

| Category-class                       | Members                                                                                                                                                                                                                                         | Decrease (removed / count↓) | Increase (added / count↑) |
| ------------------------------------ | ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- | --------------------------- | ------------------------- |
| **Content**                          | text chars, paragraph, run-with-text, **table/row/cell count**, footnote, endnote, comment entry, comment marker, bookmark, numbering ref, math, drawing/picture, tracked change (ins/del), referenced part/relationship, content-type override | 🔴 error                    | 🟠 warn                   |
| **Table shape**                      | `table R×C`, merged-cell shape                                                                                                                                                                                                                  | 🟠 warn                     | 🟠 warn                   |
| **Section geometry** _(strict only)_ | page size, orientation, margins, columns                                                                                                                                                                                                        | 🟠 warn                     | 🟠 warn                   |
| **Image shape** _(strict only)_      | extent (bucketed), wrap                                                                                                                                                                                                                         | 🟠 warn                     | 🟠 warn                   |
| **Formatting**                       | bold/italic/underline/strike/caps/hidden/color/highlight/size/vertAlign, style ref, style def, style formatting                                                                                                                                 | 🟠 warn                     | ⚪ info                   |
| **Atomic marks**                     | line/page/column break, tab, symbol, cr, soft/no-break hyphen                                                                                                                                                                                   | 🟠 warn                     | ⚪ info                   |
| **Bookkeeping**                      | package-asset bytes & part-exists, relationship target-mode/type, comment-thread aux (Extended/Ids/Extensible)                                                                                                                                  | 🟠 warn                     | ⚪ info                   |

**Why shape categories are warn, not error (and why error still catches real
loss):** a table that _grows_ `3×2 → 3×4` and a table that is _deleted then
re-added at a new shape_ are indistinguishable in the identity-free histogram —
both produce a decrease of one shape key and an increase of another. Tiering
shape **decreases** as error would therefore raise **false errors on benign
growth**. Genuine table loss is still an **error**, caught by the **Content**
class: the `table` / `table row` / `table cell` _counts_ decrease unambiguously
in the direction of loss whenever a table or cells are actually removed,
independent of shape. Section and image geometry have no Content counterpart
(they are layout, not content), so they are warn in both directions; being
strict-only, they surface as warnings only when the caller asked for spec-purist
scrutiny.

There is intentionally **no "reshape" tier** — the previous draft's third column
was fictional, since the data model produces only decreases and increases.

## Formatters

- `inventoryDiffToIssues(diff): ValidationIssue[]` — **severity-graded** per the
  matrix above, each delta tiered by `severityFor(class, direction)`. Codes:
  `inventory-content-loss` (Content decrease → error), `inventory-content-added`
  (Content increase → warn), `inventory-shape-change` (table/section/image, both
  directions → warn), `inventory-formatting-drift`, `inventory-mark-drift`,
  `inventory-bookkeeping-drift` (warn on decrease, info on increase).
  Descriptive — emitted by the standalone fingerprint, not the validator
  pass/fail path.
- `formatInventoryDiffMarkdown(diff): string` — top-level sections **Added /
  Removed / Changed**, each grouped by part `path` then `category`, every line
  prefixed with its severity (🔴/🟠/⚪). Plus a **Reshaped** subsection: a
  display-only pairing that, within a single `(path, category)` of a _shape_
  class, lists a removed shape-key alongside an added shape-key (e.g.
  `table 3×4 → 3×2`) for readability. This pairing is best-effort presentation
  and carries no identity guarantee (see Non-goals). Ends with a one-line summary
  (`N added, M removed, K changed, U unchanged; E error / W warn`).
- **CLI exit code:** `diff-docx` exits non-zero iff the diff contains any
  **error-tier** difference — i.e. any Content-class decrease — usable in CI as
  "fail if B lost content vs A."

## Repair gate interaction (decided: feed the gate)

Because the new collectors live in the shared `collectDocxSemanticInventory`,
the repair pipeline's before/after snapshots now include the richer counters.
`compareDocxSemanticInventories` stays decrease-only but splits its loss code by
severity class (via the shared `severityClassFor`):

- **Content-class decrease** → `repair-content-loss` (**error**) — unchanged
  behavior for genuine content (text, paragraph, table/row/cell count, footnote,
  comment, …).
- **Every other class decrease** (table shape, section, image, formatting,
  atomic marks, bookkeeping) → `repair-fidelity-loss` (**warn**) — new,
  non-blocking.

**Control flow:** the function iterates all `before` counters, and for each
counter whose `after` count is lower, dispatches via `severityClassFor` to the
appropriate code+severity and accumulates it. It returns the full accumulated
list. The existing early `repair-content-preserved` (info) is returned **only
when that list is empty** (no decreases of any class).

**Consequence for the manifest:** because the _new_ collector families
(atomic marks, formatting nuance, bookkeeping) map to **warn**, they add
`repair-fidelity-loss` (a warning), **not** new error codes. So `fixtures-all`
entries only flip to _failing_ when a repair drops genuine **content** (text,
table row/col, footnote, …) — which is exactly what should fail. Expect
warning-code additions across many entries and error-code changes on only a few.
Regenerate validator-side codes with `bunx tsx scripts/update-manifest.ts`;
preserve the LibreOffice **word-probe** fields (regenerate on a Word-equipped
machine via `SOFFICE_AVAILABLE=1 bun run test:fixtures:word`, or patch only the
changed code arrays to avoid clobbering word data).

**Byte-noise guard:** package-asset _bytes_ appear in the fingerprint (warn),
but repair only mutates unpacked XML and never rewrites media bytes, so the
decrease-only gate will not emit spurious `repair-fidelity-loss` on byte counts.
If that ever proves false, exclude the `part bytes` counter from the gate while
keeping it in the fingerprint.

## Testing (TDD red-green)

One vitest spec per concern, fixtures as inline `wrapDocument(...)` strings like
`tests/validators-docx.test.ts`:

1. **Collectors** (`tests/docx-diagnostics.test.ts`, extend): each family.
   Families 3+4 assert **absent under lenient, present under strict, and absent
   under word-valid** (word-valid behaves like lenient for all four families).
2. **Diff** (`tests/docx-inventory-diff.test.ts`, new): added / removed /
   changed / unchanged; a shape change produces remove+add of different keys;
   empty-vs-empty; identical-vs-identical → all unchanged; a `changed` delta with
   `after > before` is classified **increase**, with `after < before`
   **decrease**.
3. **Severity policy** (same file): `severityFor(class, direction)` —
   Content decrease → error; Content increase → warn; table/section/image (either
   direction) → warn; formatting/mark/bookkeeping decrease → warn, increase →
   info. **Critical regression test:** a table _growth_ `3×2 → 3×4` yields only
   warn-tier deltas (shape) plus a Content **increase** (more cells → warn) and
   **no error** — proving growth does not false-positive.
4. **Formatters** (same file): `inventoryDiffToIssues` emits the correct
   code+severity per delta (Content decrease → `inventory-content-loss` error;
   shape → `inventory-shape-change` warn; etc.); markdown grouping, the Reshaped
   pairing, severity prefixes, and summary line.
5. **CLI** (`tests/diff-docx.cli.test.ts`, new): exit **0** when only
   warn/info-tier diffs (e.g. a reshaped/grown table); exit **non-zero** when a
   Content-class decrease exists (e.g. a removed paragraph); expected markdown
   sections present.
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
