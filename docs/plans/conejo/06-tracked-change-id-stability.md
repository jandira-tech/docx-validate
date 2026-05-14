# Plan 06: Tracked-Change ID Stability

## Problem

The repairer regenerates `w:id` on tracked-change elements (`<w:ins>`,
`<w:del>`, `<w:rPrChange>`, `<w:pPrChange>`, `<w:tblPrChange>`,
`<w:trPrChange>`, `<w:tcPrChange>`) in `word/document.xml`. The original IDs
(which are stable document-level identifiers) are replaced with a new
monotonically-assigned sequence starting at 1.

Also affected:
- `w:id` on `<w:commentRangeStart>` and `<w:commentRangeEnd>` (handled in
  Plan 04, but ID stability applies equally here).
- `w16cid:durableId` on `<w:num>` elements in `word/numbering.xml`
  (the repairer removes `w16cid:durableId` from numbering definitions).

Changing tracked-change IDs:
1. Breaks external references (collaborative editing tools that track change
   IDs across sessions).
2. Can invalidate co-authoring session state stored in `word/settings.xml`
   (specifically `w:rsidRoot` and the rsid block).

The repairer also removes all `w:rsid*` attributes from paragraphs and runs,
which discards revision-session lineage data.

## Current detection

No validator check for tracked-change ID regeneration. The `w:id` values on
tracked-change elements are integers and are XSD-valid regardless of value.

The existing `repairParaIds()` method in `DOCXSchemaValidator`
(`src/scripts/office/validators/docx.ts`, lines 1800–1933) demonstrates
the cross-file remapping pattern needed for para IDs — a similar approach
applies to tracked-change IDs but is not currently implemented.

## Proposed fix

### Repairer side

1. **Preserve existing tracked-change IDs**: When copying `word/document.xml`,
   do not reassign `w:id` on `<w:ins>` and `<w:del>` elements.

2. **Collision avoidance**: If the repairer must inject a new tracked change
   (e.g., it is itself making a tracked insertion), allocate IDs above the
   maximum existing ID:
   ```typescript
   const maxExistingId = Math.max(
     0,
     ...insElements.map(el => parseInt(el.getAttribute("w:id") ?? "0", 10)),
     ...delElements.map(el => parseInt(el.getAttribute("w:id") ?? "0", 10)),
   );
   let nextId = maxExistingId + 1;
   ```

3. **Preserve `w:rsid*` attributes**: Do not strip revision session IDs from
   paragraphs and runs. These are informational but removing them silently
   changes the document's co-authoring fingerprint.

4. **Preserve `w16cid:durableId` on numbering**: If `word/numbering.xml`
   is regenerated, carry forward `w16cid:durableId` values from the input's
   `<w:num>` elements to the corresponding output `<w:num>` elements (matched
   by `w:abstractNumId` reference or `w:numId`).

### Validator side (new check, severity: warning under strict; info under lenient)

Add `validateTrackedChangeIds()` to `DOCXSchemaValidator`:

1. Collect all `w:id` values on `<w:ins>` and `<w:del>` elements across
   `word/document.xml`.
2. If all IDs are from a monotonic sequence starting at 1 (e.g., 1, 2, 3, …)
   with no gaps, emit `tracked-change-ids-regenerated` at `"info"` severity —
   this pattern is a strong signal of ID regeneration.

This is a heuristic check (cannot distinguish intentional assignment from
regeneration without the original file). Severity is kept at `"info"` to
avoid false positives on genuinely new documents.

**Code location**: new method `validateTrackedChangeIds()` in
`DOCXSchemaValidator` in `src/scripts/office/validators/docx.ts`.

## Acceptance criteria

- After repairer round-trip, tracked-change `w:id` values in the output
  match the input values (or are provably higher due to new injections).
- `w16cid:durableId` is present on `<w:num>` elements in the output if it
  was present in the input.
- `w:rsid*` attributes on `<w:p>` and `<w:r>` elements are retained.
- New unit test asserting `tracked-change-ids-regenerated` info for a
  sequential-from-1 ID set.

## Test fixture reference

- Numbering diff: `tests/fixtures/word-strict/second-pass/diffs/word_numbering.xml.diff`
  (shows `w16cid:durableId` removal from `<w:num>` elements)
- Document diff: `tests/fixtures/word-strict/second-pass/diffs/word_document.xml.diff`
  (shows `w:id` regeneration on tracked changes)