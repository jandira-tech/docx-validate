# Plan 05: tblLook Preservation

## Problem

The repairer removes `<w:tblLook>` elements from table properties (`<w:tblPr>`)
in `word/document.xml`. The `<w:tblLook>` element controls which conditional
formatting bands (first row, last row, first column, last column, banded rows,
banded columns) are applied from the table style.

When `<w:tblLook>` is absent, Word applies default banding, which may
differ from the original document's intent. This is a fidelity issue rather
than a compatibility error — Word opens the file without warning — but it
causes visible rendering differences in styled tables.

The diff also shows the repairer removes:
- `w:tblCellSpacing` (cell spacing within tables)
- `w:jc` alignment on tables (table-level justification)
- `w:tblInd` (table indent from margin)
- `w:shd` on table/cell level (cell shading/background)

## Current detection

No validator check. These are XSD-valid omissions.

## Proposed fix

### Repairer side

1. When copying `<w:tblPr>` from input to output, preserve all children
   except those the repairer explicitly intends to change (e.g., `<w:tblBorders>`
   if the border normalization fix from Plan 07 is applied).

2. Specifically **never remove** `<w:tblLook>` unless its value matches the
   OOXML default (`w:val="04A0"`, all bands on). Removing a `<w:tblLook>`
   that was explicitly set is always a fidelity regression.

3. Preserve `<w:tblCellSpacing>` and `<w:tblInd>` verbatim.

4. For cell shading (`<w:shd>`), preserve unless the repairer is performing
   an explicit shading normalization pass (which should be a separate,
   opt-in operation).

### Validator side (new check, severity: warning)

Add `validateTableLook()` to `DOCXSchemaValidator`:

1. Parse `word/document.xml`.
2. For each `<w:tbl>` that has a `<w:tblStyle>` reference (meaning it uses
   a named style with conditional formatting bands):
   - If `<w:tblPr>` contains no `<w:tblLook>`, emit `tbl-look-missing` at
     `"info"` severity.
3. Under `profile: "word-valid"`, keep as `"info"` (not Word-blocking).

**Code location**: new method `validateTableLook()` in `DOCXSchemaValidator`
in `src/scripts/office/validators/docx.ts`, called from `validate()`.

## Acceptance criteria

- After repairer round-trip, every `<w:tbl>` that had a `<w:tblLook>` in
  the input retains it in the output.
- New unit test in `tests/validators-docx.test.ts` under
  `describe("validateTableLook")` asserting `tbl-look-missing` info when
  a styled table lacks `<w:tblLook>`.

## Test fixture reference

- Document diff: `tests/fixtures/word-strict/second-pass/diffs/word_document.xml.diff`
  (search for `tblLook` in the diff to find specific change locations)