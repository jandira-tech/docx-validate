# Plan 07: Cell Border Normalization

## Problem

The repairer expands implicit table cell borders to explicit `<w:tcBorders>`
blocks on every `<w:tcPr>` (table cell properties), even when the cell's
borders are inherited from the table style and the explicit values are
identical to what the style would provide.

The diff shows every cell that previously had no explicit `<w:tcBorders>` now
carries a fully-expanded set:

```xml
<w:tcBorders>
  <w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/>
  <w:left w:val="single" w:sz="4" w:space="0" w:color="auto"/>
  <w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/>
  <w:right w:val="single" w:sz="4" w:space="0" w:color="auto"/>
  <w:insideH w:val="single" w:sz="4" w:space="0" w:color="auto"/>
  <w:insideV w:val="single" w:sz="4" w:space="0" w:color="auto"/>
</w:tcBorders>
```

This is semantically redundant when the values match the table style's
defaults, but it overrides style-level conditional formatting (e.g., the
`<w:tblLook>` banded-row border suppression). Adding these explicit borders
prevents Word from applying the style's conditional formatting, changing
the visual output.

## Current detection

No validator check. Explicit `<w:tcBorders>` is XSD-valid and commonly used.
The interaction with conditional formatting suppression is a semantic/rendering
concern, not a schema violation.

## Proposed fix

### Repairer side

**Do not inject explicit borders when the values match table-style defaults.**

Implementation approach:

1. Before writing `<w:tcBorders>` to a cell, resolve the table's effective
   border values from:
   a. The table's `<w:tblBorders>` (table-level defaults).
   b. The table style's `<w:tblBorders>` in `word/styles.xml`.
   c. The table's `<w:tblLook>` conditional formatting applicability.

2. If the would-be explicit borders exactly match the resolved effective
   borders, omit `<w:tcBorders>` from the output `<w:tcPr>`.

3. If explicit cell border customization is genuinely needed (e.g., the
   repairer is adding a border that was not there), write only the
   border sides that differ from the table-level default.

**Simpler alternative**: If the full border-resolution logic is too complex,
adopt a conservative strategy: if the input `<w:tcPr>` had no `<w:tcBorders>`,
the output should also have no `<w:tcBorders>`. Only retain/add explicit
borders if the input had them.

### Validator side (new check, severity: info)

Add `validateRedundantCellBorders()` to `DOCXSchemaValidator`:

1. For each table cell in `word/document.xml`:
   a. Resolve the table's border defaults (from `<w:tblBorders>` and style).
   b. Check if the cell's `<w:tcBorders>` exactly duplicates those defaults.
   c. If so, emit `cell-borders-redundant` at `"info"` severity.

This is an `"info"` check (not `"warning"`) because redundant borders are
not a correctness issue, only a fidelity/intent issue.

**Implementation note**: Full style resolution requires parsing `word/styles.xml`
in addition to `word/document.xml`. The `DOCXSchemaValidator` constructor
already accepts the `unpackedDir` path — use that to read `styles.xml` within
the validator method.

## Acceptance criteria

- After repairer round-trip on the second-pass fixture, cells that had no
  explicit borders in the input have no `<w:tcBorders>` in the output.
- New unit test in `tests/validators-docx.test.ts` under
  `describe("validateRedundantCellBorders")` that constructs a table where
  cell borders duplicate table-level defaults and asserts `cell-borders-redundant`
  info is emitted.

## Test fixture reference

- Document diff: `tests/fixtures/word-strict/second-pass/diffs/word_document.xml.diff`
  (search for `tcBorders` to find the specific expansion locations)