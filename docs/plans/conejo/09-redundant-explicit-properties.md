# Plan 09: Redundant Explicit Run and Paragraph Properties

## Problem

The repairer injects explicit run properties (`<w:rPr>`) and paragraph
properties (`<w:pPr>`) on elements that previously inherited these values
from the paragraph/character style. The diff shows:

1. **Explicit font declarations on every run**: The repairer adds
   `<w:rFonts w:ascii="Inter" w:hAnsi="Inter" w:cs="Inter"/>` to runs in
   the header and footer even though the style (`Header`, `Footer`) already
   specifies `Inter` as the default font.

2. **Explicit line spacing**: `<w:spacing w:line="276" w:lineRule="auto"/>`
   added to paragraph properties where the style already implies the same
   spacing.

3. **Explicit color and size**: `<w:color w:val="000000"/>` and `<w:sz>`
   added to runs that inherit these from the paragraph style's `<w:rPr>`.

These redundant properties:
- Inflate file size (minor issue).
- Make the document harder to re-style: changing the style's font only
  affects paragraphs/runs that do NOT have the explicit override. Explicit
  overrides win over style definitions, so runs with redundant explicit
  properties stop responding to style changes.
- Interfere with Word's "Clear Formatting" behavior.

## Current detection

No validator check. Explicit run/paragraph properties that duplicate style
defaults are XSD-valid and produce correct rendering, so this is purely a
fidelity/maintainability concern.

## Proposed fix

### Repairer side

**Do not inject explicit formatting that duplicates the resolved style default.**

Implementation approach:

1. When processing `<w:r>` elements, resolve the effective run properties
   from the paragraph style chain:
   a. Default `<w:rPr>` in `<w:docDefaults>` in `word/styles.xml`.
   b. Paragraph style's `<w:rPr>`.
   c. Character style's `<w:rPr>` (if a `<w:rStyle>` is applied).

2. For each property the repairer would write explicitly, compare the value
   to the resolved style default. If they are identical, omit the explicit
   property.

3. Apply the same logic for paragraph properties (`<w:pPr>`) and the
   paragraph style chain.

**Simpler conservative strategy**: If the input `<w:r>` had no explicit
`<w:rFonts>`, do not add one. Only inject properties that the input explicitly
set or that are required for correctness (e.g., `<w:rStyle>` references for
comment annotations).

### Validator side (new check, severity: info)

Add `validateRedundantRunProperties()` to `DOCXSchemaValidator`:

1. Parse `word/styles.xml` to build the default run property chain.
2. For each `<w:r>` in `word/document.xml` that has an explicit `<w:rFonts>`:
   - If the font values match the document default (`<w:docDefaults>/<w:rPrDefault>`),
     emit `run-props-redundant` at `"info"` severity.

Limit to `<w:rFonts>` initially (highest signal-to-noise) before expanding to
other properties.

**Code location**: new method `validateRedundantRunProperties()` in
`DOCXSchemaValidator` in `src/scripts/office/validators/docx.ts`.

## Implementation note

Style resolution requires parsing both `word/styles.xml` and `word/document.xml`.
The computation is non-trivial for deeply-nested style chains. A pragmatic
scope is:

- Check only the document default run properties vs explicit `<w:rFonts>`.
- Do not recurse into paragraph-style → character-style chains for v1.

Expand scope in a follow-up once the basic check is in place.

## Acceptance criteria

- After repairer round-trip on the second-pass fixture, header and footer runs
  do not carry explicit `<w:rFonts>` when the `Header`/`Footer` style already
  specifies the same font.
- New unit test in `tests/validators-docx.test.ts` under
  `describe("validateRedundantRunProperties")` asserting `run-props-redundant`
  info is emitted for a run whose explicit font matches `<w:docDefaults>`.

## Test fixture reference

- Header diff: `tests/fixtures/word-strict/second-pass/diffs/word_header1.xml.diff`
  (shows `<w:rFonts>` injection)
- Footer diff: `tests/fixtures/word-strict/second-pass/diffs/word_footer1.xml.diff`
  (shows `<w:rFonts>` injection on every run)
- Document diff: `tests/fixtures/word-strict/second-pass/diffs/word_document.xml.diff`
  (bulk of explicit-property injections)