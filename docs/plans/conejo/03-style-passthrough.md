# Plan 03: Style Passthrough

## Problem

The repairer regenerates `word/styles.xml` from a minimal built-in set
instead of preserving the input styles. The diff shows:

1. `<w:latentStyles>` block is completely removed (150+ `<w:lsdException>`
   entries gone).
2. Comment-related paragraph/character styles (`CommentText`, `CommentSubject`,
   `CommentReference`) are removed.
3. Heading styles (`Heading1`–`Heading6`) are renamed with capitalized display
   names and stripped of `uiPriority` and outline-level paragraph properties.
4. Custom styles present in the input are dropped entirely.
5. The minimal output set is: `Normal`, `DefaultParagraphFont`, `TableNormal`,
   `NoList` (the four ECMA-376 §17.7.4.4 defaults), plus a reduced subset of
   heading and footnote styles.

This causes:
- Comment-related styles missing → `style-default-missing` error (currently
  detected by the validator — `REQUIRED_DEFAULT_STYLES` in `docx.ts`).
- Custom styles lost → `style-ref-missing` warnings for any paragraph/run
  that references a dropped style.
- Latent-style data lost → cosmetic but irreversible fidelity loss.

## Current detection

`validateStyles()` in `DOCXSchemaValidator` (`src/scripts/office/validators/docx.ts`)
currently checks:

- `REQUIRED_DEFAULT_STYLES` (Normal, DefaultParagraphFont, TableNormal, NoList)
  via the `style-default-missing` error code.
- Referenced but undefined styles via `style-ref-missing` warning.

The `WELL_KNOWN_STYLE_DEFINITIONS` constant (lines 118–188 of `docx.ts`)
defines fallback XML for CommentReference, CommentText, CommentSubject,
Header, Footer, DefaultParagraphFont, Normal, TableNormal, NoList — these are
injected by `repairMissingStyleDefinitions()` when missing.

## Proposed fix

### Repairer side

**Preserve-and-patch strategy**:

1. Copy `word/styles.xml` from the input verbatim.
2. Run the existing `repairMissingStyleDefinitions()` pass (which already
   handles the four required defaults and the comment styles).
3. Do **not** regenerate the styles from scratch or remove the latent styles
   block.

If the repairer needs to add styles (e.g., inject a house style), use an
append-only approach: parse the existing `<w:styles>`, check whether the
target `w:styleId` already exists, and only insert if absent.

### Validator side (enhancement to existing checks)

Extend `validateStyles()` to also check:

1. `latent-styles-missing` (info): `<w:styles>` root has no `<w:latentStyles>`
   child. This is XSD-valid but unusual for Word-generated files; flag at
   `"info"` severity so it is visible without blocking the valid flag.

2. Ensure `CommentReference`, `CommentText`, `CommentSubject` are included in
   `REQUIRED_DEFAULT_STYLES` (they currently appear in `WELL_KNOWN_STYLE_DEFINITIONS`
   for repair but are not checked in the required-defaults validation pass).
   Add them to the validation check for `profile: "strict"` at minimum.

**Code location**: `DOCXSchemaValidator.validateStyles()` in
`src/scripts/office/validators/docx.ts`. The `REQUIRED_DEFAULT_STYLES`
constant at line 197 and the `repairMissingStyleDefinitions` method body are
the relevant implementation sites.

## Acceptance criteria

- After a repairer round-trip, `word/styles.xml` retains `<w:latentStyles>`,
  all custom styles from the input, and unmodified heading definitions.
- `validate(output, { profile: "strict" })` returns no `style-default-missing`
  or `style-ref-missing` errors.
- New unit test asserting `latent-styles-missing` info issue appears when
  `<w:latentStyles>` is absent.
- Existing fixture manifest updated for any existing fixtures that now emit the
  new `latent-styles-missing` info.

## Test fixture reference

- Working styles: `tests/fixtures/word-strict/second-pass/unpacked-working/word/styles.xml`
- Broken styles: `tests/fixtures/word-strict/second-pass/unpacked-broken/word/styles.xml`
- Diff: `tests/fixtures/word-strict/second-pass/diffs/word_styles.xml.diff`