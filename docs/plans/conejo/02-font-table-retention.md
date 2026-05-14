# Plan 02: Font Table Retention

## Problem

The repairer replaces the full `word/fontTable.xml` with a near-empty document
containing only the XML declaration and an empty `<w:fonts>` root element:

```xml
<?xml version="1.0" encoding="UTF-8"?>
<w:fonts xmlns:w="..." xmlns:mc="..." mc:Ignorable="...">
</w:fonts>
```

The working (Word-repaired) copy contains 5 font entries with full metadata
(`w:panose1`, `w:charset`, `w:family`, `w:pitch`, `w:sig`, `w:altName`):
`Inter`, `Times New Roman`, `Courier New`, `Aptos Display`, `Aptos`.

An empty font table causes Word to substitute all fonts with its default
(typically Calibri), silently changing document appearance. It does not
trigger the "unreadable content" warning on its own, but is a fidelity loss
that causes downstream rendering differences.

## Current detection

No validator check for empty `<w:fonts>`. The XSD schema
(`ISO-IEC29500-4_2016/wml.xsd`) allows an empty `<w:fonts>` element.

## Proposed fix

### Repairer side

**Merge strategy** (preferred over copy-through):

1. Parse the input `word/fontTable.xml` and extract all `<w:font>` children.
2. Parse the output `word/fontTable.xml` (which the repairer may have
   generated with its own minimal set).
3. For each font in the input that is not already present in the output
   (match by `w:name` attribute), append it to the output font table.
4. Write the merged font table back to `word/fontTable.xml`.

This ensures repairer-injected fonts survive, and user fonts from the
original are not silently lost.

**Namespace handling**: preserve the `mc:Ignorable` and `xmlns:*` declarations
from the input root element rather than generating a minimal set.

### Validator side (new check, severity: warning)

Add `validateFontTable()` to `DOCXSchemaValidator`
(`src/scripts/office/validators/docx.ts`):

1. Read `word/fontTable.xml` (skip if absent — font table is optional).
2. Count `<w:font>` children of `<w:fonts>`.
3. If count is 0, emit `font-table-empty` at `"warning"` severity — an empty
   font table is valid XML and Word handles it, but it almost always indicates
   a repairer bug.

Under `profile: "strict"` this should remain `"warning"` (XSD-valid).
Under `profile: "word-valid"` keep as `"warning"` (Word opens the file).

## Acceptance criteria

- After a repairer round-trip on the second-pass fixture, `word/fontTable.xml`
  contains at least the fonts present in the input (`Inter`, `Times New Roman`,
  `Courier New`).
- New unit test in `tests/validators-docx.test.ts` under
  `describe("validateFontTable")` that creates an empty `<w:fonts>` and
  asserts `font-table-empty` warning is issued.
- Existing fixture manifest (`tests/fixtures-all.manifest.json`) updated if
  any existing fixtures trigger the new warning.

## Test fixture reference

- Working font table: `tests/fixtures/word-strict/second-pass/unpacked-working/word/fontTable.xml`
- Broken font table: `tests/fixtures/word-strict/second-pass/unpacked-broken/word/fontTable.xml`
- Diff: `tests/fixtures/word-strict/second-pass/diffs/word_fontTable.xml.diff`