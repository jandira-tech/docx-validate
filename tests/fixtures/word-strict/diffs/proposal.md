# Proposal: Fix jubarte/docx-js-editor Writer Issues

## Summary

The broken document (`sample-document.broken-tables.docx`) is a snapshot of jubarte/docx-js-editor's writer output after several repair passes. Compared against the fully-repaired working document, it exhibits 16 changed files with differences spanning XML structure, table formatting, styles definitions, font tables, numbering, and relationship management. This proposal identifies what the writer added, what it removed, and proposes concrete fixes.

---

## (1) What Was ADDED by the Writer (in broken, not in working)

### A1. Files and content types added
| Addition | Location | Effect |
|----------|----------|--------|
| `docProps/custom.xml` | Package root | Custom document properties part |
| `word/commentsExtensible.xml` | Word part | Extensible comments part |
| Image content types (`png`, `jpeg`, `jpg`, `bmp`, `gif`, `svg`) | `[Content_Types].xml` | Default entries for image formats |
| `odttf` content type | `[Content_Types].xml` | Obfuscated font content type |
| Relationship `rId4` to `custom.xml` | `_rels/.rels` | Links the new custom properties part |
| `commentsExtensible.xml` Override | `[Content_Types].xml` | Content type for new part |

### A2. Explicit formatting on every text run
The writer added `w:rFonts w:ascii="Inter" w:hAnsi="Inter" w:cs="Inter"` to virtually every `<w:r>` element in `document.xml`, `header1.xml`, and `footer1.xml`. In the working document, these runs relied on the `<w:docDefaults>` font specification (which already declares Inter as the default font).

### A3. Explicit paragraph spacing defaults
The writer added `w:before="0"` to `<w:spacing>` elements throughout `document.xml`. Where the working document had `<w:spacing w:after="20"/>`, the broken version has `<w:spacing w:before="0" w:after="20"/>`.

### A4. Table cell borders on every cell
The writer added explicit `<w:tcBorders>` elements on every table cell, specifying per-edge border style, size, and color (`E2E8F0` for standard borders, `2563EB` for accent left borders on callout cells). In the working document, cell borders were inherited from the table-level `<w:tblBorders>` and the `w:tblStyle`.

### A5. Extra numbering definition
The writer added a second `w:abstractNum` (id=2) preserving the original en-dash bullet style at ilvl=0, while changing the primary abstractNum (id=1) to use a filled-circle bullet. The working document has only one abstractNum (id=0).

### A6. Namespaces added
The writer added `wpg`, `wpi`, `wps`, `wne` namespace declarations to root elements in several files. These are not referenced in any child elements.

---

## (2) What Was REMOVED by the Writer (in broken, missing from working)

### R1. Table structural properties (CRITICAL — P0)
The writer stripped the following from ALL tables in `document.xml`:

| Removed Element | Tables Affected | Effect |
|----------------|-----------------|--------|
| `w:tblStyle w:val="Normal"` | Table 1 | First table loses its style association |
| `w:tblInd` | Tables 2,4,5,6 | Tables lose left-margin offsets (e.g., `w:w="5"`, `w:w="-4"`) |
| `w:tblLook` | All 7 tables | Table style formatting bands (first row/column) lost |
| `w:tblCellMar` (table-level) | Tables 1,2,4,5,6,7 | Cell padding defaults changed |
| `w:tblPrEx` (row-level) | 9 rows across all tables | Per-row table property exceptions lost |
| `w:tblCellMar` (row-level) | 9 rows | Per-row cell margins lost |
| `w:space` on border elements | All bordered tables | Border spacing defaults to 0 instead of original values |

### R2. Font table entries (HIGH — P0)
The writer stripped all 5 font definitions from `word/fontTable.xml`:
- Inter (with altName: Cambria)
- Times New Roman (with signature metadata)
- Courier New (with signature metadata)
- Aptos Display (with signature metadata)
- Aptos (with signature metadata)

Only an empty `<w:fonts>` root remains.

### R3. Style definitions (CRITICAL — P0)
The writer removed these style definitions from `word/styles.xml`:

| Removed Style | Type | Effect |
|--------------|------|--------|
| Normal (default paragraph) | paragraph | No default paragraph style |
| DefaultParagraphFont (default character) | character | No default character style |
| TableNormal (default table) | table | No default table style |
| NoList (default numbering) | numbering | No default numbering style |
| **CommentReference** | character | **Dangling reference — Word repair dialog** |
| CommentText | paragraph | Dangling reference |
| CommentTextChar | character | Dangling reference |
| 376 latent styles | latent | Style gallery and quick-style menus affected |

The `CommentReference` style is referenced via `<w:rStyle w:val="CommentReference"/>` in both `document.xml` and `comments.xml` but has no matching definition. This is the documented root cause of the "broken-tables" fixture.

### R4. XML standalone declaration (P3)
The writer removed `standalone="yes"` from XML declarations in all 16 XML files.

### R5. Namespace declarations (P3)
The writer removed 8 `chartex` namespace declarations, `aink`, `am3d`, and `oel` from root elements.

### R6. Relationship references (P1)
The writer removed relationships for:
- `word/endnotes.xml` (rId6 in working)
- `word/theme/theme1.xml` (rId16 in working)
- `word/webSettings.xml`

The writer also changed all hyperlink relationship IDs from short numeric IDs to long pseudo-random strings.

### R7. Numbering metadata (P1)
The writer removed `nsid` and `tmpl` attributes from abstractNum elements, removed `w16cid:durableId` from num elements, and changed bullet characters.

### R8. Section properties (P2)
The writer removed `<w:cols w:space="720"/>` from `<w:sectPr>`.

### R9. Revision Save IDs (P3)
The writer removed `w:rsidR` and `w:rsidRDefault` from all `<w:p>` and `<w:tr>` elements.

### R10. Parts removed entirely
The writer did not produce these parts (present in working):
- `word/endnotes.xml`
- `word/theme/theme1.xml`
- `word/webSettings.xml`

---

## (3) Proposed Solutions

### Fix 1: Restore Table Structural Properties (P0)
**Target**: `src/scripts/office/validators/docx.ts` — add a `repairTableStructure()` method.

The writer must preserve these attributes through the write-then-repair cycle:
- `w:tblInd` on `<w:tblPr>` — must be copied from the source document's table properties
- `w:tblLook` on `<w:tblPr>` — must be preserved with original `w:val`, `w:firstRow`, `w:lastRow`, `w:firstColumn`, `w:lastColumn`, `w:noHBand`, `w:noVBand`
- `w:tblStyle` on `<w:tblPr>` — must be preserved
- `w:tblCellMar` on `<w:tblPr>` — must be preserved
- `w:tblPrEx` on `<w:tr>` — must be preserved with nested `w:tblCellMar`
- `w:space` on border elements — must be preserved

Implementation approach (in the repair pipeline):
1. Read the original document's table XML before the writer processes it
2. After the writer produces its output, re-apply table structural attributes from the original
3. Alternatively, fix the writer itself to never strip these attributes

Since the fixture represents output from the writer after multiple repair passes, the root cause is likely in jubarte's XML serialization layer. The validator's `repairMissingStyleDefinitions` already handles style injection — a similar approach is needed for table properties.

### Fix 2: Inject Missing `CommentReference` Style Definition (P0)
**Target**: Already implemented in `src/scripts/office/validators/docx.ts:115-122` as part of `WELL_KNOWN_STYLE_DEFINITIONS`.

The current code at `repairMissingStyleDefinitions` (line 1213) already injects `CommentReference` and other well-known styles when referenced but not defined. The fixture `sample-document.broken-tables.docx` was specifically created to test this repair path (per `tests/fixtures/broken/README.md:21`).

**Verification**: The working (repaired) file has the `CommentReference` style; the broken file does not. The existing repair works correctly for this specific issue.

### Fix 3: Preserve or Reconstruct Font Table (P0)
**Target**: `src/scripts/office/validators/docx.ts` — add a `repairFontTable()` method, or fix the writer.

The writer should either:
a) Preserve the original `fontTable.xml` unchanged from the source document (preferred), or
b) If it must regenerate the font table, collect all `w:rFonts` references from the document and construct minimal font entries for each referenced font family

Implementation:
```typescript
// Scan all w:rFonts references across document.xml, header*.xml, footer*.xml
// For each font name not in fontTable.xml, add a minimal entry:
// <w:font w:name="Inter">
//   <w:charset w:val="00"/>
//   <w:family w:val="swiss"/>
// </w:font>
```

### Fix 4: Preserve or Document Relationship Identity Strategy (P1)
**Target**: Writer internals (jubarte).

The writer should either:
a) Preserve original relationship IDs from the source document when the target part exists in both source and output, or
b) Use a deterministic ID scheme and ensure `sectPr` header/footer references match the generated IDs

The core requirement is consistency: `w:headerReference r:id` and `w:footerReference r:id` in `document.xml` must match the IDs in `word/_rels/document.xml.rels`.

### Fix 5: Preserve Numbering Definitions (P1)
**Target**: Writer internals.

The writer must:
- Preserve `nsid` and `tmpl` attributes on `w:abstractNum` elements
- Preserve `w16cid:durableId` on `w:num` elements
- Preserve original bullet characters (don't change en-dash to filled circle)
- Preserve `w:tplc` attributes on `w:lvl` elements (don't replace with `w15:tentative`)

### Fix 6: Writer Should Preserve Default Styles (P2)
**Target**: Writer internals.

The writer should preserve these minimum default styles from the source document:
- `Normal` (default paragraph)
- `DefaultParagraphFont` (default character)  
- `TableNormal` (default table)
- `NoList` (default numbering)

### Fix 7: Writer Should Not Strip Optional Parts (P2)
**Target**: Writer internals.

The writer should preserve these parts when present in the source:
- `word/endnotes.xml`
- `word/theme/theme1.xml`
- `word/webSettings.xml`

### Fix 8: Add XML `standalone="yes"` (P3)
**Target**: Writer's XML serializer.

Add `standalone="yes"` to all XML declarations to match OOXML canonical form.

### Fix 9: Preserve or Clean-Remove Unused Namespaces (P3)
**Target**: Writer's XML serializer.

Either preserve all namespace declarations from the source document, or produce a minimal clean set (the broken version's approach is reasonable but should also preserve any namespaces actually used in child elements).

---

## Implementation Priority

| Priority | Fix | Effort | Dependencies |
|----------|-----|--------|--------------|
| P0 | Fix 2: CommentReference style injection | Already implemented | None — needs validation |
| P0 | Fix 1: Restore table structural properties | Medium | New repair method in docx.ts |
| P0 | Fix 3: Preserve/reconstruct font table | Medium | New repair method or writer fix |
| P1 | Fix 4: Relationship ID consistency | Small | Writer fix |
| P1 | Fix 5: Preserve numbering definitions | Medium | Writer fix |
| P2 | Fix 6: Preserve default styles | Small | Writer fix |
| P2 | Fix 7: Preserve optional parts | Small | Writer fix |
| P3 | Fix 8: XML standalone | Trivial | Writer's XML serializer |
| P3 | Fix 9: Namespace hygiene | Small | Writer's XML serializer |
