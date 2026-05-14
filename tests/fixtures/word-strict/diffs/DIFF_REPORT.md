# Diff Report: working vs broken (word-strict)

**Generated**: 2026-05-13  
**Working file**: `tests:fixtures:broken:sample-document.broken-tables.docx-repaired.docx` (repaired by repair pipeline)  
**Broken file**: `sample-document.broken-tables.docx` (pre-repair snapshot from jubarte/docx-js-editor writer)  
**Total common files**: 16 — ALL differ  
**Files only in broken**: 2  
**Files only in working**: 3

---

## Group 1 — Files Only in One Side

| Present only in | File |
|-----------------|------|
| BROKEN | `docProps/custom.xml` |
| BROKEN | `word/commentsExtensible.xml` |
| WORKING | `word/endnotes.xml` |
| WORKING | `word/theme/theme1.xml` |
| WORKING | `word/webSettings.xml` |

---

## Group 2 — XML Declaration Changes (ALL 16 common XML files)

Every XML file in the working version starts with:
```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
```

Every XML file in the broken version drops `standalone="yes"`:
```xml
<?xml version="1.0" encoding="UTF-8"?>
```

**Impact**: Cosmetic. `standalone="yes"` is the OOXML canonical form; its absence is tolerated.

---

## Group 3 — Namespace Declaration Changes (ALL 16 common XML files)

The **working** version carries a large set of namespace declarations on root elements, including 8 `chartex` variants (`cx` through `cx8`), `aink`, `am3d`, `oel`, and several legacy drawing namespaces.

The **broken** version strips most of these, keeping only the core wordprocessing namespaces (`w`, `w14`, `w15`, `r`, `mc`, `v`, `o`, `w10`, `wp`, `wp14`, `m`) and adds `wpg`, `wpi`, `wps`, `wne` in some files.

**Impact**: Minor. Unused namespace declarations are harmless. The added namespaces (`wpg`, `wpi`, `wps`) are not referenced anywhere in the broken document, suggesting they were added speculatively by the repair pipeline.

---

## Group 4 — Relationship Identity Changes

### `_rels/.rels`
- Broken ADDED a relationship `rId4` to `docProps/custom.xml` (matching the extra `custom.xml` in Group 1).

### `word/_rels/document.xml.rels`
- Working has 12 relationships with short sequential IDs: `rId1` through `rId16`.
- Broken has different relationship IDs. Hyperlink references changed from short IDs (`rId7`, `rId8`, `rId12`) to long random-looking IDs (`rIdclrrgvzn-co-bssfjluef`, `rIdjuy5xlgboggzopeudzji5`, `rIdbckue47mmm_l1hfnxuyyt`).
- Working has `theme/theme1.xml` relationship (`rId16`); broken does not.
- Working has `endnotes.xml` relationship (`rId6`); broken does not.
- Header/footer references changed from `rId13`/`rId14` to `rId6`/`rId7`.

**Impact**: Functional. The header/footer reference IDs must match between `sectPr` and `.rels`. If they don't, Word displays "Document Recovery" or blank headers/footers.

---

## Group 5 — TABLE STRUCTURAL ELEMENTS REMOVED (word/document.xml) — CRITICAL

This is the namesake issue. The broken version systematically stripped table-structure attributes from every `<w:tbl>` element in the document. There are 7 tables in the document, and **all 7** show these removals:

| Element removed from broken | Found in working | Count in document |
|---|---|---|
| `w:tblStyle w:val="Normal"` | Present on table 1 (props table) | 1 |
| `w:tblInd` | Present on tables 2, 4, 5, 6 | 4 |
| `w:tblLook` | Present on all 7 tables | 7 |
| `w:tblCellMar` (table-level) | Present on 4 tables | 4 |
| `w:tblPrEx` (row-level table property exceptions) | Present on 9 rows across all tables | 9 |
| `w:tblCellMar` (row-level, inside `tblPrEx`) | Present on 9 rows | 9 |
| `w:space` on border elements | Present on all bordered tables | many |

### Specific removals by table:

1. **Table 1 (props table)**: Removed `w:tblStyle`, `w:tblLook`, `w:tblCellMar`, `w:space` from borders.
2. **Table 2 (callout table)**: Removed `w:tblInd`, `w:tblLook`, `w:tblCellMar`, `w:tblPrEx` (row-level), `w:space`.
3. **Table 3 (code block table)**: Removed `w:tblLook`, `w:tblPrEx` (row-level).
4. **Table 4 (code block table)**: Removed `w:tblInd`, `w:tblLook`, `w:tblCellMar`, `w:tblPrEx` (row-level).
5. **Table 5 (code block table)**: Removed `w:tblInd`, `w:tblLook`, `w:tblCellMar`, `w:tblPrEx` (row-level).
6. **Table 6 (code block table)**: Removed `w:tblInd`, `w:tblLook`, `w:tblCellMar`, `w:tblPrEx` (row-level).
7. **Table 7 (sign-off table)**: Removed `w:tblLook`, `w:tblCellMar`, `w:tblPrEx` (row-level).

**Impact**: CRITICAL. Word uses `w:tblInd` to offset tables from the left margin. Without it, tables align to the left margin instead of their intended position. `w:tblLook` controls table-style formatting bands (first row/column). `w:tblPrEx` and `w:tblCellMar` on rows control per-row cell margins. Removal of these causes table layout to shift and cell padding to change.

---

## Group 6 — Paragraph Attribute Changes (word/document.xml)

1. **`w:rsidR` / `w:rsidRDefault`** removed from ALL paragraphs in broken. Working has them on every `<w:p>` and `<w:tr>`.
2. **`w14:textId`** changed from `77777777` (working) to unique hex values (broken).
3. **`w:spacing`** elements now include explicit `w:before="0"` in addition to `w:after` values.
4. **`w:rFonts`** added to virtually every `<w:r>` element in broken, where working relied on the document-default font from `<w:docDefaults>`.

**Impact**: `rsidR` removal is cosmetic (these are revision save IDs). `textId` changes are fine (unique IDs). Explicit `w:before="0"` is fine. Explicit `w:rFonts` on every run is verbose but not harmful.

---

## Group 7 — Font Table Emptied (word/fontTable.xml)

**Working** defines 5 fonts: Inter, Times New Roman, Courier New, Aptos Display, Aptos — each with panose-1, charset, family, pitch, and signature metadata.

**Broken** font table is completely empty — only the root `<w:fonts>` element with namespace declarations, zero `<w:font>` children.

**Impact**: High for interoperability. Without font definitions, Word and other processors cannot map font names to their metrics, fallback behavior, or embedding hints. Documents may render with incorrect fonts on systems lacking the named fonts.

---

## Group 8 — Styles Drastically Reduced (word/styles.xml)

### Removed in broken (present in working):

| Style | Type |
|-------|------|
| Normal (default paragraph) | paragraph |
| DefaultParagraphFont (default character) | character |
| TableNormal (default table) | table |
| NoList (default numbering) | numbering |
| CommentText | paragraph |
| CommentTextChar | character |
| CommentReference | character |
| CommentSubject | (would need definition) |
| 376 latent styles (`w:latentStyles`) | latent |
| Title (original with uiPriority) | paragraph |

### Changed in broken:

- All heading styles (`Heading1`-`Heading6`) restructured: removed `w:uiPriority`, `w:semiHidden`, `w:unhideWhenUsed`, `w:outlineLvl`. Added `w:basedOn w:val="Normal"` and `w:next w:val="Normal"`.
- Heading names changed from lowercase ("heading 1") to title case ("Heading 1").
- `Strong` style changed from `w:customStyle="1" w:styleId="Strong1"` to standard `w:styleId="Strong"` with `w:basedOn`/`w:next`.
- `Hyperlink` style added `w:basedOn w:val="DefaultParagraphFont"`.
- `FootnoteReference` style added `w:basedOn w:val="DefaultParagraphFont"`.
- `FootnoteText` style added `w:basedOn w:val="Normal"` and explicit line spacing.
- `FootnoteTextChar` style lost `w:customStyle="1"`, gained `w:basedOn`.

### CRITICAL: Missing `CommentReference` style

The broken document references `<w:rStyle w:val="CommentReference"/>` in both `document.xml` and `comments.xml`, but the style definition is NOT present in `styles.xml`. This causes Word's "Open" dialog to flag a dangling style reference error.

**Impact**: CRITICAL. Missing well-known style definitions (especially `CommentReference`) cause Word to display repair dialogs on open.

---

## Group 9 — Numbering Restructured (word/numbering.xml)

- `abstractNumId` changed from `0` to `1`.
- `nsid` and `tmpl` attributes removed from abstractNum elements.
- `tplc` attributes replaced with `w15:tentative="1"`.
- Bullet character at ilvl=0 changed from en-dash (–) to filled circle (●).
- A second abstractNum (id=2) was added to preserve the old en-dash bullet style.
- `w16cid:durableId` removed from `<w:num>` elements.

**Impact**: The bullet character change is visible. The removal of `nsid`/`tmpl` and change in numbering IDs may cause list numbering to reset or lose association with their lists when opened in Word.

---

## Group 10 — Comment ID Changes

- Comment ID changed from `4` to `100`.
- Comment reference durableId changed from `456E2E6B` to `B68569E0`.
- Comment date format: working uses `2026-03-05T19:41:00Z`, broken uses `2026-03-05T19:41:08Z`.

**Impact**: Minor. Comment IDs are internal; the mismatch between `document.xml` comment markers and `comments.xml` comment definitions would cause issues, but within each file they're consistent.

---

## Group 11 — Section Properties (word/document.xml footer)

- Working: `<w:sectPr w:rsidR="00B90B26">` with `<w:cols w:space="720"/>`.
- Broken: `<w:sectPr>` (no `rsidR`) without `<w:cols>`.

**Impact**: `w:cols` removal changes column layout. In single-column documents this has no visual effect but deviates from the canon.

---

## Group 12 — Content Type Manifest ([Content_Types].xml)

- Broken added image-related Default entries (`png`, `jpeg`, `jpg`, `bmp`, `gif`, `svg`) and `odttf` (obfuscated font).
- Broken added `commentsExtensible.xml` Override, `custom.xml` Override.
- Broken removed `endnotes.xml`, `webSettings.xml`, `theme/theme1.xml` Overrides (matching Group 1 file absences).

**Impact**: Must match actual files present. The broken manifest correctly reflects which parts exist.

---

## Summary of Differences by Severity

| Severity | Group | Description |
|----------|-------|-------------|
| **P0** | Group 5 | Table structural elements removed (`w:tblInd`, `w:tblLook`, `w:tblPrEx`, `w:tblCellMar`, `w:tblStyle`) |
| **P0** | Group 8 | Missing `CommentReference` style definition causing Word repair dialog |
| **P0** | Group 7 | Font table emptied (5 fonts removed) |
| **P1** | Group 4 | Relationship ID changes (header/footer references mismatch risk) |
| **P1** | Group 9 | Numbering restructured (bullet characters changed, IDs shifted) |
| **P2** | Group 11 | Section properties: `w:cols` removed |
| **P2** | Group 8 | Removal of default styles (Normal, DefaultParagraphFont, TableNormal, NoList) |
| **P3** | Group 2 | XML declaration `standalone="yes"` removed |
| **P3** | Group 3 | Namespace declarations reduced |
| **P3** | Group 6 | `w:rsidR` removed, `w:rFonts` explicitly added to all runs |
| **P3** | Group 10 | Comment ID renumbering |
| **P3** | Group 12 | Content type manifest changes (correct for actual contents) |
