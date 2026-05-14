# Diff Report: Word-Repaired vs Repairer-Output

**Compared files:**
- **Working (Word-repaired):** `second-pass-word-repaired-sample-document.broken-tables.docx`
- **Broken (repairer output):** `second-pass-actually-broken-sample-document.broken-tables.docx`

**Date of comparison:** 2026-05-13

---

## Summary

| Metric | Count |
|--------|-------|
| Common files (exist in both) | 16 |
| Files only in Working (Word-repaired) | 3 |
| Files only in Broken (repairer output) | 2 |
| Common files with differences | 16 |
| Common files identical | 0 |

---

## A. Files Present Only in One Side

### A1. Files only in Working (Word-repaired)

| File | Description |
|------|-------------|
| `word/endnotes.xml` | Endnotes part — declared in Content_Types and document.xml.rels |
| `word/theme/theme1.xml` | Office theme — declared in Content_Types and document.xml.rels |
| `word/webSettings.xml` | Web settings part — declared in Content_Types |

### A2. Files only in Broken (repairer output)

| File | Description |
|------|-------------|
| `docProps/custom.xml` | Custom document properties — declared in Content_Types and _rels/.rels |
| `word/commentsExtensible.xml` | Comments Extensible part — declared in Content_Types and document.xml.rels |

---

## B. Differences in Common Files

### Group 1: XML Declaration (`standalone="yes"`)

**Affected files:** `_rels/.rels`, `[Content_Types].xml`, `docProps/app.xml`, `docProps/core.xml`, `word/_rels/document.xml.rels`, `word/footer1.xml`, `word/footnotes.xml`, `word/header1.xml`, `word/numbering.xml`, `word/settings.xml`, `word/styles.xml`, `word/document.xml`

**Description:** In Working, the XML declaration is `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>`. In Broken, it is `<?xml version="1.0" encoding="UTF-8"?>` — the `standalone="yes"` attribute is absent.

**Files where this is the primary difference:** `word/comments.xml`, `word/commentsExtended.xml`, `word/commentsIds.xml` retain `standalone="yes"` in both.

---

### Group 2: Namespace Reductions on Root Elements

**Affected files:** `word/comments.xml`, `word/commentsExtended.xml`, `word/commentsIds.xml`, `word/header1.xml`, `word/footer1.xml`, `word/footnotes.xml`, `word/numbering.xml`, `word/document.xml`

**Description:** In Broken, many namespace declarations present in Working are removed from the root element of each XML part. Specifically:

- Working includes namespaces like `cx` through `cx8` (chartex variants), `aink` (ink), `am3d` (model3d), `m` (math), `o` (office), `v` (VML), `w10`, `w16cex`, `w16cid`, `w16`, `w16sdtdh`, `w16sdtfl`, `w16se`, `w16du`, `wp14`, `wp`, `wpg`, `wpi`, `wps`, `oel`, `mc`.
- Broken removes many of these; retains only `mc`, `o`, `r`, `m` (in some), `v`, `wp14` (in some), `wp` (in some), `w10`, `w`, `w14`, `w15`, `wpg` (in some), `wps` (in some), `wpi` (in some).

Additionally, in `word/footnotes.xml`, `word/numbering.xml`, `word/settings.xml`, `word/document.xml`, Broken adds the `wne` namespace (`http://schemas.microsoft.com/office/word/2006/wordml`) that is absent in Working.

---

### Group 3: Package Relationships (`_rels/.rels`)

**File:** `_rels/.rels`

**Differences:**
- Working: 3 relationships (rId1=officeDocument, rId2=core-properties, rId3=extended-properties), with `standalone="yes"`
- Broken: 4 relationships — adds `rId4` for `custom-properties` pointing to `docProps/custom.xml`
- Relationship IDs use the same targets but ordering differs

---

### Group 4: Content Types (`[Content_Types].xml`)

**File:** `[Content_Types].xml`

**Differences in Broken compared to Working:**

Added defaults:
- `image/png`, `image/jpeg`, `image/jpg`, `image/bmp`, `image/gif`, `image/svg+xml`
- `application/vnd.openxmlformats-officedocument.obfuscatedFont`

Added overrides:
- `/docProps/core.xml`
- `/docProps/custom.xml`
- `/docProps/app.xml`
- `/word/fontTable.xml`
- `/word/header1.xml`
- `/word/footer1.xml`
- `/word/commentsExtensible.xml`

Removed overrides (present in Working, absent in Broken):
- `/word/webSettings.xml`
- `/word/endnotes.xml`
- `/word/theme/theme1.xml`

Attribute order differences: Working uses `PartName` then `ContentType`; Broken uses `ContentType` then `PartName`.

---

### Group 5: Document Properties — App (`docProps/app.xml`)

**File:** `docProps/app.xml`

**Differences:**
- Working: Contains full extended properties including `Template`, `TotalTime`, `Pages` (3), `Words` (815), `Characters` (4648), `Application`, `DocSecurity`, `Lines` (38), `Paragraphs` (10), `ScaleCrop`, `Company`, `LinksUpToDate`, `CharactersWithSpaces` (5453), `SharedDoc`, `HyperlinksChanged`, `AppVersion` (16.0000)
- Broken: Properties element is entirely empty — contains only the root element with `xmlns` attributes and no child elements

---

### Group 6: Document Properties — Core (`docProps/core.xml`)

**File:** `docProps/core.xml`

**Differences:**
- Working: `lastModifiedBy="Arthur Souza Rodrigues"`, `revision=1`, created and modified timestamps in 2026-05-14T01:27:00Z format
- Broken: `lastModifiedBy="Un-named"`, `revision=1` (as element with newline), created timestamp `2026-03-05T19:36:13.142Z`, modified timestamp `2026-05-10T00:07:14.347Z`
- Broken formats `revision`, `created`, `modified` as elements with internal whitespace rather than inline

---

### Group 7: Document Relationships (`word/_rels/document.xml.rels`)

**File:** `word/_rels/document.xml.rels`

**Differences:**
- Working: Uses sequential numeric IDs (rId1 through rId16). Relationships: numbering(rId1), styles(rId2), settings(rId3), comments(rId4), footnotes(rId5), **endnotes(rId6)**, hyperlinks(rId7,rId8,rId12), commentsExtended(rId10), commentsIds(rId11), header(rId13), footer(rId14), fontTable(rId15), **theme(rId16)**
- Broken: Uses rId1-rId7 for core, then non-sequential string-based IDs for hyperlinks (`rIdclrrgvzn-co-bssfjluef`, `rIdjuy5xlgboggzopeudzji5`, `rIdbckue47mmm_l1hxfnuyyt`), then rId11-rId13 for fontTable, commentsExtended, commentsIds, and **adds commentsExtensible relationship**
- Missing from Broken: endnotes relationship (rId6 in Working), theme relationship (rId16 in Working)
- Added in Broken: commentsExtensible relationship

---

### Group 8: Comments (`word/comments.xml`)

**File:** `word/comments.xml`

**Differences:**
- Comment ID: `w:id="4"` in Working; `w:id="100"` in Broken
- Comment date format: Working has `w:date="2026-03-05T19:41:00Z"`; Broken has `w:date="2026-03-05T19:41:08Z"`
- Working's comment run has `w:rsidR="00B21E0F"` and `w:rsidRDefault="00000000"`; Broken lacks these
- Comment text trailing whitespace reformatted (line76-77): Working has `Just like that!         ` inline, Broken splits to `Just like that!` followed by newline-indented close tag
- Namespace reductions as described in Group2

---

### Group 9: Comments Extended (`word/commentsExtended.xml`)

**File:** `word/commentsExtended.xml`

**Differences:**
- Working: `xmlns:wpg`, `xmlns:wpi`, `xmlns:wps`, `xmlns:o`, `xmlns:v`, `xmlns:w10`, `xmlns:w16cex`, `xmlns:w16cid`, `xmlns:w16`, `xmlns:w16du`, `xmlns:w16sdtdh`, `xmlns:w16sdtfl`, `xmlns:w16se`, `xmlns:aink`, `xmlns:am3d`
- Broken: Missing `xmlns:wpg`, `xmlns:wpi`, `xmlns:wps`, `xmlns:w16du`, `xmlns:w16sdtdh`, `xmlns:w16sdtfl`, `xmlns:w16se`, `xmlns:aink`, `xmlns:am3d`; adds `xmlns:o`, `xmlns:v`, `xmlns:w10`
- Body content (`w15:commentEx`) is identical

---

### Group 10: Comments IDs (`word/commentsIds.xml`)

**File:** `word/commentsIds.xml`

**Differences:**
- `w16cid:paraId`: `"456E2E6B"` in Working; `"B68569E0"` in Broken
- `w16cid:durableId`: `"456E2E6B"` in Working; `"05B4A74E"` in Broken
- Namespace reductions as described in Group2
- `standalone="yes"` retained in both

---

### Group 11: Header (`word/header1.xml`)

**File:** `word/header1.xml`

**Differences in Broken compared to Working:**

Removed:
- `w:rsidR="00B21E0F"` and `w:rsidRDefault="00000000"` from the paragraph element
- Trailing whitespace in text content (replaced with newline-indented close tag)

Added:
- `w:spacing w:after="0"` in paragraph properties
- `w:rFonts w:ascii="Inter" w:hAnsi="Inter" w:cs="Inter"` in run properties
- Namespace reductions as described in Group2

---

### Group 12: Footer (`word/footer1.xml`)

**File:** `word/footer1.xml`

**Differences in Broken compared to Working:**

Removed:
- `w:rsidR="00B21E0F"` and `w:rsidRDefault="00000000"` from the paragraph element
- Trailing whitespace in text content across all runs (12 text/instrText elements reformatted)

Added:
- `w:rFonts w:ascii="Inter" w:hAnsi="Inter" w:cs="Inter"` in every run's run properties (12 instances)
- Namespace reductions as described in Group2

---

### Group 13: Footnotes (`word/footnotes.xml`)

**File:** `word/footnotes.xml`

**Differences in Broken compared to Working:**

Removed:
- `w:rsidR` and `w:rsidRDefault` attributes from paragraph elements
- `standalone="yes"`

Added/Changed:
- `mc:Ignorable="w14 w15 wp14"` attribute on root element
- `wne` namespace added
- Paragraph IDs changed: separator id=-1 uses `paraId="00000044"` (was `"07760C62"`), continuationSeparator id=0 uses `paraId="00000046"` (was `"797448A0"`)
- Each footnote paragraph gets explicit `<w:pPr>` with `<w:spacing w:after="0" w:line="240" w:lineRule="auto"/>`
- Each footnote reference run gets `<w:rPr>` with `<w:rStyle w:val="FootnoteReference"/>` and `<w:footnoteRef/>`

---

### Group 14: Font Table (`word/fontTable.xml`)

**File:** `word/fontTable.xml`

**Differences:**
- Working: Contains 5 font definitions:
  - Inter (altName Cambria, panose1, charset, family roman, notTrueType)
  - Times New Roman (panose1, charset, family roman, sig)
  - Courier New (panose1, charset, family modern, sig)
  - Aptos Display (panose1, charset, family swiss, sig)
  - Aptos (panose1, charset, family swiss, sig)
- Broken: Empty `<w:fonts>` element — no font definitions, fewer namespace declarations

---

### Group 15: Settings (`word/settings.xml`)

**File:** `word/settings.xml`

**Differences:**

In Working, present; in Broken, absent:
- `w:zoom w:percent="150"`
- `w:defaultTabStop w:val="720"`
- `w:characterSpacingControl w:val="doNotCompress"`
- `w:footnotePr` (with footnote ids -1 and 0)
- `w:endnotePr` (with endnote ids -1 and 0)
- Multiple compatSettings beyond compatibilityMode: `overrideTableStyleFontSizeAndJustification`, `enableOpenTypeFeatures`, `doNotFlipMirrorIndents`, `differentiateMultirowTableHeaders`, `useWord2013TrackBottomHyphenation`
- `w:rsids` (rsidRoot and 4 rsid values)
- `m:mathPr` (full math properties with font, break, margin, etc.)
- `w:themeFontLang w:val="en-US"`
- `w:clrSchemeMapping`
- `w:decimalSymbol w:val="."`
- `w:listSeparator w:val=","`
- `w15:docId w15:val="{C147CAA8-2A90-E641-9A2B-CDACA5D2EC1B}"`

In Broken, present; in Working, absent:
- Additional namespaces: `wpc`, `wp14`, `wp`, `wpg`, `wpi`, `wne`, `wps`
- `w:evenAndOddHeaders w:val="false"`

Both have:
- `w:displayBackgroundShape`
- `w:compat` with `w:compatSetting` for `compatibilityMode` (attribute order differs: Working has `name/uri/val`, Broken has `val/uri/name`)

---

### Group 16: Numbering (`word/numbering.xml`)

**File:** `word/numbering.xml`

**Differences in Broken compared to Working:**

First abstractNum changes:
- `w:abstractNumId` changed from `"0"` to `"1"`
- `w:nsid` removed
- `w:tmpl` removed
- `w:tplc` on each level replaced with `w15:tentative="1"`
- `w:multiLevelType w:val="hybridMultilevel"` added

Second abstractNum changes:
- `w:abstractNumId` changed from `"1"` to `"2"`
- `w:nsid` removed
- `w:tmpl` removed
- `w:tplc` on level0 replaced with `w15:tentative="1"`
- Levels 1 through 8 (all with `w:numFmt w:val="decimal"`, empty `w:lvlText`, `w:lvlJc w:val="left"`) are entirely removed

Num instances:
- `w16cid:durableId` removed from both `w:num` elements
- `w:abstractNumId` values updated to match new abstractNum IDs: `w:val="0"` → `w:val="1"`, `w:val="1"` → `w:val="2"`
- `w:lvlOverride`/`w:startOverride` unchanged

---

### Group 17: Styles (`word/styles.xml`)

**File:** `word/styles.xml`

**Differences:** Working has ~407 lines of style definitions; Broken has ~30 lines.

**In Working, present; in Broken, absent:**
- Entire `<w:latentStyles>` section (376 `lsdException` entries)
- Styles: `CommentText` (paragraph, annotation text), `CommentTextChar` (character), `CommentReference` (character)
- Style: `Strong1` (custom paragraph style with bold)

**In Working, present; in Broken, relocated/reordered:**
- `Title` style: In Working after Heading6; in Broken first before Heading1
- `Normal` (default paragraph): In Working first; in Broken after CommentReference
- `DefaultParagraphFont` (default character): In Working after Heading6; in Broken after Normal
- `TableNormal` (default table): In Working after DefaultParagraphFont; in Broken after DefaultParagraphFont
- `NoList` (default numbering): In Working after TableNormal; in Broken after TableNormal

**In Broken, present; in Working, absent or different:**
- `Strong` style: In Broken uses `w:styleId="Strong"`; in Working uses `w:styleId="Strong1"` with `w:customStyle="1"`
- Heading styles lose `w:uiPriority`, `w:semiHidden`, `w:unhideWhenUsed`, `w:outlineLvl` properties
- Style names lose leading lowercase: `"heading 1"` → `"Heading 1"`, etc.
- `FootnoteText` paragraph style gains `<w:spacing w:after="0" w:line="240" w:lineRule="auto"/>`

**Style ordering in Broken vs Working:**
- Broken ordering: Title, Heading1-6, Strong, ListParagraph, FootnoteText, FootnoteReference, FootnoteTextChar, CommentReference, Normal, DefaultParagraphFont, TableNormal, NoList
- Working ordering: Normal, headings, DefaultParagraphFont, TableNormal, NoList, Title, Strong1, ListParagraph, footnote styles, comment styles

**`CommentReference` style in Broken:**
- Present in Broken (referenced by document.xml run for comment anchor), but was present in Working at its expected position too. The diff shows its definition is retained but its position changed.

---

### Group 18: Main Document (`word/document.xml`)

**File:** `word/document.xml`

This is the largest diff (3,197 lines). The differences are systematic and repetitive.

#### 18a. Revision Session ID Removal
- **Affects:** Every `<w:p>` and `<w:tr>` element
- **Change:** Working contains `w:rsidR="00B21E0F"`, `w:rsidRDefault="00000000"`, `w:rsidR="00D427AF"` attributes; Broken has none of these
- **Scope:** All paragraphs (~70+) and all table rows (~18+)

#### 18b. Explicit Font Specification (`w:rFonts`)
- **Affects:** Nearly every `<w:r>` element in body text
- **Change:** Broken adds `<w:rFonts w:ascii="Inter" w:hAnsi="Inter" w:cs="Inter"/>` to runs that lack it in Working
- **Count:** 77 instances
- **Context:** Working relies on style inheritance for font; Broken makes it explicit per run

#### 18c. Explicit Font Size (`w:sz`/`w:szCs`)
- **Affects:** Body text runs that lack explicit size in Working
- **Change:** Broken adds `<w:sz w:val="22"/>` and `<w:szCs w:val="22"/>` to many runs
- **Context:** Working relies on the document default style (sz=22); Broken makes it explicit

#### 18d. Paragraph Spacing Normalization (`w:spacing w:before="0"`)
- **Affects:** Nearly every paragraph
- **Change:** Broken adds `w:before="0"` to spacing elements where Working only specifies `w:after` or `w:before`
- **Count:** 56 instances of `w:spacing w:before="0" w:after="0"` added

#### 18e. Individual Cell Borders (`w:tcBorders`)
- **Affects:** Every table cell (`<w:tc>`) in every table
- **Change:** Broken adds explicit `<w:tcBorders>` inside `<w:tcPr>` for every cell, with per-edge specifications
- **Count:** 78 instances
- **Border patterns vary:**
  - Most data cells: all 4 edges with `w:sz="1" w:color="E2E8F0"`
  - Info/callout cells (full-width rows): left edge with `w:sz="8" w:color="2563EB"` (blue accent), other edges `w:sz="1" w:color="E2E8F0"`
  - Sign-off table cells: all edges with `w:val="none" w:sz="0"`
- **Working:** No `<w:tcBorders>` — cell borders are either absent entirely or inherited

#### 18f. Table Borders — `w:space` Attribute Removal
- **Affects:** 5 of 6 tables (tables 1-5 that have borders)
- **Change:** Working has `w:space="0"` on every border edge element (`w:top`, `w:left`, `w:bottom`, `w:right`, `w:insideH`, `w:insideV`); Broken omits `w:space="0"`
- **Count:** 30 border elements lose `w:space="0"`
- **Exception:** The 6th table (sign-off) has no borders in Working, but Broken adds `w:val="none" w:sz="0"` borders; this table does not appear in the space removal pattern

#### 18g. Table Style Reference Removal (`w:tblStyle`)
- **Affects:** First table only
- **Change:** Working has `<w:tblStyle w:val="Normal"/>` in the first table's `<w:tblPr>`; Broken removes it
- **Other tables:** No `w:tblStyle` in either version

#### 18h. Table Look Removal (`w:tblLook`)
- **Affects:** All 6 tables
- **Change:** Working has `<w:tblLook w:val="04A0" w:firstRow="1" w:lastRow="0" w:firstColumn="1" w:lastColumn="0" w:noHBand="0" w:noVBand="1"/>` in every table's `<w:tblPr>`; Broken removes it from all tables

#### 18i. Cell Shading — `w:color` Attribute Removal
- **Affects:** All table cells with `<w:shd>`
- **Change:** Working has `w:color="auto"` on all `<w:shd>` elements; Broken omits `w:color="auto"`, keeping only `w:val="clear" w:fill="..."`
- **Context:** `w:color="auto"` is the default foreground color; its removal is a simplification

#### 18j. Hyperlink Relationship ID Changes
- **Affects:** 3 hyperlinks in the document body
- **Changes:**
  - npm link: `r:id="rId7"` → `r:id="rIdclrrgvzn-co-bssfjluef"`
  - github link: `r:id="rId8"` → `r:id="rIdjuy5xlgboggzopeudzji5"`
  - PLUGINS.md link: `r:id="rId12"` → `r:id="rIdbckue47mmm_l1hxfnuyyt"`

#### 18k. Comment Reference ID Changes
- **Affects:** Comment range markers and reference in the document body
- **Changes:** All comment IDs changed from `w:id="4"` to `w:id="100"`:
  - `w:commentRangeStart w:id="4"` → `w:commentRangeStart w:id="100"`
  - `w:commentRangeEnd w:id="4"` → `w:commentRangeEnd w:id="100"`
  - `w:commentReference w:id="4"` → `w:commentReference w:id="100"`

#### 18l. Tracked Change ID Changes
- **Affects:** All `<w:del>` and `<w:ins>` elements
- **Changes:** IDs shifted by +9 or +10:
  - `w:del w:id="0"` → `w:del w:id="10"`
  - `w:ins w:id="1"` → `w:ins w:id="11"`
  - `w:del w:id="2"` → `w:del w:id="12"`
  - `w:ins w:id="3"` → `w:ins w:id="13"`
  - `w:del w:id="5"` → `w:del w:id="14"`
  - `w:ins w:id="6"` → `w:ins w:id="15"`

#### 18m. Table Row Paragraph/Text IDs Regenerated
- **Affects:** All table rows
- **Change:** `w14:paraId` and `w14:textId` values on `<w:tr>` elements are regenerated. They refer to different IDs in Working vs Broken (e.g., `7E9492D4`/`77777777` → `570F6BE0`/`00000035`)

#### 18n. Last Rendered Page Break Removal
- **Affects:** 2 locations
- **Change:** Working contains `<w:lastRenderedPageBreak/>` before the Comments section text and before the "Ref methods" heading; Broken removes both

#### 18o. Text Whitespace Reformatting
- **Affects:** All text runs in the document
- **Change:** Working stores trailing whitespace inside `w:t` elements (e.g., `"...text         "`). Broken splits the closing tag to a new line (e.g., `"...text\n        "`)
- **Context:** This is a formatting-only difference; the semantic content is identical

#### 18p. List Paragraph Indentation (`w:ind`)
- **Affects:** 4 paragraphs in the "Ground rules" / contributing section using `ListParagraph` style
- **Change:** Broken adds `<w:ind w:left="640" w:hanging="320"/>` to list paragraph properties; Working lacks this explicit indentation

#### 18q. Section Properties Changes
- **Affects:** `<w:sectPr>` at end of document body
- **Changes:**
  - `w:rsidR="0008393D"` removed from `<w:sectPr>`
  - `w:headerReference r:id="rId13"` → `r:id="rId6"`
  - `w:footerReference r:id="rId14"` → `r:id="rId7"`
  - End-of-body paragraph (`w14:paraId="47A3E9AC"`) removed in Broken
  - `w:cols w:space="720"` removed in Broken

#### 18r. Sign-off Table Borders
- **Affects:** The final table (Sign-off table)
- **Change:** Working has no `<w:tblBorders>` on this table (only `w:tblLook`). Broken adds explicit `<w:tblBorders>` with `w:val="none" w:sz="0"` for all edges, and also adds `<w:tcBorders>` with the same none/zero values to each cell
