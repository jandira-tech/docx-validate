# Proposal: Observed Differences Between Word-Repaired and Repairer Output

**Compared files:**
- **Working (Word-repaired):** `second-pass-word-repaired-sample-document.broken-tables.docx`
- **Broken (repairer output):** `second-pass-actually-broken-sample-document.broken-tables.docx`

---

## 1. What Was Added (present in repairer output, absent in Word-repaired)

### Whole-file additions
- `docProps/custom.xml` — custom document properties part
- `word/commentsExtensible.xml` — comments extensible part

### Namespace additions
- `wne` namespace (`http://schemas.microsoft.com/office/word/2006/wordml`) added to root elements of `word/footnotes.xml`, `word/numbering.xml`, `word/settings.xml`, `word/document.xml`
- Additional namespaces on `word/settings.xml`: `wpc`, `wp14`, `wp`, `wpg`, `wpi`, `wne`, `wps`

### Content Type additions
- Default entries for: `image/png`, `image/jpeg`, `image/jpg`, `image/bmp`, `image/gif`, `image/svg+xml`, `application/vnd.openxmlformats-officedocument.obfuscatedFont`
- Override entries for: `/docProps/core.xml`, `/docProps/custom.xml`, `/docProps/app.xml`, `/word/fontTable.xml`, `/word/header1.xml`, `/word/footer1.xml`, `/word/commentsExtensible.xml`

### Package relationships
- `rId4` relationship added for `custom-properties` targeting `docProps/custom.xml`

### Document relationships
- `commentsExtensible` relationship added
- Hyperlink relationship IDs changed from sequential numeric (`rId7`, `rId8`, `rId12`) to hash-based string IDs

### Settings additions
- `w:evenAndOddHeaders w:val="false"`

### Styles changes
- `CommentReference` character style present (needed by comment anchor runs)

### Numbering changes
- `w:multiLevelType w:val="hybridMultilevel"` added to first abstractNum
- `w15:tentative="1"` added to list levels (replacing `w:tplc`)
- abstractNum IDs incremented: 0→1, 1→2

### Footnotes
- `mc:Ignorable="w14 w15 wp14"` attribute on root element
- `<w:spacing w:after="0" w:line="240" w:lineRule="auto"/>` on footnote paragraphs
- `<w:rPr>` with `FootnoteReference` style on footnote reference runs

### Main document additions
- `w:rFonts w:ascii="Inter" w:hAnsi="Inter" w:cs="Inter"` on ~77 runs
- `w:sz w:val="22"` / `w:szCs w:val="22"` on body text runs
- `w:spacing w:before="0"` on ~56 paragraphs
- `<w:tcBorders>` on 78 table cells (with per-edge border specifications)
- `<w:tblBorders>` with `w:val="none" w:sz="0"` on sign-off table
- `<w:ind w:left="640" w:hanging="320"/>` on 4 list paragraphs
- Comment ID changed: `w:id="4"` → `w:id="100"` (in document body and comments.xml)
- Tracked change IDs shifted: +9/+10 offset
- Table row paragraph/text IDs regenerated (`w14:paraId`, `w14:textId`)

---

## 2. What Was Removed (present in Word-repaired, absent in repairer output)

### Whole-file removals
- `word/endnotes.xml` — endnotes part
- `word/theme/theme1.xml` — office theme
- `word/webSettings.xml` — web settings

### Content Type removals
- Override entries for: `/word/webSettings.xml`, `/word/endnotes.xml`, `/word/theme/theme1.xml`

### XML declaration attribute removal
- `standalone="yes"` removed from XML declarations in 12 files

### Namespace removals
- Namespace declarations removed from root elements across 8 XML parts: `cx`, `cx1`-`cx8` (chartex), `aink`, `am3d`, `w16du`, `w16sdtdh`, `w16sdtfl`, `w16se`, and others

### Document properties emptied
- `docProps/app.xml`: all child elements removed (Template, TotalTime, Pages, Words, Characters, Application, etc.)
- `docProps/core.xml`: metadata fields reformatted inline→block; `lastModifiedBy` changed from "Arthur Souza Rodrigues" to "Un-named"; timestamps differ

### Settings removals
- `w:zoom w:percent="150"`
- `w:defaultTabStop w:val="720"`
- `w:characterSpacingControl w:val="doNotCompress"`
- `w:footnotePr`, `w:endnotePr`
- Multiple compatSettings beyond `compatibilityMode`
- `w:rsids` (rsidRoot + 4 rsid values)
- `m:mathPr`
- `w:themeFontLang`, `w:clrSchemeMapping`, `w:decimalSymbol`, `w:listSeparator`
- `w15:docId`

### Styles removals
- Entire `<w:latentStyles>` section (376 `lsdException` entries)
- `CommentText` paragraph style
- `CommentTextChar` character style
- `CommentReference` character style
- `Strong1` custom paragraph style (replaced by `Strong` built-in style)
- `w:uiPriority`, `w:semiHidden`, `w:unhideWhenUsed`, `w:outlineLvl` from heading styles
- Style name lowercasing (`"heading 1"` → `"Heading 1"`)

### Font table
- All 5 font definitions removed (Inter, Times New Roman, Courier New, Aptos Display, Aptos)

### Numbering removals
- `w:nsid` from both abstractNums
- `w:tmpl` from both abstractNums
- `w:tplc` from list levels
- `w16cid:durableId` from num elements
- Levels 1-8 from second abstractNum

### Main document removals
- `w:rsidR`, `w:rsidRDefault` attributes from all paragraphs and table rows
- `w:tblLook` from all 6 tables
- `w:tblStyle w:val="Normal"` from first table
- `w:space="0"` from ~30 border edge elements
- `w:color="auto"` from all `<w:shd>` elements
- `<w:lastRenderedPageBreak/>` from 2 locations
- `w:rsidR="0008393D"` from `<w:sectPr>`
- `w:cols w:space="720"` from `<w:sectPr>`
- End-of-body paragraph (`w14:paraId="47A3E9AC"`) from `<w:sectPr>`

---

## 3. Proposed Solutions

### Issue 3a: Whole-file additions/removals (`word/endnotes.xml`, `word/theme/theme1.xml`, `word/webSettings.xml` vs `docProps/custom.xml`, `word/commentsExtensible.xml`)

**Observation:** The repairer adds `docProps/custom.xml` and `word/commentsExtensible.xml` but strips `word/endnotes.xml`, `word/theme/theme1.xml`, and `word/webSettings.xml`. These three files existed in the original broken document (found in the first-pass `unpacked-broken` directory) and were retained by Word's repair. The repairer removes them entirely.

**Possible cause:** The repairer may not preserve parts that pass through a filter or regeneration step where only a subset of known part types are retained.

**Proposed fix:** The repair path (in `DOCXSchemaValidator.repair()`) should preserve all parts present in the input document that are valid OOXML, even if they are not currently "needed" for the repair logic. Only parts explicitly detected as corrupt should be removed.

### Issue 3b: Font table emptied

**Observation:** The repairer outputs an empty `<w:fonts>` element with zero font definitions. Working has 5 font definitions including the primary font "Inter".

**Possible cause:** During repair, font definitions may be correctly identified by validation but the rewrite step may fail to serialize them back. The validator may detect missing fonts but the repair step produces an empty font table rather than one populated from the original input.

**Proposed fix:** Investigate the font reconstruction logic in the repair path. If the original document has font definitions in `word/fontTable.xml`, the repair should carry them forward. If fonts need to be added for styles referenced in the document, they should be appended rather than replacing the entire table.

### Issue 3c: Styles stripped — latentStyles and custom styles removed

**Observation:** 376 `lsdException` entries from `<w:latentStyles>` are removed, along with `CommentText`, `CommentTextChar`, and `CommentReference` style definitions. `Strong1` is replaced by built-in `Strong`.

**Possible cause:** The repairer may regenerate the style definitions from a minimal template rather than preserving and augmenting the original styles. This minimal template includes only the styles the repairer "knows about" (Title, Headings, Normal, TableNormal, NoList, etc.) and drops all others.

**Proposed fix:** The style repair should preserve existing style definitions from the input and only add missing ones. If style regeneration is necessary, it should start from a copy of the original styles and supplement, not replace.

### Issue 3d: Comment IDs regenerated

**Observation:** Comment IDs in `word/comments.xml` and `word/document.xml` change from `"4"` to `"100"`. Comment paragraph/durable IDs also change.

**Possible cause:** The repairer may renumber comments during repair, starting from a base offset (100) rather than preserving original IDs.

**Proposed fix:** Preserve original comment IDs where they are valid. If comments must be renumbered (e.g., due to conflicts), document the rationale.

### Issue 3e: Table structural properties stripped

**Observation:** `w:tblLook`, `w:tblStyle`, and `w:space="0"` are removed from tables. `w:color="auto"` is removed from cell shading.

**Possible cause:** The repairer may strip attributes that it considers redundant with default values. `w:space="0"` is the default for table border spacing. `w:color="auto"` is the default foreground color. `w:tblLook` and `w:tblStyle` may be removed as part of a table normalization step.

**Proposed fix:** Default-value attributes (`w:space="0"`, `w:color="auto"`) are harmless to remove — they have no visual effect. However, `w:tblLook` (particularly `w:val="04A0"` with `firstRow/firstColumn/noVBand`) defines conditional formatting behavior and should not be stripped. The repair should preserve `w:tblLook` from the input.

### Issue 3f: Tracked change IDs shifted

**Observation:** All `w:del` and `w:ins` element IDs are shifted by +9 or +10.

**Possible cause:** The repairer may prepend or insert additional tracked changes, causing a renumbering. Or it may apply a base offset to avoid conflicts with other IDs used in the document.

**Proposed fix:** Tracked change IDs should be preserved from the original unless conflicts require renumbering. If renumbering is needed, cross-reference with the corresponding `w:commentReference` IDs to maintain internal consistency.

### Issue 3g: Explicit cell borders added (78 instances)

**Observation:** The repairer adds `<w:tcBorders>` to every table cell, with border properties inferred from the table-level `<w:tblBorders>`.

**Possible cause:** This may be an intentional normalization step — converting table-level borders to cell-level borders for consistency or to work around validation issues.

**Proposed fix:** If cell-level borders are the repairer's intentional output format, this is not a defect. However, the repairer could consider preserving the original table-level border model when it passes validation, and only expanding to per-cell borders when the table-level borders are corrupt.

### Issue 3h: Relationship ID scheme changed

**Observation:** Working uses sequential numeric relationship IDs. Broken uses string-based hash IDs for hyperlinks.

**Possible cause:** The repairer may regenerate relationship IDs using a different scheme (concatenated string hashes) rather than preserving or rebuilding sequential numeric IDs.

**Proposed fix:** Relationship IDs should use a consistent, deterministic scheme. The sequential numeric scheme (rId1, rId2, ...) is the most common in OOXML documents and should be preferred.
