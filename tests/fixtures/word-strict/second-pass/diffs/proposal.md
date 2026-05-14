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
- `FootnoteText` paragraph style gains `<w:spacing w:after="0" w:line="240" w:lineRule="auto"/>`

### Numbering changes
- `w:multiLevelType w:val="hybridMultilevel"` added to first abstractNum
- `w15:tentative="1"` added to list levels (replacing `w:tplc`)
- abstractNum IDs incremented: 0→1, 1→2

### Footnotes
- `mc:Ignorable="w14 w15 wp14"` attribute on root element
- `<w:spacing w:after="0" w:line="240" w:lineRule="auto"/>` on footnote paragraphs
- `<w:rPr>` with `FootnoteReference` style on footnote reference runs

### Main document additions
- `w:rFonts w:ascii="Inter" w:hAnsi="Inter" w:cs="Inter"` on approximately 77 runs
- `w:sz w:val="22"` / `w:szCs w:val="22"` on body text runs
- `w:spacing w:before="0"` on approximately 56 paragraphs
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
- Namespace declarations removed from root elements across 8 XML parts: `cx`, `cx1` through `cx8` (chartex variants), `aink`, `am3d`, `w16du`, `w16sdtdh`, `w16sdtfl`, `w16se`, and others

### Document properties emptied
- `docProps/app.xml`: all child elements removed (Template, TotalTime, Pages, Words, Characters, Application, DocSecurity, Lines, Paragraphs, ScaleCrop, Company, LinksUpToDate, CharactersWithSpaces, SharedDoc, HyperlinksChanged, AppVersion)
- `docProps/core.xml`: metadata fields reformatted inline→block; `lastModifiedBy` changed from "Arthur Souza Rodrigues" to "Un-named"; timestamps differ (created: 2026-05-14→2026-03-05; modified: 2026-05-14→2026-05-10)

### Settings removals
- `w:zoom w:percent="150"`
- `w:defaultTabStop w:val="720"`
- `w:characterSpacingControl w:val="doNotCompress"`
- `w:footnotePr`, `w:endnotePr`
- Multiple compatSettings beyond `compatibilityMode`: `overrideTableStyleFontSizeAndJustification`, `enableOpenTypeFeatures`, `doNotFlipMirrorIndents`, `differentiateMultirowTableHeaders`, `useWord2013TrackBottomHyphenation`
- `w:rsids` (rsidRoot + 4 rsid values)
- `m:mathPr`
- `w:themeFontLang`, `w:clrSchemeMapping`, `w:decimalSymbol`, `w:listSeparator`
- `w15:docId`

### Styles removals
- Entire `<w:latentStyles>` section (376 `lsdException` entries)
- `CommentText` paragraph style
- `CommentTextChar` character style
- `CommentReference` character style
- `Strong1` custom paragraph style (replaced by built-in `Strong`)
- `w:uiPriority`, `w:semiHidden`, `w:unhideWhenUsed`, `w:outlineLvl` from heading styles
- Style name lowercasing (`"heading 1"` format → `"Heading 1"` format)

### Font table
- All 5 font definitions removed (Inter, Times New Roman, Courier New, Aptos Display, Aptos)

### Numbering removals
- `w:nsid` from both abstractNums
- `w:tmpl` from both abstractNums
- `w:tplc` from list levels
- `w16cid:durableId` from num elements
- Levels 1 through 8 from second abstractNum

### Main document removals
- `w:rsidR`, `w:rsidRDefault` attributes from all paragraphs and table rows
- `w:tblLook` from all 6 tables
- `w:tblStyle w:val="Normal"` from first table
- `w:space="0"` from approximately 30 border edge elements
- `w:color="auto"` from all `<w:shd>` elements
- `<w:lastRenderedPageBreak/>` from 2 locations
- `w:rsidR="0008393D"` from `<w:sectPr>`
- `w:cols w:space="720"` from `<w:sectPr>`
- End-of-body paragraph (`w14:paraId="47A3E9AC"`) from `<w:sectPr>`

---

## 3. Proposed Solutions

### Issue 3a: Whole-file additions and removals (`word/endnotes.xml`, `word/theme/theme1.xml`, `word/webSettings.xml` vs `docProps/custom.xml`, `word/commentsExtensible.xml`)

**Observation:** The repairer adds `docProps/custom.xml` and `word/commentsExtensible.xml` but strips `word/endnotes.xml`, `word/theme/theme1.xml`, and `word/webSettings.xml`. These three files existed in the original broken document (confirmed in the first-pass unpacked-broken directory) and were retained by Word's repair. The repairer removes them entirely.

**Possible cause:** The repair path in `DOCXSchemaValidator.repair()` may regenerate parts from a minimal template rather than passing through all parts present in the input. Parts not explicitly accounted for in the repair logic are dropped.

**Proposed fix:** The repair path should preserve all parts present in the input document that are structurally valid. Only parts explicitly detected as corrupt or unrecoverable should be removed. If the repairer maintains an allowlist of part types, it should be expanded to include all standard OOXML part types.

### Issue 3b: Font table emptied

**Observation:** The repairer outputs an empty `<w:fonts>` element with zero font definitions. Word's repair retains 5 font definitions including the primary font "Inter". The repairer later adds `w:rFonts w:ascii="Inter"` to runs in the document body — referencing a font that is not defined in the font table.

**Possible cause:** The font table may be regenerated from scratch rather than preserved from input. The regeneration produces an empty table because no font-reconstruction logic exists.

**Proposed fix:** The font table from the input document should be preserved as-is during repair. If the repairer needs to add font entries (e.g., for fonts referenced in styles or runs but missing from the table), it should append to the existing table rather than replacing it entirely.

### Issue 3c: Styles stripped — latentStyles, CommentText, CommentTextChar, and CommentReference removed

**Observation:** 376 `lsdException` entries from `<w:latentStyles>` are removed. `CommentText`, `CommentTextChar`, and `CommentReference` style definitions are removed. `Strong1` custom style is replaced by built-in `Strong`.

**Possible cause:** The repairer may regenerate styles from a built-in default set rather than preserving the original style definitions. The default set includes only basic styles (Normal, headings, Title, TableNormal, NoList, ListParagraph) and drops all custom, comment, and latent style entries.

**Proposed fix:** The style repair should preserve existing style definitions from the input document. Only styles that are truly corrupt or missing should be added. If full regeneration is unavoidable, it should start from a copy of the original styles and supplement rather than replace.

### Issue 3d: Comment IDs regenerated

**Observation:** Comment IDs in `word/comments.xml` and `word/document.xml` change from `"4"` to `"100"`. Paragraph IDs and durable IDs in `word/commentsIds.xml` also change.

**Possible cause:** The repairer may be renumbering comments from a base offset rather than preserving original identifiers.

**Proposed fix:** Original comment IDs should be preserved where structurally valid. If ID conflicts or gaps exist that require renumbering, this should be treated as a specific fix rather than a blanket renumbering.

### Issue 3e: Table structural properties stripped

**Observation:** The repairer removes `w:tblLook`, `w:tblStyle`, and `w:space="0"` from tables. `w:color="auto"` is removed from cell shading attributes.

**Possible cause:** Some of these are default values (`w:space="0"`, `w:color="auto"`) and their removal is structurally harmless. However, `w:tblLook` with `w:val="04A0"` and `firstRow/firstColumn/noVBand` flags controls conditional table formatting and should not be removed.

**Proposed fix:** `w:tblLook` attributes should be preserved from the input document. Default-value attributes (`w:space="0"`, `w:color="auto"`) can be safely omitted. If the repairer is intentionally normalizing table structure, `w:tblLook` represents formatting intent and should be retained.

### Issue 3f: Tracked change IDs shifted

**Observation:** All `w:del` and `w:ins` element IDs shift by +9 or +10 offset (e.g., `w:id="0"` → `w:id="10"`, `w:id="1"` → `w:id="11"`).

**Possible cause:** The repairer may be applying a base offset to tracked change IDs, possibly to avoid collisions with other document identifiers.

**Proposed fix:** Tracked change IDs should be preserved from the original document. If renumbering is required due to conflicts, the new IDs should maintain internal consistency with any cross-referenced elements.

### Issue 3g: Explicit cell borders added (78 instances)

**Observation:** The repairer adds `<w:tcBorders>` to every table cell, with per-edge border properties. The original document relies on table-level `<w:tblBorders>` for border rendering.

**Possible cause:** This appears to be an intentional behavior in the repairer — expanding table-level borders to individual cell-level borders. This may be a normalization step for validation purposes.

**Proposed fix:** If this is an intentional normalization, it is not a defect per se. However, the repairer could consider a less invasive approach: preserve the original table-level border model when it passes validation, and only expand to per-cell borders when the borders are structurally corrupt.

### Issue 3h: Relationship ID scheme changed

**Observation:** Word uses sequential numeric relationship IDs (rId1 through rId16). The repairer uses string-based hash IDs for hyperlink relationships (e.g., `rIdclrrgvzn-co-bssfjluef`).

**Possible cause:** The repairer may generate relationship IDs using a non-sequential scheme, potentially hashing the relationship target to produce a unique identifier.

**Proposed fix:** Relationship IDs should use a consistent, deterministic scheme. Sequential numeric IDs (rId1, rId2, ...) are the most common convention in OOXML and should be preferred unless there is a specific reason to use an alternative scheme.

### Issue 3i: Redundant explicit properties (rFonts, sz, spacing, ind)

**Observation:** The repairer adds explicit `w:rFonts`, `w:sz`/`w:szCs`, `w:spacing w:before="0"`, and `w:ind` to runs and paragraphs where Word relies on style inheritance.

**Possible cause:** The repairer may be resolving style inheritance explicitly rather than relying on the style cascade. This could be because font definitions were stripped (see Issue 3b), rendering style-based font resolution unreliable.

**Proposed fix:** If the font table is preserved (see Issue 3b fix), explicit per-run font properties may not be necessary. The repairer should only add explicit formatting where the style cascade would produce incorrect results.
