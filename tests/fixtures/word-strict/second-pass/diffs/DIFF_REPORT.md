# DIFF REPORT: Word-Repaired vs Repairer Output

Comparison of unpacked XML from:

- **Working** (Word-repaired): `second-pass-word-repaired-sample-document.broken-tables.docx`
  → `tests/fixtures/word-strict/second-pass/unpacked-working/`
- **Broken** (repairer output): `second-pass-actually-broken-sample-document.broken-tables.docx`
  → `tests/fixtures/word-strict/second-pass/unpacked-broken/`

All 16 common XML files differ. 2 files are broken-only; 3 files are working-only.

---

## Files only in BROKEN

| File | Note |
|------|------|
| `docProps/custom.xml` | Custom properties part added by repairer |
| `word/commentsExtensible.xml` | Comment extensible part with durableId mismatch |

## Files only in WORKING

| File | Note |
|------|------|
| `word/endnotes.xml` | Endnotes part dropped by repairer |
| `word/theme/theme1.xml` | Theme part dropped by repairer |
| `word/webSettings.xml` | Web settings part dropped by repairer |

---

## Group 1 — XML Declaration

Both sides use `<?xml version="1.0" encoding="UTF-8"?>`. Working adds
`standalone="yes"` on several parts; broken omits it.

Files affected: `word/document.xml`, `_rels/.rels`, `[Content_Types].xml`,
`word/styles.xml`, and others.

## Group 2 — Namespace Reductions

Repairer removes unused namespace declarations from root elements. Examples:

- `word/commentsExtended.xml`: removes `xmlns:am3d`, `xmlns:v`, others;
  updates `mc:Ignorable` accordingly.
- `word/footer1.xml`, `word/header1.xml`: removes multiple `xmlns:*`
  declarations, produces a shorter root element.

Files affected: `word/commentsExtended.xml`, `word/footer1.xml`,
`word/header1.xml`, `word/footnotes.xml`, `word/fontTable.xml`.

## Group 3 — Package Relationships (`_rels/.rels`)

Working: 3 relationships (`rId1`–`rId3`): office document, core props,
app props.

Broken: 4 relationships (`rId1`–`rId4`): adds custom-properties relationship
targeting `docProps/custom.xml`. Relationship IDs renumbered.

## Group 4 — Content Types (`[Content_Types].xml`)

Working: overrides for document, styles, settings, webSettings, numbering,
footnotes, endnotes, comments, commentsExtended, commentsIds, header, footer,
theme, core props, app props.

Broken: adds image defaults (png, jpeg, jpg, bmp, gif, svg); removes endnotes,
webSettings, and theme overrides; adds commentsExtensible; removes standalone
from XML declaration.

## Group 5 — Document Properties (`docProps/app.xml`)

Working: full `<Properties>` element with metadata (TotalTime, Pages, Words,
Characters, Lines, Paragraphs, AppVersion, etc.).

Broken: self-closing empty `<Properties/>` element — all metadata dropped.

## Group 6 — Core Properties (`docProps/core.xml`)

Working: `cp:lastModifiedBy` = `Arthur Souza Rodrigues`, modified timestamp
`2026-05-14T01:27:00Z`.

Broken: `cp:lastModifiedBy` = `Un-named`, modified timestamp
`2026-05-10T00:07:14.347Z`. Both set `cp:revision` = `1`.

## Group 7 — Document Relationships (`word/_rels/document.xml.rels`)

Working: relationships include theme, numbering, endnotes, footnotes,
commentsIds, header, footer, settings, styles, comments, commentsExtended,
webSettings, and external hyperlinks. Uses non-sequential IDs.

Broken: relationships renumbered to `rId1`–`rIdN`; removes theme, endnotes,
webSettings; adds commentsExtensible relationship.

## Group 8 — Comments (`word/comments.xml`)

Working: `<w:comment w:id="4" …>` (original ID), date `2026-03-05T19:41:00Z`,
paragraph has `w14:textId` attribute.

Broken: `<w:comment w:id="100" …>` (renumbered), date format different, no
`w14:textId` on paragraph.

## Group 9 — Comments Extended (`word/commentsExtended.xml`)

Namespace declarations differ. Both contain `<w15:commentEx w15:paraId="456E2E6B"
w15:done="0"/>`. The `mc:Ignorable` attribute differs due to namespace changes.

## Group 10 — Comments IDs (`word/commentsIds.xml`)

Working: `<w16cid:commentId w16cid:paraId="456E2E6B" w16cid:durableId="456E2E6B"/>`.
(Both attributes equal to the comment paragraph's `w14:paraId`.)

Broken: `<w16cid:commentId w16cid:paraId="B68569E0" w16cid:durableId="05B4A74E"/>`.
(Different values — not aligned with `456E2E6B` from `comments.xml`.)

**This mismatch causes `comment-thread-commentid-paraid-orphan` errors.**

## Group 11 — Header (`word/header1.xml`)

Working: full namespace set on `<w:hdr>`, paragraph has border and alignment
attributes, run has no explicit font.

Broken: reduced namespace set, `<w:spacing>` added to paragraph properties,
`<w:rFonts w:ascii="Inter" …>` injected into run.

## Group 12 — Footer (`word/footer1.xml`)

Similar to header: broken version injects `<w:rFonts>Inter</w:rFonts>` into
every run. Text content split across multiple XML lines.

## Group 13 — Footnotes (`word/footnotes.xml`)

Working: separator footnotes use empty `<w:separator/>` runs with no paragraph
properties.

Broken: separator footnotes have `FootnoteReference`-styled runs with
`<w:spacing>` in paragraph properties.

## Group 14 — Font Table (`word/fontTable.xml`)

Working: 5 font entries — Inter, Times New Roman, Courier New, Aptos Display,
Aptos — each with `w:panose1`, `w:charset`, `w:family`, `w:pitch`, `w:sig`.

Broken: empty `<w:fonts>` element. All font entries removed.

**Fidelity regression: all fonts fall back to Word default.**

## Group 15 — Settings (`word/settings.xml`)

Working: full settings including zoom, tabs, footnote/endnote props, rsids,
math properties, theme font language, color scheme, decimal/list separators,
`w15:docId`.

Broken: minimal settings with only `<w:evenAndOddHeaders>` and a
`<w:compatSetting>`. All other settings removed.

## Group 16 — Numbering (`word/numbering.xml`)

Working: 2 `abstractNum` definitions (IDs `0` and `1`), using `w:tmpl` and
per-level `w:tplc` attributes. `<w:num>` elements have `w16cid:durableId`.

Broken: 2 `abstractNum` definitions (IDs `1` and `2`), `w:tmpl` and `w:tplc`
removed, levels marked `w15:tentative="1"`. `<w:num>` elements lack
`w16cid:durableId`.

## Group 17 — Styles (`word/styles.xml`)

Working: full styles document including `<w:latentStyles>` (~150 entries),
all heading styles with `uiPriority` and outline level, comment-related styles,
custom styles.

Broken: regenerated minimal set — Normal, DefaultParagraphFont, TableNormal,
NoList, Title, Heading1–6 (renamed, stripped of uiPriority/outline level).
No latent styles. No comment styles. No custom styles.

**Causes `style-default-missing` errors for CommentReference, CommentText,
CommentSubject.**

## Group 18 — Main Document (`word/document.xml`)

18 sub-groups of differences:

| Sub-group | Description |
|-----------|-------------|
| 18a | rsid attributes removed from all paragraphs and runs |
| 18b | `w14:textId` removed from paragraphs that had it |
| 18c | Explicit `<w:rFonts>` injected on runs that inherited font from style |
| 18d | Explicit `<w:sz>` and `<w:szCs>` injected on runs |
| 18e | Explicit `<w:spacing>` injected on paragraphs |
| 18f | `<w:tblLook>` removed from `<w:tblPr>` on tables |
| 18g | `<w:tblCellSpacing>` removed from `<w:tblPr>` |
| 18h | `<w:jc>` removed from `<w:tblPr>` |
| 18i | `<w:tblInd>` removed from `<w:tblPr>` |
| 18j | `<w:shd>` removed from cell properties |
| 18k | `<w:tcBorders>` added to every cell (was inherited from style) |
| 18l | Tracked-change `w:id` values regenerated from 1 |
| 18m | Comment-range `w:id` values changed (4→100 on commentRangeStart/End) |
| 18n | Hyperlink `r:id` values changed due to relationship renumbering |
| 18o | `w:numId` references updated to match renumbered `<w:num>` elements |
| 18p | Whitespace reformatting within XML text nodes |
| 18q | `w:color w:val="auto"` removed from some runs |
| 18r | `<w:lang>` removed from default run properties in styles |