# Proposal: Repairer Fidelity Fixes (Second-Pass Analysis)

Comparison of Word-repaired vs. repairer output for
`sample-document.broken-tables.docx`. This document is descriptive only;
implementation plans are under `docs/plans/conejo/`.

---

## Section 1 — What the Repairer Added

| Category | Details |
|----------|---------|
| New files | `docProps/custom.xml`, `word/commentsExtensible.xml` |
| Namespace reductions | Removed unused `xmlns:am3d`, `xmlns:v`, and others from root elements |
| Explicit run fonts | `<w:rFonts w:ascii="Inter" …/>` injected on header/footer runs |
| Explicit paragraph spacing | `<w:spacing w:line="276" w:lineRule="auto"/>` added to paragraph properties |
| Explicit cell borders | Full `<w:tcBorders>` block added to every table cell |
| Comment renumbering | `w:id` on `<w:comment>` changed from `4` to `100` |
| New content types | Image defaults (`png`, `jpeg`, `jpg`, `bmp`, `gif`, `svg`) added to `[Content_Types].xml` |
| Relationship additions | `custom-properties` relationship added to `_rels/.rels` |
| Relationship renumbering | All `<Relationship Id="…">` reassigned from `rId1` sequentially |

---

## Section 2 — What the Repairer Removed

| Category | Details |
|----------|---------|
| Whole parts removed | `word/endnotes.xml`, `word/theme/theme1.xml`, `word/webSettings.xml` |
| Font table emptied | All `<w:font>` entries removed from `word/fontTable.xml` |
| `<w:latentStyles>` block | Entire block (~150 lsdException entries) removed from `word/styles.xml` |
| Comment-related styles | `CommentText`, `CommentSubject`, `CommentReference` removed from styles |
| Table structural properties | `<w:tblLook>`, `<w:tblCellSpacing>`, `<w:jc>`, `<w:tblInd>`, `<w:shd>` removed |
| Settings content | Most of `word/settings.xml` removed (rsids, math props, docId, etc.) |
| Revision session IDs | `w:rsid*` attributes stripped from all paragraphs and runs |
| Tracked-change IDs | `w:id` on `<w:ins>` and `<w:del>` regenerated from 1 |
| Numbering durable IDs | `w16cid:durableId` removed from `<w:num>` in `word/numbering.xml` |
| Heading style metadata | `uiPriority` and outline level removed from heading styles |

---

## Section 3 — Proposed Solutions

### Issue 1 — Whole-file preservation

**Problem**: Repairer drops valid OOXML parts (`word/endnotes.xml`,
`word/theme/theme1.xml`, `word/webSettings.xml`) that it does not process.

**Proposed fix**: Apply a copy-through strategy for all parts the repairer
does not explicitly modify. Preserve the corresponding `<Relationship>` and
`<Override>` entries.

See: `docs/plans/conejo/01-whole-file-preservation.md`

---

### Issue 2 — Font table retention

**Problem**: Repairer replaces `word/fontTable.xml` with an empty `<w:fonts>`
document, discarding all font metadata.

**Proposed fix**: Merge input fonts into the output font table rather than
replacing. For each `<w:font>` in the input not present in the output (by
`w:name`), append it verbatim.

See: `docs/plans/conejo/02-font-table-retention.md`

---

### Issue 3 — Style passthrough

**Problem**: Repairer regenerates `word/styles.xml` from a minimal built-in
set, dropping `<w:latentStyles>`, custom styles, and comment-related styles.
This triggers `style-default-missing` errors.

**Proposed fix**: Copy `word/styles.xml` from input verbatim and run only
the `repairMissingStyleDefinitions()` pass to add any truly missing required
defaults. Never remove the latent styles block or custom style definitions.

See: `docs/plans/conejo/03-style-passthrough.md`

---

### Issue 4 — Comment ID stability

**Problem**: Repairer renumbers `w:id` on `<w:comment>` elements and generates
a mismatched `w16cid:paraId` in `word/commentsIds.xml`. The mismatch causes
`comment-thread-commentid-paraid-orphan` and `comment-thread-durableid-orphan`
errors (both fatal under all profiles including `word-valid`).

**Proposed fix**: Preserve original comment `w:id` values. Ensure
`commentsIds.xml` uses the `w14:paraId` from the comment's internal `<w:p>`,
not a newly generated value. Either align `commentsExtensible.xml` or drop it.

See: `docs/plans/conejo/04-comment-id-stability.md`

---

### Issue 5 — tblLook preservation

**Problem**: Repairer removes `<w:tblLook>` from `<w:tblPr>` on tables,
preventing conditional formatting bands from applying. Also removes
`<w:tblCellSpacing>`, `<w:jc>`, `<w:tblInd>`, and `<w:shd>`.

**Proposed fix**: Preserve all `<w:tblPr>` children verbatim unless the
repairer explicitly modifies that specific element. Never remove `<w:tblLook>`
from a table that had one.

See: `docs/plans/conejo/05-tbllook-preservation.md`

---

### Issue 6 — Tracked-change ID stability

**Problem**: Repairer regenerates `w:id` on `<w:ins>` and `<w:del>` elements
and removes `w:rsid*` attributes and `w16cid:durableId` from numbering.

**Proposed fix**: Preserve existing tracked-change `w:id` values; allocate
new IDs only above the existing maximum. Preserve `w:rsid*` attributes and
numbering `w16cid:durableId` values.

See: `docs/plans/conejo/06-tracked-change-id-stability.md`

---

### Issue 7 — Cell border normalization

**Problem**: Repairer injects explicit `<w:tcBorders>` on every cell, even
when the borders are fully inherited from the table style. This overrides
conditional formatting suppression from `<w:tblLook>`.

**Proposed fix**: Do not inject explicit cell borders when the values match
the table-level defaults. If the input had no explicit `<w:tcBorders>`, the
output should also have none.

See: `docs/plans/conejo/07-cell-border-normalization.md`

---

### Issue 8 — Relationship ID scheme

**Problem**: Repairer renumbers all relationship IDs sequentially from `rId1`.
If any reference in the XML parts is not updated to match, `rels-broken`
errors occur. Even when consistently updated, renumbering is unnecessary.

**Proposed fix**: Preserve existing relationship IDs for unchanged targets.
Assign new IDs (above existing maximum) only for relationships introduced by
the repairer itself.

See: `docs/plans/conejo/08-relationship-id-scheme.md`

---

### Issue 9 — Redundant explicit properties

**Problem**: Repairer injects explicit `<w:rFonts>`, `<w:spacing>`, `<w:color>`,
and `<w:sz>` on runs and paragraphs where these values are already provided by
the paragraph/character style. This breaks re-styling and inflates file size.

**Proposed fix**: Before writing an explicit run property, resolve the value
from the style chain. If the explicit value matches the style-inherited default,
omit the explicit property.

See: `docs/plans/conejo/09-redundant-explicit-properties.md`
