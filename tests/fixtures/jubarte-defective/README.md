# jubarte-defective (jubarte AST-roundtrip outputs that docx-validate v0.1.3 flags)

DOCX files produced by **jubarte's `docx → ast → docx` roundtrip** that
docx-validate v0.1.3 marks **invalid**. Collected so the validator's behavior
on these stays under test. `_manifest.json` has the exact errors per file.

Excluded by design: jubarte bugs that were already **fixed** and now produce
valid output (missing-`w:author` comments; reserved-`xml:`-namespace
`ns0:space` math) — per "we need them failing", fixed ones don't belong here.

## IMPORTANT: most of these are NOT jubarte defects

Verified against **real Microsoft Word** (the `probe-word-fixtures.ts` oracle):

| file | validator | real Word | verdict |
|---|---|---|---|
| `jubarte_comments.docx` | invalid (`comment-thread-count-mismatch`) | **clean-open** | **validator false positive** |
| `jubarte_Redline_CiceroDo_v_plate(30.docx` | invalid (`comment-thread-count-mismatch`) | (threaded, same class) | **likely validator false positive** |
| `jubarte_external_..._numbering-implicit-numid.docx` | invalid (broken `Scratch.dot` ref) | n/a | source defect, preserved verbatim |
| `jubarte_sectpr-headerref.docx` | invalid (orphan `docProps/custom.xml`) | n/a | source defect, preserved verbatim |
| `jubarte_nested-comments-marker.docx` | invalid (`w:t` missing `xml:space`) | n/a | source defect, preserved verbatim |

### The validator bug these expose

`comment-thread-count-mismatch` (src/scripts/office/validators/docx.ts, rule 4)
requires `commentRangeStart/End/Reference` counts to equal the **total** number
of `<w:comment>` entries. But **reply** comments in a thread legitimately have
**no range markers of their own** — they attach to the parent's range via
`commentsExtended.xml` `w15:paraIdParent`. So any threaded document trips the
rule, even though Word opens it cleanly (`comments.docx` proven clean-open).

**Fix:** compute `expected` excluding replies — a comment whose `<w15:commentEx>`
has a non-null `w15:paraIdParent` is a reply and must not require its own range
markers. `commentsExtendedXml` is already loaded before rule 4.

This matters beyond the validator: jubarte must be **Word-first**, and these
files prove jubarte's output is Word-correct while the validator over-rejects.
Driving any "regenerate-on-invalid" logic off this rule would corrupt valid
documents until the rule is Word-aligned.
