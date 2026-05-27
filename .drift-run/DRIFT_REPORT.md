# Fixture repair-drift report (profile: strict)

- Fixtures processed: **410** (6 could not be processed)
- Repaired (repairs > 0): **160**
- Showed drift (any add/remove/change): **57**
- **Content loss (error-tier): 1**

## Content loss (error-tier — a content-bearing element decreased)

| Fixture | repairs | -content | drift (a/r/c) |
|---|---|---|---|
| `external/open-xml-sdk/mcdoc.docx` | 4 | 6 | 0/10/0 |

### Loss detail (first few per fixture)

**`external/open-xml-sdk/mcdoc.docx`**
- document structure 'drawing reference': 1 → 0 (element(s)).
- document structure 'DrawingML placement': 1 → 0 (element(s)).
- document structure 'paragraph': 2 → 0 (element(s)).
- document structure 'run': 2 → 0 (element(s)).
- document structure 'VML picture': 1 → 0 (element(s)).

## Non-loss drift (warn/info only — fidelity/shape/metadata changed, no content lost)

| Fixture | repairs | drift (a/r/c) | warns |
|---|---|---|---|
| `broken/header-with-rid-no-sidecar.docx` | 9 | 1/0/0 | 1 |
| `unknown/unknown/header-with-rid-no-sidecar.docx` | 22 | 1/0/0 | 1 |
| `word-regenerate-invalid/original/broken/header-with-rid-no-sidecar.docx` | 9 | 1/0/0 | 1 |
| `word-regenerate-invalid/original/unknown/unknown/header-with-rid-no-sidecar.docx` | 22 | 1/0/0 | 1 |
| `broken/empty.missing-content-type.docx` | 6 | 4/0/0 | 0 |
| `broken/endnotes.paraid-overflow.docx` | 16 | 4/0/0 | 0 |
| `broken/sample-document.broken-tables.docx` | 85 | 5/0/0 | 0 |
| `broken/sample-document.empty-rels.docx` | 103 | 5/0/0 | 0 |
| `broken/sample-document.id-overflow.docx` | 156 | 5/0/0 | 0 |
| `broken/sample-document.missing-paraids.docx` | 98 | 5/0/0 | 0 |
| `external/apache-poi/bug56075-changeTracking_on.docx` | 6 | 4/0/0 | 0 |
| `external/docx-templates/anchor-empty.docx` | 3 | 3/0/0 | 0 |
| `external/docx-templates/for-row1.docx` | 21 | 3/0/0 | 0 |
| `external/docx-templates/for1.docx` | 11 | 3/0/0 | 0 |
| `external/docx-templates/for1customDelimiter.docx` | 11 | 3/0/0 | 0 |
| `external/docx-templates/for1inline.docx` | 7 | 3/0/0 | 0 |
| `external/docx-templates/for1inlineWithSpaces.docx` | 7 | 3/0/0 | 0 |
| `external/docx-templates/for1scalars.docx` | 11 | 3/0/0 | 0 |
| `external/docx-templates/for2.docx` | 19 | 3/0/0 | 0 |
| `external/docx-templates/for3.docx` | 25 | 3/0/0 | 0 |
| `external/docx-templates/htmls.docx` | 3 | 3/0/0 | 0 |
| `external/docx-templates/if-row1.docx` | 37 | 3/0/0 | 0 |
| `external/docx-templates/if.docx` | 23 | 3/0/0 | 0 |
| `external/docx-templates/if2.docx` | 39 | 3/0/0 | 0 |
| `external/docx-templates/ifInline.docx` | 15 | 3/0/0 | 0 |
| `external/docx-templates/invalidCommand.docx` | 13 | 3/0/0 | 0 |
| `external/docx-templates/invalidIf.docx` | 15 | 3/0/0 | 0 |
| `external/docx-templates/literalXml.docx` | 7 | 3/0/0 | 0 |
| `external/docx-templates/longText.docx` | 9 | 3/0/0 | 0 |
| `external/docx-templates/missingEndIf.docx` | 3 | 3/0/0 | 0 |
| `external/docx-templates/noQuery.docx` | 3 | 3/0/0 | 0 |
| `external/docx-templates/noQuerySimpleInserts.docx` | 3 | 3/0/0 | 0 |
| `external/docx-templates/replaceTwoImages.docx` | 19 | 3/0/0 | 0 |
| `external/docx-templates/simpleQuery.docx` | 3 | 3/0/0 | 0 |
| `external/docx-templates/simpleQuerySimpleInserts.docx` | 3 | 3/0/0 | 0 |
| `external/docx-templates/splitDelimiters.docx` | 13 | 3/0/0 | 0 |
| `external/docx-templates/zipGeneration.docx` | 13 | 1/0/0 | 0 |
| `external/superdoc/behavior/sd-2517-localized-heading-styles.docx` | 11 | 3/0/0 | 0 |
| `external/superdoc/doc-api-stories/table-style-options-roundtrip.docx` | 28 | 2/0/0 | 0 |
| `external/superdoc/evals/comments-doc.docx` | 5 | 2/0/0 | 0 |
| `external/superdoc/super-editor/gdocs-comments-export.docx` | 6 | 2/0/0 | 0 |
| `external/superdoc/super-editor/gdocs-single-comment.docx` | 4 | 2/0/0 | 0 |
| `external/superdoc/super-editor/gdocs-tracked-changes.docx` | 7 | 2/0/0 | 0 |
| `external/superdoc/super-editor/Google Docs Originated comments & TCs.docx` | 21 | 2/0/0 | 0 |
| `external/superdoc/super-editor/missing-sectpr.docx` | 5 | 3/0/0 | 0 |
| `external/superdoc/super-editor/nested-comments-gdocs.docx` | 8 | 2/0/0 | 0 |
| `external/superdoc/super-editor/nested-comments.docx` | 22 | 2/0/0 | 0 |
| `external/superdoc/super-editor/pagination-blank.docx` | 3 | 2/0/0 | 0 |
| `external/superdoc/super-editor/sd-1707-list-enter-track-changes-with-br.docx` | 5 | 2/0/0 | 0 |
| `external/superdoc/super-editor/shape_group.docx` | 55 | 3/0/0 | 0 |
| `external/superdoc/super-editor/superdoc-hyperlink-cases.docx` | 35 | 3/0/0 | 0 |
| `word-regenerate-invalid/original/broken/sample-document.broken-tables.docx` | 85 | 5/0/0 | 0 |
| `word-regenerate-invalid/original/broken/sample-document.empty-rels.docx` | 103 | 5/0/0 | 0 |
| `word-regenerate-invalid/original/broken/sample-document.id-overflow.docx` | 156 | 5/0/0 | 0 |
| `word-regenerate-invalid/original/broken/sample-document.missing-paraids.docx` | 98 | 5/0/0 | 0 |
| `word-strict/sample-document.broken-tables.docx` | 85 | 5/0/0 | 0 |

## Could not process

- `external/apache-poi/crash-517626e815e0afa9decd0ebb6d1dee63fb9907dd.docx`: Corrupted zip: can't find end of central directory
- `external/superdoc/encryption/encrypted-advanced-text.docx`: Can't find end of central directory : is this a zip file ? If it is, see https://stuk.github.io/jszip/documentation/howto/read_zip.html
- `external/superdoc/encryption/encrypted-hello.docx`: Can't find end of central directory : is this a zip file ? If it is, see https://stuk.github.io/jszip/documentation/howto/read_zip.html
- `word-regenerate-invalid/original/external/apache-poi/crash-517626e815e0afa9decd0ebb6d1dee63fb9907dd.docx`: Corrupted zip: can't find end of central directory
- `word-regenerate-invalid/original/external/superdoc/encryption/encrypted-advanced-text.docx`: Can't find end of central directory : is this a zip file ? If it is, see https://stuk.github.io/jszip/documentation/howto/read_zip.html
- `word-regenerate-invalid/original/external/superdoc/encryption/encrypted-hello.docx`: Can't find end of central directory : is this a zip file ? If it is, see https://stuk.github.io/jszip/documentation/howto/read_zip.html
