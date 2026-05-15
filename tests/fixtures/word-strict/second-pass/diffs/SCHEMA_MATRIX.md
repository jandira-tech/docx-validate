# Schema Matrix: Broken-Tables Word Warning

Goal: determine whether the Word "unreadable content" warning in the two
broken-table pairs is diagnosable by XML syntax/XSD alone, or whether the
validator needs Word-parity checks beyond schema validation.

## Pair Outcomes

| Pair        | Fixture                                                                              | Word outcome | XML/XSD outcome                    | Updated validator outcome                                                              |
| ----------- | ------------------------------------------------------------------------------------ | ------------ | ---------------------------------- | -------------------------------------------------------------------------------------- |
| second-pass | `unpacked-broken` / `second-pass-actually-broken-sample-document.broken-tables.docx` | warns        | every mapped XML part is XSD-valid | fails with `comment-thread-commentid-paraid-orphan`, `comment-thread-durableid-orphan` |
| second-pass | `unpacked-working` / `second-pass-word-repaired-sample-document.broken-tables.docx`  | clean        | every mapped XML part is XSD-valid | passes                                                                                 |
| first-pass  | `tests/fixtures/broken/sample-document.broken-tables.docx`                           | warns        | every mapped XML part is XSD-valid | fails with comment-thread codes plus existing id/style/table-row checks                |
| first-pass  | `tests/fixtures/working/sample-document.afterword-repaired.docx`                     | clean        | every mapped XML part is XSD-valid | passes                                                                                 |

## Relevant Schema Equivalents

| OOXML surface                                                          | XSD equivalent                         | XSD result                       | What the XSD enforces                                                    | What it does not express                                                      |
| ---------------------------------------------------------------------- | -------------------------------------- | -------------------------------- | ------------------------------------------------------------------------ | ----------------------------------------------------------------------------- |
| `word/comments.xml`                                                    | `ISO-IEC29500-4_2016/wml.xsd`          | broken and repaired valid        | local WML element/attribute structure                                    | whether extension parts point at this comment's paragraph id                  |
| `word/commentsExtended.xml`                                            | `microsoft/wml-2012.xsd`               | broken and repaired valid        | `w15:commentsEx` contains `w15:commentEx` entries with `w15:paraId`      | whether each `w15:paraId` agrees with `commentsIds.xml`                       |
| `word/commentsIds.xml`                                                 | `microsoft/wml-cid-2016.xsd`           | broken and repaired valid        | each `w16cid:commentId` has `paraId` and `durableId` long-hex attributes | whether `paraId` exists in `comments.xml`                                     |
| `word/commentsExtensible.xml`                                          | `microsoft/wml-cex-2018.xsd`           | broken valid; absent in repaired | each `w16cex:commentExtensible` has a `durableId` and optional metadata  | whether `durableId` exists in `commentsIds.xml`                               |
| namespace declarations/order on extension roots                        | no useful XSD equivalent               | broken and repaired valid        | namespace bindings are syntactically legal                               | declaration order, unused namespace declarations, and Word-specific tolerance |
| `[Content_Types].xml` and `.rels` entries for comments extension parts | OPC content-types / relationships XSDs | broken and repaired valid        | local package manifest/relationship shape                                | semantic validity of the target extension content                             |

## Isolated Finding

The second-pass broken sample has internally inconsistent comment-extension
identity data:

- `comments.xml` comment paragraph: `w14:paraId="456E2E6B"`
- `commentsExtended.xml`: `w15:commentEx w15:paraId="456E2E6B"`
- `commentsIds.xml`: `w16cid:commentId w16cid:paraId="B68569E0" w16cid:durableId="05B4A74E"`
- `commentsExtensible.xml`: `w16cex:commentExtensible w16cex:durableId="8A13236F"`

The Word-repaired second-pass sample removes `commentsExtensible.xml` and keeps
`commentsIds.xml` aligned with `456E2E6B`.

## Schema Substitution

`src-to-update-schemas/scripts/office/schemas/` was substituted into
`src/scripts/office/schemas/` and then compared with `diff -qr`; there are no
byte-level differences in the schema tree after substitution. The schema set is
therefore not sufficient by itself for these two warning pairs.
