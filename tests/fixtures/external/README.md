# External DOCX Test Fixtures

This directory holds DOCX test fixtures imported from third-party open-source
projects. They supplement the smaller intentionally-broken corpus in
`../broken/` with real-world malformed and edge-case files used by other
OOXML toolchains' own test suites.

Files retain upstream filenames where possible (with explicit notes for any
renames in the table below) so they can be cross-referenced.
Per-vendor `NOTICE` / `LICENSE` files document the attribution requirements.

Some files in this set are **negative-tests**: the validator correctly returns
`ok=true` for them because the OPC/OOXML spec permits the construct in
question. They are kept to assert that the validator does not over-validate
(i.e. reject valid-but-unusual documents).

## Provenance

| File                                                             | Source repo            | Upstream license | Why it's here                                                                                                                                                                         |
| ---------------------------------------------------------------- | ---------------------- | ---------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `apache-poi/bug59378.docx`                                       | apache/poi             | Apache-2.0       | OPC compliance edge case: `Default Extension="rels"` correctly covers `_rels/.rels`. Validator (correctly) accepts. Kept as a negative-test asserting we don't over-validate.         |
| `apache-poi/MultipleBodyBug.docx`                                | apache/poi             | Apache-2.0       | Multiple `<w:body>` elements — structural violation                                                                                                                                   |
| `apache-poi/51921-Word-Crash067.docx`                            | apache/poi             | Apache-2.0       | Known parser-crash input — hardening                                                                                                                                                  |
| `apache-poi/crash-517626e815e0afa9decd0ebb6d1dee63fb9907dd.docx` | apache/poi             | Apache-2.0       | Fuzzer-derived crash input — hardening                                                                                                                                                |
| `apache-poi/bug56075-changeTracking_on.docx`                     | apache/poi             | Apache-2.0       | Tracked-changes corpus                                                                                                                                                                |
| `open-xml-sdk/EmptyRelationshipElement.docx`                     | dotnet/Open-XML-SDK    | MIT              | Empty `<Relationship>` tag — broken-rels                                                                                                                                              |
| `open-xml-sdk/5Errors.docx`                                      | dotnet/Open-XML-SDK    | MIT              | Five intentional schema validation errors                                                                                                                                             |
| `open-xml-sdk/InvalidDocProps.docx`                              | dotnet/Open-XML-SDK    | MIT              | Malformed `docProps` part                                                                                                                                                             |
| `open-xml-sdk/InvalidDocPropsct.docx`                            | dotnet/Open-XML-SDK    | MIT              | Content-type mismatch on docProps                                                                                                                                                     |
| `open-xml-sdk/UnknownElement.docx`                               | dotnet/Open-XML-SDK    | MIT              | Unknown XML element under mc:ignorable                                                                                                                                                |
| `open-xml-sdk/mcdoc.docx`                                        | dotnet/Open-XML-SDK    | MIT              | Markup Compatibility / AlternateContent                                                                                                                                               |
| `open-xml-sdk/Strict01.docx`                                     | dotnet/Open-XML-SDK    | MIT              | ISO OOXML Strict conformance class                                                                                                                                                    |
| `docx4j/header-no-rels.docx`                                     | plutext/docx4j         | Apache-2.0       | Header part with no `.rels` sidecar. Per OPC §9.3.1, parts with no outgoing r:id refs don't require sidecars. Validator (correctly) accepts. Kept as a negative-test.                 |
| `docx4j/hyperlink_dupe.docx`                                     | plutext/docx4j         | Apache-2.0       | Upstream filename suggests duplicate hyperlink rIds, but the fixture has zero r:id refs in document.xml and no hyperlink Relationships. Validator (correctly) accepts. Negative-test. |
| `docx4j/NumberingImplicitNumId.docx`                             | plutext/docx4j         | Apache-2.0       | Numbering with implicit numId                                                                                                                                                         |
| `mammoth-js/strict-format.docx`                                  | mwilliamson/mammoth.js | BSD-2-Clause     | ISO OOXML Strict conformance class, single body paragraph. Used by `tests/strict-namespace-paragraph-count.test.ts`.                                                                  |
| `mammoth-js/text-box.docx`                                       | mwilliamson/mammoth.js | BSD-2-Clause     | DOCX with a text box; inner paragraphs must not be counted as body paragraphs. Used by `tests/textbox-paragraph-count.test.ts`.                                                       |

## Excluded sources

- **pandoc** (jgm/pandoc) — GPL-2+ license on the repo. Test files cannot be
  imported without legal review.
- **SheetJS/libreoffice_test-files** — license on the mirror is unconfirmed
  and the underlying provenance of some LibreOffice test files is unclear.

## Source inventory

The full survey that produced this set lives at
`tmp/external-fixtures-inventory.md` (gitignored).
