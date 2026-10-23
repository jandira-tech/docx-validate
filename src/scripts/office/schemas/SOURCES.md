# Bundled schema provenance

Every XSD under this directory is vendored so validation runs fully offline.
This file records where each _newly added_ subtree came from. The pre-existing
`ecma/`, `ISO-IEC29500-4_2016/`, `mce/` and `microsoft/` trees predate this
record and are not re-listed here.

## `iso-strict/` — ISO/IEC 29500 Strict conformance class schemas

- Source: `OfficeOpenXML-XMLSchema-Strict.zip`, extracted from
  `ECMA-376-1_5th_edition_december_2016.zip` downloaded from
  `https://ecma-international.org/wp-content/uploads/ECMA-376-1_5th_edition_december_2016.zip`
- Edition: ECMA-376 Part 1, 5th edition (December 2016), ISO/IEC 29500-1:2016
  Strict conformance class.
- Licence/terms: (c) Ecma International. Distributed with the standard for
  implementation use; see the ECMA-376 Part 1 front matter and
  <https://ecma-international.org/publications-and-standards/>.
- Local edits: `wml.xsd` and `shared-math.xsd` import the XML namespace
  without a schemaLocation; the import now points at `xml.xsd`, copied
  verbatim from `ISO-IEC29500-4_2016/xml.xsd` (the W3C XML namespace schema),
  so the xml:space attribute ref resolves offline. `wml.xsd` also had three
  `default="off"` attribute defaults that are invalid against its own
  boolean-only `ST_OnOff` union; they are patched to `default="false"`
  (same boolean value) because libxml2 refuses to compile the schema
  otherwise. `xml.xsd` is bundled verbatim.
- Files (sha256 of the file as bundled, i.e. after the edits noted above):

| file                                                 | sha256                                                           |
| ---------------------------------------------------- | ---------------------------------------------------------------- |
| iso-strict/dml-chart.xsd                             | 189b9631129a7d6c9f971b42b94c8942d260f11f81615f427db059bbf442cdc3 |
| iso-strict/dml-chartDrawing.xsd                      | 1943abd37970b25b355890f183c19a3b1a9b67a1420633030eb50ed5b278bec7 |
| iso-strict/dml-diagram.xsd                           | 179cd2fa4b981e1e2faa9ce99bb18c5d4d896dc1edb144af50f5cc8c52eb8569 |
| iso-strict/dml-lockedCanvas.xsd                      | 9c3038dd6ccc3472378be0e7c4e3e34013ee7a5619555e7463dd6d7e1d9bde1e |
| iso-strict/dml-main.xsd                              | 078871240cff323ddd22a395281421b737fa61508fd39ffaa2947ba2bc01175e |
| iso-strict/dml-picture.xsd                           | d36c91b21f80ec504cd05d40737e7e2104713d8ed7c6c8af51e5c6b778fd667c |
| iso-strict/dml-spreadsheetDrawing.xsd                | 1135493946a4d86773f1421e3453fa94bdebe575f6746e0b09a716360b44b925 |
| iso-strict/dml-wordprocessingDrawing.xsd             | 725d64911076b34f1b2d60004cd14866fc0a399ab9808dc09415197a9e4dd144 |
| iso-strict/pml.xsd                                   | 63f991936d4b11ecdab09442043d2806f1b19383e6c49b13c64186bcf485fb0b |
| iso-strict/shared-additionalCharacteristics.xsd      | 2642f009105af888c2374aa550593d7bccc0a953579a4ff70b9a2fbbb76757b2 |
| iso-strict/shared-bibliography.xsd                   | 380653d1a4e9760717c33c4d643a042d7f22d070ef012fac51d3fced3ffe36dd |
| iso-strict/shared-commonSimpleTypes.xsd              | 02df434395eb22d0e13aa2f6445b7f7fbecd3687c6dff238823bbdd3bd9cdf6a |
| iso-strict/shared-customXmlDataProperties.xsd        | e5679dbc139f9e192b7754863401f6a429002c10d8189ed68237c4972080125c |
| iso-strict/shared-customXmlSchemaProperties.xsd      | 32034ef469e73047751ca00f88c4b6e62dd47a7a655b72daf7f95a1c5015f228 |
| iso-strict/shared-documentPropertiesCustom.xsd       | e04e42fb42404c6e9025fe235b67b0f726b5735d62ac799ec4dca0fa67e99b04 |
| iso-strict/shared-documentPropertiesExtended.xsd     | 62fbc970cd9f08faad142e59085c3656b5248fd97039d9bc35d24c31f032fb62 |
| iso-strict/shared-documentPropertiesVariantTypes.xsd | 105b7556e4d40ff808410764b4e31193fceb0f5b73191517c45166e0403d8cf6 |
| iso-strict/shared-math.xsd                           | ab77ec926879b8577f0bc26a523eb07159556fc34cecc281d68aeac2722c0ad9 |
| iso-strict/shared-relationshipReference.xsd          | 82409aacec09c8672eec2f8d7e44bec10c450649f31603659a8229b122a01163 |
| iso-strict/sml.xsd                                   | dc4b61faf2b62d2caf875faf40bb14a8f3c0cf3dc52763a2db394a28e79634b4 |
| iso-strict/wml.xsd                                   | 5519f32b847180883ad572cba62203fdf0885d43d7005c9d7e7ecee1afbcee8e |
| iso-strict/xml.xsd                                   | a539aa2fb154fa50e0f5cc97e6ad7cbc66f8ec3e3746f61ec6a8b0d5d15ecdf2 |

## `dublincore/` — Dublin Core schemas imported by `ecma/fourth-edition/opc-coreProperties.xsd`

- Source: Dublin Core Metadata Initiative,
  `http://dublincore.org/schemas/xmls/qdc/2003/04/02/` — the exact release the
  OPC core-properties schema references.
- Edition: DCMI XML schemas, 2003-04-02 release.
- Licence/terms: (c) Dublin Core Metadata Initiative; published for unrestricted
  implementation use (see <https://www.dublincore.org/specifications/>).
- Local edits: `dc.xsd` and `dcterms.xsd` had their
  `schemaLocation="http://www.w3.org/2001/03/xml.xsd"` rewritten to
  `../ISO-IEC29500-4_2016/xml.xsd` so the import closure stays on disk.
  `dcmitype.xsd` is bundled verbatim. `opc-coreProperties.xsd` was likewise
  rewritten to import `../../dublincore/dc.xsd` and `../../dublincore/dcterms.xsd`.
- Files (sha256 of the file as bundled, i.e. after the rewrite noted above):

| file                    | sha256                                                           |
| ----------------------- | ---------------------------------------------------------------- |
| dublincore/dc.xsd       | 0211363702998068fd633ef97376c53a855c9ed36ce9c68d6d2c38c8f530f676 |
| dublincore/dcmitype.xsd | 4980845d0bedcf45e707fa64eb47492ae4f95ac4ae42f5c3c68b58c0bc9e8e49 |
| dublincore/dcterms.xsd  | 6d3f15ca58d6882419efaa94f2fee9d8fc97c10beffd43783a68139a06d6cd6b |
