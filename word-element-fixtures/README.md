# word-element-fixtures — one .docx per element declared in wml.xsd

A generated corpus: for every unique element name declared in
`src/scripts/office/schemas/ISO-IEC29500-4_2016/wml.xsd` (global and local
`xsd:element` declarations, de-duplicated by name — 613 elements), this
directory holds one minimal `.docx`, named `<element>.docx`.

Each fixture is a three-part package — `[Content_Types].xml`, `_rels/.rels`,
`word/document.xml` — whose body holds the bare element:

```xml
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body><w:ELEMENT/></w:body>
</w:document>
```

The element sits directly under `w:body`, so block-level elements validate
cleanly while run- and property-level ones draw schema findings; the verdict
per fixture is recorded, not forced.

## Manifest

`manifest.jsonl` — one row per fixture:

```json
{
    "element": "bookmarkStart",
    "file": "bookmarkStart.docx",
    "sha256": "…",
    "profile": "lenient",
    "valid": true,
    "issueCount": 0,
    "errorCodes": [],
    "warningCodes": []
}
```

- `sha256` — hash of the `.docx` bytes as generated.
- `profile` — the validation profile the row records (`lenient`; the `word-valid`
  profile of `.docx` runs no XSD walk, and `strict` validates these
  Transitional parts against the same `ISO-IEC29500-4_2016/wml.xsd`).
- `valid` / `issueCount` / `errorCodes` / `warningCodes` — the outcome of
  `new Validate({ profile: "lenient", suffix: ".docx" }).run(bytes)`.

## Regenerate

```bash
pnpm exec tsx scripts/generate-element-fixtures.ts
```

writes the fixtures into this directory and rewrites `manifest.jsonl`.
