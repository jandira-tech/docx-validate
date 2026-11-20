# word-element-fixtures — one minimal valid .docx per element declared in wml.xsd

A generated corpus: for every unique element name declared in
`src/scripts/office/schemas/ISO-IEC29500-4_2016/wml.xsd` (global and local
`xsd:element` declarations, de-duplicated by name — 613 elements), this
directory holds one `.docx`, named `<element>.docx`.

Each fixture is the **absolute minimal** package that still passes validation
while remaining representative of the element's XML rule: the element is kept
and placed at the schema-correct position — across same-name schema
declarations the shortest path wins, and within one declaration the earliest
part root in the list below wins (`w:document` first, so primary-document
paths are preferred over equally long header/footer ones). Part roots, in
priority order (mirrors `ROOT_PARTS` in `scripts/repair-element-fixtures.ts`):
`w:document`, `w:glossaryDocument`, `w:styles`, `w:settings`, `w:numbering`,
`w:fonts`, `w:webSettings`, `w:comments`, `w:footnotes`, `w:endnotes`, `w:hdr`,
`w:ftr`, `w:recipients` — through the bundled XSD content-model
graph — with exactly the required child elements and required attributes the
schema demands, plus the smallest extras the semantic checks ask for (a
comments part for comment markers, a footnotes part for a footnote reference,
a relationship sidecar for an `r:id`, …). Nothing more is added.

Example — `t.docx` is just:

```xml
<w:document xmlns:w="…"><w:body><w:p><w:r><w:t>x</w:t></w:r></w:p></w:body></w:document>
```

## Manifest

`manifest.jsonl` — one row per fixture:

```json
{
    "element": "t",
    "file": "t.docx",
    "sha256": "…",
    "profile": "lenient",
    "valid": true,
    "issueCount": 0,
    "errorCodes": [],
    "warningCodes": []
}
```

- `sha256` — hash of the `.docx` bytes as generated.
- `profile` / `valid` / `issueCount` / `errorCodes` / `warningCodes` — the
  outcome of `new Validate({ profile: "lenient", suffix: ".docx" }).run(bytes)`.
  All 613 fixtures are `valid: true` under `lenient`.

## Regenerate

```bash
pnpm exec tsx scripts/generate-element-fixtures.ts   # bare-element corpus
pnpm exec tsx scripts/repair-element-fixtures.ts     # minimal valid repair (rewrites fixtures + manifest.jsonl)
```

`repair-element-fixtures.ts` builds each fixture from the schema (path search,
required-content synthesis) and validates it; it supports
`--out/--results/--profile/--limit/--only/--resume`.
