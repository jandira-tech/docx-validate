# word-element-invalid-fixtures — one minimally-broken .docx per element declared in wml.xsd

A generated corpus: a copy of every `word-element-fixtures/<element>.docx`
(613 elements), each mutated by the **absolute minimum** so that it **fails the
validation rule specific to that element** — while still containing the element
in its proper context, i.e. remaining representative of the element's XML rule.

The mutation is applied to the target element itself, in this priority order
(the first one whose lenient run is invalid with an error naming the element
`{…}name` is accepted):

1. `remove-first-attribute` — drop one of the element's (required) attributes
   → “attribute … is required but missing” on the element.
2. `remove-first-child` — drop one required child element
   → “Missing child element(s)” on the element.
3. `add-undeclared-attribute` — add `w:bogus="1"`
   → “attribute … is not allowed/declared” on the element.

All 613 fixtures are `valid: false` under the `lenient` profile, and every one
of them has an `xsd-error` whose message names the element itself.

## Manifest

`manifest.jsonl` — one row per fixture:

```json
{
    "element": "tbl",
    "file": "tbl.docx",
    "sha256": "…",
    "profile": "lenient",
    "valid": false,
    "mutation": "remove-first-child",
    "issueCount": 2,
    "errorCodes": ["xsd-error", "xsd-summary"],
    "errorMessages": ["…"]
}
```

## Regenerate

```bash
pnpm exec tsx scripts/break-element-fixtures.ts   # reads word-element-fixtures/, writes this dir + manifest.jsonl
```

Supports `--src/--out/--results/--profile/--limit/--only/--resume`.
