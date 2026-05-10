# SuperDoc Fixtures

Mirror of the `.docx` fixtures from
[Harbour-Enterprises/SuperDoc](https://github.com/Harbour-Enterprises/SuperDoc) (`main` branch).

These were downloaded for cross-validator coverage — SuperDoc is an in-browser
DOCX editor with its own export pipeline, so its fixtures exercise XML shapes
the upstream ECMA/ISO toolchains never produce.

## Layout

| Subdirectory | Source path in SuperDoc | Count |
| --- | --- | --- |
| `super-editor/` | `packages/super-editor/src/editors/v1/tests/data/` | 143 |
| `super-editor/diffing/` | `…/tests/data/diffing/` | 26 |
| `behavior/` | `tests/behavior/{fixtures,tests}/**` | 36 |
| `evals/` | `evals/fixtures/docs/` | 10 |
| `doc-api-stories/` | `tests/doc-api-stories/tests/**/fixtures/` | 5 |
| `cli-legacy/` | `apps/cli/src/__tests__/fixtures-cli-legacy/` | 2 |
| `encryption/` | `…/v1/core/ooxml-encryption/fixtures/` | 2 |
| `layout-engine/shapes/` | `packages/layout-engine/test-fixtures/shapes/` | 2 |

Total: **226** files. Tests dispatch them through `runValidators` and snapshot
which ones pass / fail XSD + redlining checks.

## Expected behaviour

These are real-world documents (and a couple of *intentionally* malformed
ones — see `broken-list.docx`, `broken-complex-list.docx`, etc.). It is normal
for some to fail validation. The CI assertion is on the *shape* of the failure
set, not blanket success — see `tests/fixtures-superdoc.test.ts` for the
manifest of expected failures and the rationale for each.
