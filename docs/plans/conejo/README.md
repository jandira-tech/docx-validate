# Plan: Conejo — Repairer Fidelity Improvements

Systematic implementation plan derived from the second-pass diff analysis
(`tests/fixtures/word-strict/second-pass/`), comparing a Word-repaired DOCX
against the repairer output.

## Background

A side-by-side comparison of:

- **Working** (Word-repaired): `second-pass-word-repaired-sample-document.broken-tables.docx`
- **Broken** (repairer output): `second-pass-actually-broken-sample-document.broken-tables.docx`

revealed 9 categories of repairer behavior that cause Microsoft Word to show
an "unreadable content" warning or otherwise degrade the output. The
`word-valid` validation profile added in this PR detects several of these
at runtime (see `src/scripts/office/validators/docx.ts`); the remaining items
require changes to the repairer pipeline itself.

## Issue Index

| # | Plan | Validator detection code | Status |
|---|------|--------------------------|--------|
| 1 | [Whole-file preservation](./01-whole-file-preservation.md) | none currently | planned |
| 2 | [Font table retention](./02-font-table-retention.md) | none currently | planned |
| 3 | [Style passthrough](./03-style-passthrough.md) | `style-default-missing` (partial) | planned |
| 4 | [Comment ID stability](./04-comment-id-stability.md) | `comment-thread-commentid-paraid-orphan`, `comment-thread-durableid-orphan` | planned |
| 5 | [tblLook preservation](./05-tbllook-preservation.md) | none currently | planned |
| 6 | [Tracked-change ID stability](./06-tracked-change-id-stability.md) | none currently | planned |
| 7 | [Cell border normalization](./07-cell-border-normalization.md) | none currently | planned |
| 8 | [Relationship ID scheme](./08-relationship-id-scheme.md) | none currently | planned |
| 9 | [Redundant explicit properties](./09-redundant-explicit-properties.md) | none currently | planned |

## Relevant source files

- `src/scripts/office/validators/docx.ts` — DOCX validator (2039 lines)
- `src/scripts/office/validators/base.ts` — base validator with XSD validation
- `src/scripts/office/validate.ts` — CLI + `validate()` entry point
- `src/lib/types.ts` — `Profile`, `ValidationResult`, `ValidationIssue` types
- `tests/fixtures/word-strict/second-pass/` — fixture corpus for these issues
- `tests/validators-docx.test.ts` — unit tests for the DOCX validator
- `tests/validate.test.ts` — integration tests including `word-valid` profile