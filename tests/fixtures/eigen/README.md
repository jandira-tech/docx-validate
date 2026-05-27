# Eigen real-world DOCX specimens

Real-world `.docx` exports from the Plate/SuperDoc ("eigenpal") editor,
imported from the former top-level `fixtures/eigen-extended/` collection (which
nothing referenced). They are kept together here, separate from the curated
synthetic fixtures in `../broken/` and `../working/`, because they share a
single near-universal, benign export quirk — `id-paraid-overflow` (and
sometimes `style-default-missing`) — rather than each demonstrating one
deliberate defect.

## Naming

Each file was renamed deterministically from its own content by
`scripts/apply-fixture-names.ts --descriptor content-first`, following
`<subject>.<descriptor>.docx`:

- **subject** — slugified `dc:title` (or first paragraph text), e.g.
  `customer-satisfaction-survey-q4`, `book-catalog`, `red-bold-text-demo`.
- **descriptor** — the file's distinguishing content feature when it has one
  (`suggesting-insertions`, `suggesting-deletions`, `suggesting-mixed-edits`,
  `table`, `comment-…`); otherwise the validation error code it trips
  (`id-paraid-overflow`, `style-default-missing`).

Exact-content duplicates (by `sha256` of `word/document.xml`) were dropped
during import; colliding names are disambiguated with a numeric suffix
(`…-2.docx`).

## Expected validator outcomes

Pinned in `tests/fixtures-all.manifest.json` and asserted by
`tests/fixtures-all-strict.test.ts` / `tests/fixtures-all-lenient.test.ts`.
See the design spec at
`docs/superpowers/specs/2026-05-27-repo-cleanup-fixture-naming-design.md`.
