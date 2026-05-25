# Plan 01: Whole-File Preservation

## Problem

The repairer discards entire OOXML parts that are present in the input but that
it does not explicitly handle. The second-pass diff confirms three parts are
silently dropped:

- `word/endnotes.xml` — present in working, absent in repairer output
- `word/theme/theme1.xml` — present in working, absent in repairer output
- `word/webSettings.xml` — present in working, absent in repairer output

When these parts are referenced from `word/_rels/document.xml.rels` but the
files are missing, Word shows an "unreadable content" warning.

The repairer also drops the corresponding `<Override>` entries from
`[Content_Types].xml` and the `<Relationship>` entries from
`word/_rels/document.xml.rels`, compounding the issue.

## Current detection

No dedicated validator check exists. The `rels-broken` issue code in
`src/scripts/office/validators/base.ts` fires when a relationship target is
missing, but only if the relationship entry itself is retained. Because the
repairer removes both the file and its relationship entry, the broken-target
check never fires.

## Proposed fix

### Repairer side

1. **Copy-through strategy**: After unpacking the input `.docx`, copy any
   part whose path matches the allowlist below directly to the output, without
   further processing, unless the repairer explicitly needs to modify it.

   Allowlist (parts to preserve verbatim unless the repairer modifies them):
   - `word/endnotes.xml`
   - `word/theme/theme1.xml`
   - `word/webSettings.xml`
   - `word/glossary/` subtree
   - `word/charts/` subtree
   - `word/embeddings/` subtree
   - `docProps/` subtree (except `docProps/core.xml` which the repairer may
     update for author/timestamp)

2. **Relationship preservation**: When a part is preserved via copy-through,
   its relationship entry in `word/_rels/document.xml.rels` and its
   `<Override>` in `[Content_Types].xml` must also be preserved verbatim.

3. **Validation gate**: After the repairer run, call `validate()` with
   `profile: "word-valid"` and fail if `rels-broken` issues appear.

### Validator side (new check)

Add `validateOrphanedRelationships()` to `DOCXSchemaValidator`
(`src/scripts/office/validators/docx.ts`) that:

1. Reads every `.rels` file.
2. For each internal `<Relationship>` whose `Target` does not start with
   `http://` or `https://`, resolves the target path relative to the `.rels`
   file location.
3. Reports `rels-target-missing` (error) for any resolved path that does not
   exist in the unpacked directory.

This is narrower than the existing `rels-broken` check in `base.ts` and fires
even when the relationship entry exists but the file was dropped.

## Acceptance criteria

- `tests/fixtures/word-strict/second-pass/unpacked-working/word/endnotes.xml`,
  `word/theme/theme1.xml`, and `word/webSettings.xml` all survive a repairer
  round-trip.
- `validate(output, { profile: "word-valid" })` returns `valid: true` for the
  repaired second-pass fixture.
- New unit test in `tests/validators-docx.test.ts` under
  `describe("validateOrphanedRelationships")` covering the missing-target case.

## Test fixture reference

- Working unpacked: `tests/fixtures/word-strict/second-pass/unpacked-working/`
- Broken unpacked: `tests/fixtures/word-strict/second-pass/unpacked-broken/`
- Diff: `tests/fixtures/word-strict/second-pass/diffs/_only_in_working.txt`