# Plan 04: Comment ID Stability

## Problem

The repairer renumbers `w:id` on `<w:comment>` elements. In the second-pass
diff:

- Working: `<w:comment w:id="4" ...>` (original Word-assigned ID)
- Broken: `<w:comment w:id="100" ...>` (repairer-assigned ID)

The corresponding `<w:commentRangeStart>`, `<w:commentRangeEnd>`, and
`<w:commentReference>` elements in `word/document.xml` must all use the same
`w:id`. If the repairer renumbers the comment but does not update all
references, the comment markers become orphaned and Word warns or silently
discards the comment.

Additionally, the comment para identity chain across extension files is broken:

| File | Attribute | Working value | Broken value |
|------|-----------|---------------|--------------|
| `comments.xml` | `w14:paraId` on `<w:p>` inside comment | `456E2E6B` | `456E2E6B` (same) |
| `commentsExtended.xml` | `w15:paraId` | `456E2E6B` | `456E2E6B` (same) |
| `commentsIds.xml` | `w16cid:paraId` | `456E2E6B` | `B68569E0` (**different**) |
| `commentsExtensible.xml` | `w16cex:durableId` | (absent in working) | `8A13236F` (**orphaned**) |

This mismatch is now detected by:
- `comment-thread-commentid-paraid-orphan` — `commentsIds.xml` paraId not
  found in any `<w:comment>` paragraph
- `comment-thread-durableid-orphan` — `commentsExtensible.xml` durableId not
  found in `commentsIds.xml`

Both are fatal under all profiles including `word-valid` (confirmed in
`tests/validate.test.ts`: "word-valid profile keeps commentsIds/commentsExtensible
mismatches fatal").

## Current detection

`validateCommentThreading()` in `DOCXSchemaValidator`
(`src/scripts/office/validators/docx.ts`) — added in this PR — checks the
full comment identity chain across all four comment XML parts.

## Proposed fix

### Repairer side

**Preserve comment IDs**:

1. Do not renumber `w:id` on `<w:comment>` elements. The existing ID is a
   stable cross-reference key used by `<w:commentRangeStart>`,
   `<w:commentRangeEnd>`, and `<w:commentReference>` in `document.xml`.

2. When the repairer must create a new comment (not repair an existing one),
   assign IDs that do not collide with existing IDs:
   ```typescript
   const existingIds = new Set(
     comments.map(c => parseInt(c.getAttribute("w:id") ?? "0", 10))
   );
   const nextId = Math.max(0, ...existingIds) + 1;
   ```

3. **Para-ID chain consistency**: When the repairer copies or rewrites
   `commentsIds.xml`, it must use the `w14:paraId` value from the `<w:p>`
   inside the corresponding `<w:comment>`, not generate a new random value.
   The mapping is:
   ```
   comments.xml  →  w:comment[w:id=X]/w:p[w14:paraId=Y]
   commentsIds.xml → w16cid:commentId[w16cid:paraId=Y]
   ```

4. If `commentsExtensible.xml` is preserved (copy-through per Plan 01), its
   `w16cex:durableId` values must match the `w16cid:durableId` values in
   `commentsIds.xml` for the same comment. If the repairer cannot ensure this,
   drop `commentsExtensible.xml` entirely (the working sample does this).

### Validator side (already implemented)

The `validateCommentThreading()` method added in this PR covers:
- `comment-thread-commentid-paraid-orphan`
- `comment-thread-durableid-orphan`
- `comment-thread-count-mismatch`
- `paraid-missing-element`

No additional validator changes needed beyond what this PR adds.

## Acceptance criteria

- After repairer round-trip, `comments.xml` retains original `w:id` values.
- `commentsIds.xml` has `w16cid:paraId` matching the `w14:paraId` of the
  comment paragraph in `comments.xml`.
- `commentsExtensible.xml` is either absent or has `durableId` values that
  exist in `commentsIds.xml`.
- `validate(output, { profile: "lenient" })` returns no
  `comment-thread-commentid-paraid-orphan` or `comment-thread-durableid-orphan`
  errors.
- Existing test `"word-valid profile keeps commentsIds/commentsExtensible
  mismatches fatal"` in `tests/validate.test.ts` continues to pass against
  the broken fixture and the fixed output passes.

## Test fixture reference

- Comments diff: `tests/fixtures/word-strict/second-pass/diffs/word_comments.xml.diff`
- CommentsIds diff: `tests/fixtures/word-strict/second-pass/diffs/word_commentsIds.xml.diff`
- Schema matrix: `tests/fixtures/word-strict/second-pass/diffs/SCHEMA_MATRIX.md`