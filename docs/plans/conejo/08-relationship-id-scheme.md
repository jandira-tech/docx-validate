# Plan 08: Relationship ID Scheme

## Problem

The repairer reassigns `Id` attributes on `<Relationship>` elements in both
`_rels/.rels` and `word/_rels/document.xml.rels`, replacing the original IDs
with a fresh sequential set starting at `rId1`.

The diff shows:
- Working `_rels/.rels`: `rId1`, `rId2`, `rId3` (3 relationships)
- Broken `_rels/.rels`: `rId1`, `rId2`, `rId3`, `rId4` (4 relationships,
  including a new `custom-properties` entry)
- Working `word/_rels/document.xml.rels`: IDs include non-sequential values
  such as `rId9`, `rId10`, etc.
- Broken `word/_rels/document.xml.rels`: IDs renumbered to `rId1`, `rId2`, …

The problem: `<Relationship Id="…">` values are cross-referenced from the XML
parts themselves. For example, `<w:hyperlink r:id="rId5">` in `document.xml`
must match the `Id="rId5"` in `document.xml.rels`. If the repairer renumbers
the relationships but fails to update all referencing attributes, hyperlinks
and other targets become broken (`rels-broken` error).

## Current detection

The existing `validateRelationships()` check in
`src/scripts/office/validators/base.ts` (inferred from the `rels-broken` code)
detects references where the relationship ID does not exist in the `.rels`
file. However, it fires based on the referenced `r:id` in the document not
matching any relationship ID, which requires consistent updates on both sides.

A repairer that renumbers consistently (updates both the `.rels` file and all
references in the XML parts) avoids the `rels-broken` error, but still causes
unnecessary diff noise and can break external tools that hold relationship IDs
as stable references.

## Proposed fix

### Repairer side

**Preserve existing relationship IDs**:

1. Parse the input `.rels` file.
2. Build a map `{ target → id }` from the input.
3. When writing the output `.rels`, reuse the existing `Id` for each
   relationship whose `Target` is unchanged.
4. For new relationships introduced by the repairer, allocate IDs that are
   above the existing maximum:
   ```typescript
   const existingIds = Array.from(rels.values()).map(id =>
     parseInt(id.replace(/^rId/, ""), 10)
   ).filter(n => !isNaN(n));
   let nextId = Math.max(0, ...existingIds) + 1;
   ```

5. Apply the same strategy to `word/_rels/document.xml.rels`.

**Rationale**: Relationship IDs are effectively a lookup key. There is no
OOXML requirement that they be sequential or start at `rId1`; any unique
string value is legal per the OPC specification. Preserving existing IDs
costs nothing and eliminates an entire category of potential reference breaks.

### Validator side (new check, severity: info)

Add `validateRelationshipIdStability()` to `DOCXSchemaValidator`:

1. Read `word/_rels/document.xml.rels` and `_rels/.rels`.
2. Collect all `Id` values.
3. If all IDs match the pattern `rId{n}` where n is sequential from 1 with no
   gaps, emit `rel-ids-sequential` at `"info"` severity — heuristic signal of
   ID regeneration.

This is `"info"` only (not blocking) since sequential IDs are valid.

**Code location**: new method in `DOCXSchemaValidator` in
`src/scripts/office/validators/docx.ts`, or in `BaseSchemaValidator` in
`src/scripts/office/validators/base.ts` if it should apply to all package
types.

## Acceptance criteria

- After repairer round-trip, relationship `Id` values in `.rels` files match
  the input values for unchanged relationships.
- New relationships added by the repairer have IDs above the input maximum.
- `validate(output, { profile: "lenient" })` returns no `rels-broken` errors.
- Hyperlinks in `word/document.xml` that referenced `r:id="rIdX"` in the
  input still resolve correctly in the output.

## Test fixture reference

- Package rels diff: `tests/fixtures/word-strict/second-pass/diffs/_rels_.rels.diff`
- Document rels diff: `tests/fixtures/word-strict/second-pass/diffs/word__rels_document.xml.rels.diff`