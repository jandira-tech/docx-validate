# Provable-benign repair drift + repair-validity fixes

**Date:** 2026-05-26
**Branch:** stacked on `feat/inventory-fingerprint`
**Status:** Design approved — pending implementation plan

## Problem

A repair-drift survey over 416 `.docx` fixtures (strict profile) found: 410
processed, 57 drifted after repair, of which 56 were _classified_ benign by the
severity model and 1 showed content loss (`external/open-xml-sdk/mcdoc.docx`).
Investigation revealed two defects and one trust gap:

1. **Repair emits invalid XML.** `repairMissingParaIds` re-stamps a `w14:paraId`
   on a paragraph that already has one, producing a **duplicate attribute** that
   makes `word/document.xml` unparseable and unopenable in Word. (This is what
   surfaced as "content loss" — the inventory collector skips the unparseable
   part, dropping every counter in it to zero.)
2. **`diff-docx` crashes** on malformed input (`Attribute w14:paraId redefined`)
   instead of reporting it.
3. **"Benign" is asserted, not proven.** The 56 non-loss drifters are labeled
   benign by a category-based severity classifier. We must _prove_ each is
   genuinely invisible to an ordinary end-user (the Springfield-Illinois
   reasonable-person standard) — no user-visible change to the rendered document.

## Goals

- Fix defect (1): repair must never emit a duplicate `w14:paraId`/`w14:textId`.
- Fix defect (2): `diff-docx` reports unparseable parts as a finding, never throws.
- Establish a **strict, testable invariant** — _repair must not change the
  visible projection of a document_ — and prove every non-loss drifter satisfies
  it. Any violation is a repair defect.

## Non-goals (this iteration)

- **Render-diff validation is deferred.** Rendering before/after in
  Word/LibreOffice (PDF text + per-page image hash) is the eventual ground-truth
  check that the visible-projection model faithfully matches Word. It requires a
  Word/LibreOffice machine (`SOFFICE_AVAILABLE`) and is documented as a future
  validation step, not built here.
- No change to the severity classifier or the inventory diff from the prior
  feature.

## The core abstraction: `collectVisibleProjection`

The "ordinary Word user" is operationalized as an explicit, **ordered,
normalized** projection of a document containing _only what a reader perceives on
the page_. A repair-drift is **proven benign iff the projection is unchanged**.

### INCLUDE — user-visible

- **Body block sequence** (paragraphs and tables) **in document order**.
- **Per paragraph:** final rendered text (tracked-changes _final_ view —
  insertions in, deletions out), with `<w:tab/>` → `\t` and `<w:br>` → `¶`;
  `pStyle` value; visible `pPr` — alignment (`jc`), indentation (`ind`), spacing
  (`spacing`), numbering (`numPr` → `numId`/`ilvl`), borders/shading.
- **Per run:** visible text + visible `rPr` — bold/italic/underline/strike/
  dstrike/caps/smallCaps, color, highlight, font size (`sz`), fonts (`rFonts`),
  vertical align, `rStyle`. **Hidden runs (`vanish`) contribute no text** (not
  rendered).
- **Tables:** rows × columns, merges (`gridSpan`/`vMerge`), visible cell
  properties (borders, shading, width), and recursively projected cell content.
- **Images/drawings:** extent (bucketed to nearest 1000 EMU) + the **resolved
  relationship target** (the image part name), _not_ the `rId` string.
- **Hyperlinks:** display text + **resolved target**.
- **Headers/footers:** projected content + which section references them.
- **Page geometry:** size, orientation, margins, columns, section breaks.
- **Numbering:** the number formats referenced by used `numId`/`ilvl`.
- **Comments:** anchor position + comment text (an ordinary user sees comment
  marks/bubbles).

### EXCLUDE — invisible plumbing

`w14`/`w15`/`w16cid`/`w16cex` `paraId`/`textId`/`durableId`; all `rsid*`;
relationship `Id` strings (the _target_ is captured, the id is not); content-type
declarations; namespace declarations; attribute order and the element order of
property containers; XML indentation and insignificant inter-element whitespace;
`docProps` metadata (not rendered on the page).

### Tricky-case decisions (approved)

- **Tracked changes:** projection is the **final** rendered view — accept
  insertions, drop deletions.
- **Hidden text (`w:vanish`):** excluded (not rendered).
- **`xml:space="preserve"`:** significant — preserved, because it changes visible
  spacing.
- **Comments:** included as visible.

### Determinism

The projection is a deterministic, order-preserving structure (nested arrays /
plain objects) suitable for deep-equality and for a readable delta. Insignificant
whitespace and excluded attributes are normalized away so that a
plumbing-only repair yields an **identical** projection.

## Proof procedure (strict invariant)

For each non-loss drifter (unpacked `before` = original, `after` = repaired):

```
provenBenign  ⇔  deepEqual(collectVisibleProjection(before),
                           collectVisibleProjection(after))
```

- **Empty delta → proven benign.** Recorded with the (empty) projection diff as
  evidence.
- **Non-empty delta → repair DEFECT** (per the approved "any visible change =
  bug" rule), recorded with the exact visible delta (which block/paragraph/run
  changed, before→after). These escalate as defects to fix in the repairer; this
  iteration _reports_ them rigorously (fixing each is follow-up work, except the
  already-known paraId defect which is fixed here).

The known content-loss case (`mcdoc.docx`) is expected to fail the proof until
defect (1) is fixed; after the fix its projection must be invariant.

## Defect (1) fix — duplicate `w14:paraId`

**Root cause** (`src/scripts/office/validators/docx.ts`, `repairMissingParaIds`,
~lines 2413–2432): the existence check reads `getAttributeNS(W14_NAMESPACE,
"paraId")` (namespace + local name), but the write is `setAttributeNS(
W14_NAMESPACE, "w14:paraId", …)`. In `mcdoc.docx` the paragraph's existing
`w14:paraId` is bound so the namespace-only lookup returns empty; repair then
stamps a second `w14:paraId`, yielding a duplicate **qualified-name** attribute
and invalid XML.

**Fix:** make the existence check robust to namespace-binding quirks by adding a
qualified-name fallback (mirroring the codebase's `wordChildAttr` triple-fallback
pattern):

```
hasParaId = elem.getAttributeNS(W14_NAMESPACE, "paraId") || elem.getAttribute("w14:paraId")
hasTextId = elem.getAttributeNS(W14_NAMESPACE, "textId") || elem.getAttribute("w14:textId")
```

When an existing paraId/textId is detected by _either_ form, the element falls
into the no-op case and is not re-stamped — eliminating the duplicate.

## Defect (2) fix — `diff-docx` hardening

`runDiffDocx` currently lets `collectDocxSemanticInventory` (→ `parseXml`) throw
on malformed parts. Harden so an unparseable input is reported as a finding, not
a crash: catch the error, return `{ code: 1, markdown: "<error: unparseable part
…>" }` (or an `inventory-unparseable` line), so the CLI exits non-zero with a
clear message. The same robustness should not silently hide real corruption — it
reports it.

## Architecture / modules

| File                                                       | Change | Responsibility                                                                                                                                   |
| ---------------------------------------------------------- | ------ | ------------------------------------------------------------------------------------------------------------------------------------------------ |
| `src/scripts/office/validators/docx.ts`                    | modify | Fix `repairMissingParaIds` existence check (defect 1).                                                                                           |
| `scripts/diff-docx.ts`                                     | modify | Catch unparseable parts; report instead of throw (defect 2).                                                                                     |
| `src/scripts/office/validators/docx-visible-projection.ts` | new    | `collectVisibleProjection(unpackedDir, profile)` + `diffVisibleProjections(before, after)` returning a structured, readable delta.               |
| `scripts/prove-benign.ts`                                  | new    | Over a `.drift-run`-style copies/repaired tree (or fresh from fixtures): run projection-invariance per non-loss drifter, emit `BENIGN_PROOF.md`. |

`collectVisibleProjection` reuses `src/lib/xml-helpers.ts` and the same
namespace/`directWordChild`/`wordChildAttr` helper style as the inventory; it is
a separate module (distinct concern: perceptible-content projection vs counter
inventory).

## Output

`BENIGN_PROOF.md`:

- Aggregate: N proven benign (empty projection delta), M defects (non-empty
  delta), with the strict invariant stated.
- Proven-benign table: fixture + "visible projection identical".
- Defect table: fixture + the visible delta (block path, before → after).

## Testing (TDD red-green)

- **Defect 1:** a `mcdoc`-style paragraph that already carries `w14:paraId`
  (bound such that the namespace-only lookup misses it) → after `repairParaId`,
  the paragraph has exactly **one** `w14:paraId` and `document.xml` re-parses.
  Red first (reproduce the duplicate), then fix.
- **Defect 2:** `runDiffDocx` on an input with a malformed part returns a
  non-zero `code` and a report string; does not throw.
- **Projection:**
    - A plumbing-only change (add `w14:paraId`, reorder attributes, re-declare a
      namespace, reflow whitespace) → `diffVisibleProjections` empty.
    - A visible change (delete a paragraph's text, drop bold, resize a table,
      swap an image target) → non-empty delta naming the change.
    - Hidden (`vanish`) text contributes no projected text; `xml:space="preserve"`
      spacing is retained.
- **prove-benign:** a tiny before/after pair with plumbing-only drift → proven
  benign; a pair with a visible change → defect with delta.

## Delivery (stacked commits)

1. **Commit 1** — defect (1) fix + regression test (`tests/validators-docx.test.ts`).
2. **Commit 2** — defect (2) `diff-docx` hardening + test.
3. **Commit 3** — `collectVisibleProjection` + `diffVisibleProjections` + unit tests.
4. **Commit 4** — `prove-benign.ts`; run it over the 56 non-loss drifters; commit
   `BENIGN_PROOF.md`; reclassify any failures as defects in the report.

## Deferred / future (documented, not built)

- `scripts/render-diff.ts`: `SOFFICE_AVAILABLE`-gated render-diff (PDF text +
  per-page image hash) to validate the projection model against real Word/
  LibreOffice output. Run on a Word machine; if render-diff and projection ever
  disagree, the projection model has a gap to close.
- Fixing any _new_ repair defects surfaced by `prove-benign` (beyond paraId) is
  follow-up work, one fix per defect, each TDD.
