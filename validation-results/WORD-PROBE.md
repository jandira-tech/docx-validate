# Word-adjudicated truth: docx-validate vs real Microsoft Word

Real MS Word (macOS, `scripts/probe-word-fixtures.ts`) is the ground-truth oracle.

## Full-corpus probe (current) — all 573 fixtures, `word-valid` profile
Raw log: `word-probe-all.jsonl`. Our validator agrees with real Word on
**567/573 = 98.95%**.

Word outcomes: 508 clean-open, 43 unreadable-content-warning, 17 open-error,
2 password-required, 2 unknown-dialog, 1 timeout.

Mismatches (6 records = ~3 distinct files, each duplicated across fixture trees):
- **2 false positives** (we reject, Word opens clean): `tiny-picture.broken-rels.docx`
  variants — broken `media/*.png` reference that Word tolerates.
- **4 false negatives** (we pass, Word rejects): `comments.unmatched-comment-marker.docx`
  ×2, `footnotes.duplicate-paraid.docx` ×2 — malformed `docProps/app.xml` integers.

These 6 are NOT safely fixable — see `broken-word/FINDINGS.md`. Word is
non-deterministic on the relevant defects (same malformed docProps field opens one
file and breaks another), so any escalation creates offsetting false positives or
false negatives. Left as documented known edge cases.

Shipped from this analysis: `ignorable-undeclared` → error under `word-valid`
(fixed `Sample Document.docx`, Word-verified, 0 false positives).

## Earlier disagreement-set probe (superseded, kept for history)
`word-probe-disagreements.jsonl` — the 200 fixtures where we and the OpenXML SDK
disagreed, probed before the full-corpus run and before the `ignorable-undeclared`
fix. Superseded by the full-corpus numbers above.
