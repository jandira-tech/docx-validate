# broken-word: files Microsoft Word rejects but docx-validate (word-valid) passed

Source of truth = real MS Word (macOS probe). These 3 fixtures gave Word
"unreadable content" yet our `word-valid` profile said valid (false negatives).
`reference/` = pristine copies (do not edit). `playground/` = experiment copies.

## Proven root causes (each confirmed by changing ONE thing and re-opening in Word)

1. comments.unmatched-comment-marker.docx (unknown/unknown + word-regenerate/original)
   - NOT the orphan comment markers. Word opens the file CLEAN with
     `commentRangeStart w:id="999999"` and all comment mismatches still present.
   - Real cause: malformed `docProps/app.xml` — pretty-print whitespace polluted
     integer values, e.g. `<TotalTime>3\n  </TotalTime>`, invalid for `xs:int`.
     We already detect this as `xsd-error` (but only `warning` under word-valid).
   - Proof: `playground/comments.appxml-fixed.PROVEN-clean-in-word.docx`
     (only app.xml whitespace collapsed) → Word clean-open.

2. Sample Document.docx (eigen-extended)
   - NOT the 104 overflowing w14:textId values (they stay; Word opens clean).
   - Real cause: `mc:Ignorable="w14 w15 wp14"` references undeclared prefixes
     (only `w14` is declared). We detect this as `ignorable-undeclared`.
   - Proof: `playground/sample.ignorable-fixed.PROVEN-clean-in-word.docx`
     (mc:Ignorable pruned to declared-only) → Word clean-open.

## Why "escalate the comment codes" is the wrong fix
- comment-orphan-start / comment-marker-missing / comment-thread-count-mismatch
  appear on Word-CLEAN files too (broken/ and word-regenerate/regenerated copies).
  Escalating them → false positives, and misses the real cause.

## Safe, evidence-backed fix (word-valid profile only)
- A) `ignorable-undeclared` → error. Safe: 0 Word-clean files carry it.
- B) `xsd-error` → error ONLY when the message is an invalid atomic/simple-type
     VALUE ("is not a valid value of the atomic type") on a docProps part.
     Distinguishes the Word-breaker from the tolerated docProps `NamespaceError`
     (annotations_import_2.docx, Word-clean). The ISO-pedantry false positives
     (hint=cs, zoom@percent, sdt ordering) are on word/*.xml, not docProps values.
