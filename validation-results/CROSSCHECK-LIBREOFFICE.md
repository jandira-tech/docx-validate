# Third-validator cross-check: LibreOffice vs Word

`scripts/crosscheck-libreoffice.ts` runs the fixtures through LibreOffice headless
(`--convert-to pdf`) — a second real, independent Office engine — to see whether it
reproduces Microsoft Word's accept/reject decision. Results:
`validation-results/libreoffice-crosscheck.jsonl`.

## Result: LibreOffice is MORE lenient than Word
Of the **65** fixtures Word does **not** open cleanly, LibreOffice **opened 55** and
rejected only **10**. The 10 it rejects are the *universally corrupt* files:
malformed XML, missing-namespace, a fuzzer-corrupted zip (apache-poi crash), and the
two encrypted documents. Everything else Word rejects, LibreOffice renders fine.

## The disputed false-negatives: 3 engines side against Word
For the false-negatives where our validator and Word disagree
(`comments.unmatched-comment-marker.docx`, `footnotes.duplicate-paraid.docx`):

| Engine | Verdict | Notes |
|---|---|---|
| **Microsoft Word** | **rejects** | "unreadable content" — uniquely strict |
| LibreOffice | **opens** | renders a PDF, rc=0 |
| OpenXML SDK | does not flag the cause | XSD whitespace-collapse accepts the docProps values |
| docx-validate (libxml2) | passes | warns on docProps integers, which are NOT the real cause |

**Three independent validators agree these files are openable; only Word rejects
them.** No off-the-shelf validator (SDK, libxml2, LibreOffice) reproduces Word's
strictness on these documents.

## What that means for the residual false-negatives
The remaining false-negatives are not a gap in our validator that another tool can
close — they are Word being uniquely fussy about its own repackaging tool's
pretty-printed `docProps/app.xml` (the cause is some whitespace in that part — see
`broken-word/FINDINGS.md` — but it is XSD-valid and tolerated by every other engine).
Keeping them as documented known edge cases is the correct call; chasing a rule to
match Word's idiosyncrasy would only fire on output from that one tool.

## Note
The earlier "2 of the 5 reject for structural reasons" finding holds:
`header-with-rid-no-sidecar` and `text-box.orphan-part` reject for missing-sidecar /
orphan-part reasons (NOT docProps) and our validator already flags them — they are
aligned, not false-negatives.
