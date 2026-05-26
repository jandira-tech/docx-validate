# docx-validate vs @ooxml-tools/validate (Microsoft OpenXML SDK)

Baseline = OpenXML SDK `Microsoft365` (its default; closest to current Word).
573 fixtures. 18 are unreadable packages (encrypted/corrupt/broken-zip) → no schema
verdict from either side → excluded from the matrix. 555 comparable.

## File-level valid/invalid matrix

| | SDK valid | SDK invalid |
|---|--:|--:|
| **ours valid** | 290 | 14 (false negatives) |
| **ours invalid** | 33 (false positives) | 218 |

Agreement: 508/555 = **91.5%**.

## (1) Are we detecting correctly?

Yes at the verdict level, but the *mechanism* differs and that matters:
- Of the 218 "both invalid", only **61 (28%)** were flagged by our XSD layer. The other
  **157** rest on our **custom semantic checks** (`id-paraid-overflow`, duplicate paraId,
  orphan comments, …), NOT on XSD.
- That is a strength, not luck: the ISO XSD types `w:id`/paraId as **unbounded integer**,
  so pure-XSD can't catch the int32 overflow the SDK reports as `Sch_AttributeUnionFailedEx`.
  Our bespoke checks catch that same real defect the XSD cannot express.

## (2) False positives — 33 files (we reject, SDK accepts)

Pervasive lint does NOT cause these: `rel-ids-sequential` (554 files) is **info**;
`paraid-missing-element`, `style-default-missing`, etc. are **warning**. None flip `valid`.

The 33 are driven by error-severity codes:
- **`xsd-error` (24 files)** — our schema is the *literal ISO/IEC 29500-4:2016 Transitional XSD*,
  which is **stricter than what Word actually emits and the SDK accepts**. Confirmed cases:
  - `w:rFonts/@w:hint="cs"` rejected — ISO `ST_Hint` = {default, eastAsia}; Word/SDK also allow `cs`.
  - `w:zoom` "requires `@w:percent`" — ISO `CT_Zoom/@percent use="required"`; Word omits it.
  - `w:sdtPr` child ordering (`w:alias`) — ISO sequence stricter than Word.
  - ~13 of the 24 are **real, valid SuperDoc documents** → genuine false positives.
  - The rest are intentionally-broken fixtures (orphan parts, missing content-type, overflow) →
    we are arguably *right* and the SDK is just lenient.
- `ws-missing-preserve`, `rels-*`, `comment-thread-*`, `ct-undeclared-ext` — our structural
  checks. Correct on broken fixtures; over-strict on a few real docs (`ws-missing-preserve`).

## (3) False negatives — 14 files (we accept, SDK rejects)

- **2** are the documented Strict-conformance skip (`Strict01.docx` 788 errs, `strict-format.docx`
  58 errs) — we skip XSD entirely on Strict-namespace docs (`xsd-strict-skipped`). Big blind spot.
- **12** are real schema issues neither our XSD nor our custom checks catch:
  - `w:ins/@w:id` int-overflow (`1778776725749`) — our overflow check covers paraId/textId, not
    revision-mark ids. (T09)
  - `w:highlight` ordering in `numbering.xml` rPr (`Sch_InvalidElementContentExpectingComplex`).
  - `m:oMathPara` placement under `w:body`.
  - Microsoft-extension part `w16cex:commentsExtensible` incomplete content — we don't validate it.
  - VML `v:oval/@ID` under `mc:Fallback` — we don't validate VML attributes.

## Bottom line

Our verdict matches Microsoft's reference validator on 91.5% of files. The disagreements are
**systematic, not random**: false positives come from the ISO XSD being stricter than Word; false
negatives come from (a) the Strict skip and (b) gaps in content-model ordering, MS-extension parts,
VML, and revision-id ranges. Our custom semantic checks add real coverage XSD can't (id overflow,
comment/relationship integrity).
