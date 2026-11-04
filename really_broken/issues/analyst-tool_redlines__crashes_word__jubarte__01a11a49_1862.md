# [jubarte 1862] failed: sectPrChange cached w:pgMar omits required w:gutter

## Summary

`tool_redlines/crashes_word/jubarte__01a11a49_1862.docx` was a jubarte-produced tracked-changes
redline (`w:author="Redline"`, `w:date="1970-01-01T00:00:00Z"`, per the analyst's quoted XML).
The analyst reported exactly one violation: the cached pre-change `w:sectPr` jubarte wrote inside
a `w:sectPrChange` (`w:id="918"`) carries a `w:pgMar` with only six of the seven attributes
`CT_PageMar` requires — `w:gutter` is missing. The live `w:pgMar` on the same section was complete.

**This session disputes the violation as filed — not because the rule is wrong, but because the
specimen is gone and the site can no longer be read.** The live Word check reports outcome
`failed`: `really_broken/tool_redlines/` contains only `repair_prompt/` (no `crashes_word/`
directory), a repo-wide `find … -name "*1862*"` returns nothing but a coincidental scratch hash,
the unpacked copy `really_broken/unpacked/analyst-tool_redlines__crashes_word__jubarte__01a11a49_1862`
does not exist, and `really_broken/manifest.csv` does not exist either. Only the analyst's quoted
XML and the proven-cause records survive. The schema rule itself is real, and the exact reported
shape was re-proven on the sibling jubarte specimen (see Verification).

## Word behaviour

- **Live result (this ask):** never opened — outcome `failed`, the file does not exist. Re-verified
  first-hand: `ls really_broken/tool_redlines/` shows only `repair_prompt/`; the analyst's own
  command now fails — `pnpm exec tsx src/scripts/office/validate.ts
  really_broken/tool_redlines/crashes_word/jubarte__01a11a49_1862.docx` → `Error: … does not exist`.
- **Manifest row:** gone with the manifest (`really_broken/manifest.csv` does not exist). Per the
  analyst's earlier reading it said category `crashes_word`, `local_word_error` = `Microsoft Word
  got an error: Connection is invalid`, `d1_failed_in: soffice` — the AppleScript error text for a
  lost automation connection (Word quit or the event timed out mid-run). It says nothing about the
  document; the same string drove the false `crashes_word` label on
  `crashes_word__docxodus__01a11a49_0888.md` in this lane.
- **Prior assessment (proven):** `docs/really-broken/assessments.jsonl:22` (and
  `inspection-log.jsonl:22`) — status `proven`, defect *"None. Microsoft Word opens the original
  file with no warning, error, or offer to repair."*, change `none`, `word.original:
  "clean-open"`. **The established Word cause is that nothing in this file breaks Word.** The
  missing `w:gutter` is schema-invalid but harmless on open: Word defaults a missing gutter to 0,
  which is exactly why docx-validate's dedicated check reports it as an advisory warning, not an
  error (see below).

## The violation — `sectprchange-pgmar-missing-gutter`

**Offending XML** (analyst-quoted, `word/document.xml:22653`; no longer verifiable — the part is
gone with the specimen):

```xml
<w:sectPrChange w:author="Redline" w:date="1970-01-01T00:00:00Z" w:id="918">
  <w:sectPr>
    <w:pgSz w:h="16838" w:w="11906" w:orient="portrait"/>
    <w:pgMar w:bottom="1417" w:top="1417" w:left="1134" w:right="1134" w:header="708" w:footer="720"/>
  </w:sectPr>
</w:sectPrChange>
```

**Schema rule** — read directly at `src/scripts/office/schemas/ISO-IEC29500-4_2016/wml.xsd:1451-1459`
(`CT_PageMar`; `gutter` is line 1458):

```xml
<xsd:complexType name="CT_PageMar">
  <xsd:attribute name="top" type="ST_SignedTwipsMeasure" use="required"/>
  <xsd:attribute name="right" type="s:ST_TwipsMeasure" use="required"/>
  <xsd:attribute name="bottom" type="ST_SignedTwipsMeasure" use="required"/>
  <xsd:attribute name="left" type="s:ST_TwipsMeasure" use="required"/>
  <xsd:attribute name="header" type="s:ST_TwipsMeasure" use="required"/>
  <xsd:attribute name="footer" type="s:ST_TwipsMeasure" use="required"/>
  <xsd:attribute name="gutter" type="s:ST_TwipsMeasure" use="required"/>
</xsd:complexType>
```

**Proposed fix (generator):** when jubarte constructs the cached pre-change `w:sectPr` it records
inside `w:sectPrChange`, it must emit the complete seven-attribute `w:pgMar`, writing
`w:gutter="0"` when the source value is absent — the same default Word's own writers emit.
**Fix target: jubarte.**

**Why disputed:** the rule reproduces (quoted above) and the shape reproduces (below), but the
offending site in this document cannot be re-read; the snippet above is the analyst's quote, not a
file this session inspected.

**Correction to the analyst's check-catalog claim.** The analyst wrote that "the CLI has no
dedicated check for a missing pgMar gutter — the defect surfaces only through the generic xsd-error
code." That is wrong against this repo: `word-page-margin-attributes`
(`src/scripts/office/checks/page-margin-attributes.ts:42-76`, registered in
`src/scripts/office/checks/index.ts:89` inside `WORD_RULE_CHECKS`) is exactly that check. It walks
every `w:pgMar` in the story parts in both `WORD_PARAGRAPH_NAMESPACES`, names the
`w:sectPrChange` placement explicitly (lines 30-35), and reports each missing attribute as an
advisory **warning** — its comment (lines 14-18) records why: "Word silently defaults the missing
ones — header/footer to 720 twips, gutter to 0." The analyst's xsd-error-only output implies a
schema-walking profile (`--strict`/`--lenient`); under lenient the Word rules do not run, and the
default `word-valid` profile runs no XSD for DOCX — the bare command as quoted would have produced
neither the xsd pair nor (per its own docstring at `src/scripts/office/validate.ts:167`) anything
but the warning.

## Where the fixes belong and why

- **jubarte (the only fix target).** The generated redline XML is schema-invalid: a `w:pgMar` with
  six attributes where `CT_PageMar` requires seven. The sibling specimen
  `really_broken/unpacked/tool_redlines__repair_prompt__jubarte__01a11a49_0792/word/document.xml`
  shows jubarte's own `w:sectPrChange` (same `w:author="Redline"`, same epoch date) normally
  carrying `w:gutter="0"` — so the omission on 1862 is per-document, either an inherited source
  omission passed through verbatim or a serializer edge case; either way the generator that writes
  the cached copy owns completing it.
- **Not docx-validate.** No gap and no false positive: the default `word-valid` profile flags the
  exact shape with the dedicated advisory check (proven below), and `--strict` surfaces the
  `xsd-error`. The advisory severity is deliberate and matches the proven `clean-open`.
- **Not `repair`.** The source document is not the patient — the generated redline is — and Word
  opens the file clean, so there is nothing to repair for Word's sake.
- **Lane bookkeeping (not a code fix).** The `crashes_word` label is a false positive: the manifest
  row's `Connection is invalid` is an automation-transport failure and the proven assessment is
  `clean-open`. The specimen should be restored or the lane row retired; the sibling issue
  `really_broken/issues/crashes_word__docxodus__01a11a49_0888.md` records the same double failure.

## Verification

Run this session, on a scratch copy of the sibling jubarte specimen with `w:gutter` stripped from
exactly the `w:pgMar` inside its `w:sectPrChange` (the reported shape reconstructed, since 1862 is
gone):

1. `pnpm exec tsx src/scripts/office/validate.ts <stripped>.docx --json` — default `word-valid`
   profile: check `word-page-margin-attributes` **executed** and emitted exactly one warning on
   `word/document.xml`:
   *"The w:pgMar of the w:sectPr recorded by a w:sectPrChange is missing the required w:gutter
   attribute. CT_PageMar requires it; write w:gutter explicitly — Word's own writers emit
   w:header="720" w:footer="720" w:gutter="0" for the defaults."*
2. `pnpm exec tsx src/scripts/office/validate.ts <stripped>.docx --strict` —
   `ERROR [word/document.xml]: Element '…pgMar': The attribute '…gutter' is required but missing.`
   — the same finding the analyst reported, confirming their run used a schema-walking profile.
3. `xmllint --noout --schema src/scripts/office/schemas/ISO-IEC29500-4_2016/wml.xsd
   <stripped>/word/document.xml` — libxml2 confirms the violation independently (the sibling part
   carries unrelated pre-existing errors — `gridCol/@w:type`, `shd/@w:fill`, `bottom` — so the
   repo engine's per-finding view above is the cleaner reference).

**To close the dispute on 1862 itself**, restore the specimen (or its unpacked copy) and re-run
step 1 — the warning must name `word/document.xml` and the `w:sectPrChange` placement — and re-run
the two-run Word probe per `AGENTS.md` ("Always check twice, with logs"), expecting `clean-open`
twice, which the proven assessment already recorded. Absent the specimen, the violation stands as
analyst-quoted only.
