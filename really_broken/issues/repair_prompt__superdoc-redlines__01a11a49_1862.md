# [superdoc-redlines 1935] repair-prompt: docProps/custom.xml stored with no content type

## Summary

`tool_redlines/repair_prompt/superdoc-redlines__01a11a49_1935.docx` (sha256
`26e68a0198005576857fde48558602274b84788ead2b76c25f0a64f91446543f`) is a redline
produced by **superdoc-redlines**. The package stores `docProps/custom.xml` — the
generator's own metadata stamp (`SuperdocVersion=1.10.0`, `DocumentGuid=b852578e-…`) —
but `[Content_Types].xml` names that part by no `<Override>` and declares no
`<Default Extension="xml">`, so the part carries no content type at all.

- **Live Word outcome (this run):** `repair-prompt` — the recovery dialog
  "Word found unreadable content in superdoc-redlines__01a11a49_1935.docx"; answering
  No was followed by "Word experienced an error trying to open the file" and no
  document window.
- **Manifest:** `word:corrupt` — "Word could not load it in 2 passes: document loaded
  empty (Word could not read it)" (previous-generation corpus
  `maybe_future_corpus/really_broken/manifest.csv` row 116; this repo's
  `really_broken/manifest.csv` has no 1935 row and the packed file was absent from
  `really_broken/tool_redlines/repair_prompt/` — a concurrent process was populating
  it. I restored the file from the corpus copy; the sha256 above matches both.)
- **docx-validate:** `valid: false` with exactly **one error**, `ct-uncovered-part`
  @ `docProps/custom.xml` — agreeing with the live Word refusal.

## Word behaviour

- **Live result (this ask):** repair-prompt outcome; the observed dialog is quoted in
  the Summary above. No crash report; Word quit cleanly.
- **Prior assessment — proven** (`docs/really-broken/assessments.jsonl` line 126,
  mirrored in `docs/really-broken/inspection-log.jsonl` line 126):
  - original bytes → Word `unreadable-content-warning`, validator `valid: false`
    with codes `[ct-uncovered-part]`;
  - bytes + `<Override PartName="/docProps/custom.xml"
    ContentType="application/vnd.openxmlformats-officedocument.custom-properties+xml"/>`
    added → Word `clean-open`, validator `valid: true`, and **the custom properties
    themselves survived** (the existing Overrides for the relationship parts and for
    `image1.png` stayed too).
  
  The live `repair-prompt` and the recorded `unreadable-content-warning` are the same
  Word behaviour (recovery offer) read at different points; they agree.

## Violation 1 — `ct-uncovered-part`: `docProps/custom.xml` matches no Override and no Default

**Offending XML** (read from the unpacked package,
`really_broken/unpacked/repair_prompt__superdoc-redlines__01a11a49_1862`):

```xml
<!-- [Content_Types].xml — the ONLY Default (line 3); the 13 Overrides (lines 4–16)
     end at /docProps/app.xml and /docProps/core.xml. None names /docProps/custom.xml,
     and there is no Default for the xml extension: -->
<Default Extension="png" ContentType="image/png"/>
<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>

<!-- stored part docProps/custom.xml (523 bytes in the zip; well-formed per xmllint): -->
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties"
            xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
  <property name="SuperdocVersion" fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="2">
    <vt:lpwstr>1.10.0</vt:lpwstr>
  </property>
  <property name="DocumentGuid" fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="3">
    <vt:lpwstr>b852578e-66a5-485e-8a85-56a5c0d72ea8</vt:lpwstr>
  </property>
</Properties>
```

I cross-checked all 15 stored parts against the stream: `docProps/custom.xml` is the
**only** uncovered one (every other part — the `.rels` parts, `word/*`, both
`docProps/{app,core}.xml`, `image1.png` — has an Override; the png also has the
Default). One site, as reported.

**Rule** — ECMA-376 Part 2 (5th ed., Dec 2021) §7.2.3.2.1, "Media Types stream
markup — General":

> "For all parts of the package other than Relationships parts (6.5.2), the Media
> Types stream shall specify either: One matching Default element, or One matching
> Override element, or Both a matching Default element and a matching Override
> element, in which case, the Override element takes precedence."

(The `.rels` exception is moot here — this package Overrides its `.rels` parts
anyway; `docProps/custom.xml` is not a relationships part, so the rule binds.)

This rule is **prose no bundled XSD expresses**: `opc-contentTypes.xsd`'s `CT_Types`
(`src/scripts/office/schemas/ecma/fourth-edition/opc-contentTypes.xsd:11-16`) is only
an unbounded `choice` of `Default`/`Override` elements with pattern-constrained
attributes — nothing relates stream entries to stored part names. Confirmed
empirically: `xmllint --noout --schema …/opc-contentTypes.xsd [Content_Types].xml`
reports the stream **validates**, so pure schema validation is blind to this defect;
it takes the dedicated coverage pass in `src/scripts/office/validators/base.ts:1673-1693`
(`resolveContentType(...) !== null` per stored part) to catch it — and it does.

**Proposed fix** — add one Override to `[Content_Types].xml` (the proven minimal
change; a `Default Extension="xml"` would also cover it but is the blunter
instrument and is not what the assessment tested):

```xml
<Override PartName="/docProps/custom.xml"
          ContentType="application/vnd.openxmlformats-officedocument.custom-properties+xml"/>
```

**Fix target:** `superdoc-redlines` — the generator writes `docProps/custom.xml`
(its own version/GUID stamp) into the package but never declares that part's content
type when it builds the Media Types stream. Every part the packager emits must be
registered in the same pass.

## Where the fixes belong and why

- **superdoc-redlines (generator bug — the fix belongs here).** The redline's
  package is genuinely invalid against ECMA-376 Part 2 §7.2.3.2.1: the packager
  stores a part it never registers. When superdoc-redlines stamps
  `docProps/custom.xml`, it must also emit the custom-properties Override shown
  above. The proven assessment shows this single line is the whole difference
  between a repair prompt and a clean open, with the stamp's data retained.
- **Not docx-validate.** The validator already reports exactly this defect
  (`ct-uncovered-part`, error severity, `base.ts:1691`), with no false positives on
  this file — the run's only other findings are warnings (`rels-unreferenced` ×1,
  `paraid-missing-element` ×5, `style-default-missing` ×3, `word-shading-color` ×30,
  `word-cell-margin-order` ×60), none of them the refusal cause. No validator gap.
- **Not "repair".** This is a generator-produced redline artifact, not a broken
  source document handed to the tool.

## Verification

Re-verified in this session, from the restored bytes (sha256 matches the corpus
manifest row and the proven assessment's file):

1. `pnpm exec tsx src/scripts/office/unpack.ts …1935.docx really_broken/unpacked/repair_prompt__superdoc-redlines__01a11a49_1862 --merge-runs false --simplify-redlines false` — 13 XML files, OK.
2. `pnpm exec tsx src/scripts/office/validate.ts …1935.docx --json` — `valid: false`,
   exactly 1 error: `ct-uncovered-part` @ `docProps/custom.xml` (message: "Part
   'docProps/custom.xml' has no content type: matched by no <Override PartName> and
   no <Default Extension> in [Content_Types].xml…"); warnings as listed above.
3. `xmllint --noout --schema src/scripts/office/schemas/ecma/fourth-edition/opc-contentTypes.xsd` on the stream — **validates** (defect invisible to the XSD);
   `xmllint --noout` on `docProps/custom.xml` — well-formed.
4. `unzip -l` cross-check — `docProps/custom.xml` is the only stored part matched by
   neither an Override nor the png Default.
5. Established Word cause (proven, carried from the assessment): **adding the
   custom-properties Override for `/docProps/custom.xml` turns the repair prompt
   into a clean open** with the custom properties retained; the generator-side test
   is to repack with that one Override added and confirm `clean-open` plus
   `validate()` `valid: true` — exactly the assessment's `1935-custom-type.docx`
   experiment. A regression fixture for this lane would pin the restored file's
   sha256 to the recorded `word:corrupt` observation and the fixed variant to
   `clean-open`.
