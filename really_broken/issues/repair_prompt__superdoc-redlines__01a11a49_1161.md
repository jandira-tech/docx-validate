# [superdoc-redlines 1161] repair-prompt: two empty `w:drawing` elements

## Summary

`tool_redlines/repair_prompt/superdoc-redlines__01a11a49_1161.docx` (99,681 bytes,
sha256 `9e1e7727653d461ca8815dd44d12d1462a60217c8051f5635444ce56238ca459`) is a
tracked-changes redline produced by **superdoc-redlines** (repair_prompt lane, jubarte
artifact 1161): `w:del`/`w:ins` markup over a Ukrainian-language document. Microsoft
Word refuses to open it — the live probe saw the open-error dialog, no recovery offer,
no crash — and this repo's validator reports exactly the defect that causes it:
**two `word-empty-drawing` errors on `word/document.xml`** (I re-ran the validator
this session; exit 1, invalid). The manifest row (`really_broken/manifest.csv` in the
source corpus, line 110) records `word:convert_failed` — "document loaded empty
(Word could not read it)". Live Word, the corpus manifest, and the proven in-repo
assessment all agree: Word cannot open this file.

## Word behaviour

- **Live result (this session's probe):** refused without a recovery offer. The only
  dialog was `Word experienced an error trying to open the file. Try these
  suggestions. … (superdoc-redlines__01a11a49_1161.docx)` (OK only); after OK no
  windows or documents remained. No crash report; Word quit cleanly. The live record
  carries `agrees: true`. (Its `outcome` field is spelled `repair-prompt` while the
  dialog text is the open-error wording; the substance either way is refusal.)
- **Corpus manifest row 110:** `tool_redline, tool_redlines/repair_prompt/
  superdoc-redlines__01a11a49_1161.docx, repair_prompt, superdoc-redlines,
  9e1e7727…8ca459, …, word:convert_failed, "Word could not load it in 2 passes:
  document loaded empty (Word could not read it)"`.
- **Prior assessment: PROVEN** (`docs/really-broken/assessments.jsonl` line 121 and
  `docs/really-broken/inspection-log.jsonl` line 121, source `word-ab`). A minimal-
  change Word A/B experiment established the cause:
  - original file → `open-error`;
  - `1161-drop-both.docx` — **both empty `w:drawing` elements deleted, both
    `mc:AlternateContent` wrappers kept** → `clean-open`, and the validator is
    `valid` with no error codes;
  - `1161-leave-first.docx` / `1161-leave-second.docx` — either empty drawing left
    in place → still `open-error`, validator still reports `word-empty-drawing`.
  Conclusion: **each empty `w:drawing` independently makes Word refuse the file**;
  the AlternateContent wrappers, the unreferenced `word/numbering.xml`, and
  `docProps/custom.xml` are not blockers.

## Violation: `empty-wdrawing-no-anchor-or-inline` (2 sites, `word-empty-drawing`)

**Offending XML** — `word/document.xml` lines 905–911 and 966–972 (the `<w:drawing/>`
at lines 908 and 969), verified by reading the unpacked part
(`really_broken/unpacked/tool_redlines__repair_prompt__superdoc-redlines__01a11a49_1161`):

```xml
<w:r>
  <mc:AlternateContent>
    <mc:Choice Requires="wps">
      <w:drawing/>
    </mc:Choice>
  </mc:AlternateContent>
</w:r>
```

These are the only two drawings and the only two `mc:Choice` elements in the part;
there is no `mc:Fallback`. Each `w:drawing` is empty and is the sole child of its
`mc:Choice`, so no `wp:anchor` or `wp:inline` satisfies the content model anywhere.

**Schema rule** — `src/scripts/office/schemas/ISO-IEC29500-4_2016/wml.xsd` lines
1227–1232:

```xsd
<xsd:complexType name="CT_Drawing">
  <xsd:choice minOccurs="1" maxOccurs="unbounded">
    <xsd:element ref="wp:anchor" minOccurs="0"/>
    <xsd:element ref="wp:inline" minOccurs="0"/>
  </xsd:choice>
</xsd:complexType>
```

Caveat, verified this session with `/usr/bin/xmllint` (libxml 20913): raw wml.xsd
**validates** a minimal specimen holding a bare `<w:drawing/>` in a `w:r`
(`tmp/…/bare-drawing.xml validates`, exit 0) — libxml2 reads the inner
`minOccurs="0"` on both branches as permitting an empty choice — and rejects the
mc-wrapped specimen only because raw wml.xsd has no mc elements in its content
model (this repo validates after markup-compatibility preprocessing). So the XSD
citation names the ECMA-376 intent; the operative mechanical detector here is the
repo check `word-empty-drawing` (`src/scripts/office/checks/word-drawing-anchors.ts:33-83`,
which skips only drawings whose parent is a `w:p`, line 65 — these have parent
`mc:Choice`), backed by the proven Word refusal above.

**Proposed fix (generator):** superdoc-redlines must not emit a `w:drawing` — or the
`mc:AlternateContent`/`mc:Choice Requires="wps"` wrapper around one — for a shape it
cannot populate. Either emit a complete drawing (a `wp:anchor` or `wp:inline` child,
as CT_Drawing requires) or drop the drawing run entirely. For this file the proven
repair is: delete both empty `w:drawing` elements (the wrappers may stay) — that
opens clean.

**Fix target:** `superdoc-redlines` — the redline's XML is genuinely invalid and the
generator must change.

## Where the fixes belong and why

- **docx-validate — nothing to fix.** The check exists and catches both sites with
  no false positive: each flagged `w:drawing` has no element child and a non-`w:p`
  parent, exactly the check's documented contract. I re-ran
  `pnpm exec tsx src/scripts/office/validate.ts really_broken/tool_redlines/repair_prompt/superdoc-redlines__01a11a49_1161.docx`
  this session: exactly 4 output lines — two `ERROR [word/document.xml]: An empty
  <w:drawing> has no child element…` with their explanation lines — exit code 1.
- **superdoc-redlines — the fix belongs here.** The generator emitted two
  content-free drawing placeholders that violate CT_Drawing and that Word refuses
  (proven: removing them opens clean, leaving either one still refuses).
- **Not repair-lane for a source document.** The patient is the redline output, and
  the defect was introduced by the redline tool, not present in a source document.
  Non-blocking oddities the assessment deliberately set aside (GUID-valued `w:id` on
  `w:del`/`w:ins`, undeclared `w:authorEmail`/`w:rsidR`/`w:rsidRPr` attributes)
  stayed present in the clean-opening experiment and are not part of this issue.

## Verification

- **Established Word cause (already run, recorded in the assessment above):**
  deleting only the two empty `w:drawing` elements opens clean and validates;
  leaving either one still produces `open-error` + `word-empty-drawing`. That is
  the minimal-change experiment, done as an A/B with two runs per reading (`word-ab`).
- **Reproduction commands run this session** (paths from the repo root):
  - `pnpm exec tsx src/scripts/office/validate.ts really_broken/tool_redlines/repair_prompt/superdoc-redlines__01a11a49_1161.docx`
    → 2 × `word-empty-drawing` ERROR, exit 1.
  - `grep -n 'w:drawing' …/word/document.xml` → only lines 908 and 969, both empty.
  - `xmllint --noout --schema src/scripts/office/schemas/ISO-IEC29500-4_2016/wml.xsd <specimen>`
    on the two minimal specimens above → bare empty drawing validates (libxml2
    quirk), mc-wrapped one fails only on the mc elements.
- **For the generator fix:** a superdoc-redlines regression test asserting no
  emitted `w:drawing` lacks a `wp:anchor`/`wp:inline` child. This file already
  serves docx-validate as a real-world `word-empty-drawing` detector specimen
  (validator catches it; live Word refusal agrees).
