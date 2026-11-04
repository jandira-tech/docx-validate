# [superdoc-redlines 0767] repair-prompt: hyperlink r:id values name no Relationship

## Summary

`tool_redlines/repair_prompt/superdoc-redlines__01a11a49_0767.docx` (sha256
`16adaa517fe48669b397a1cb74d09bfd597ab55985e621a25d1c95eb00593b5f`, 199,482 bytes) is a
tracked-changes redline produced by **superdoc-redlines**: a legal-citation document whose
`w:ins`/`w:del` runs are stamped with GUID `w:id`s, `w:author="Live Bench"` and
`w:authorEmail`, and whose citation links are wrapped in `w:hyperlink` elements.

Live Microsoft Word (this run) offers recovery: *"Word found unreadable content in
superdoc-redlines__01a11a49_0767.docx. Do you want to recover the contents of this
document?"* — answering **No** is followed by *"Word experienced an error trying to open the
file"* and no document window. Outcome `repair-prompt`. No `really_broken/manifest.csv` row
exists yet (the file is not on disk as of this session), so there is no manifest outcome to
compare against; the live result stands alone and agrees with the proven assessment below.

The dominant defect: **eight `w:hyperlink` elements in `word/document.xml` carry `r:id`
values that name no Relationship** — the part's relationships file defines only `rId1`–`rId14`
and contains no hyperlink-type relationship at all.

## Word behaviour

- **Live run (this session):** recovery dialog → No → open error → no document. Outcome
  `repair-prompt`.
- **Prior assessment** — `docs/really-broken/assessments.jsonl:117` (mirrored at
  `docs/really-broken/inspection-log.jsonl:117`), status **proven**: removing only those
  eight `r:id` attributes from the `w:hyperlink` start tags (every other byte of every part
  unchanged) took Word from `unreadable-content-warning` to `clean-open`. That is the
  established Word cause for this file's refusal.
- The docx-validate check naming this shape is `word-hyperlink-relationship`
  (`src/scripts/office/checks/word-hyperlink.ts`), whose header cites Word session 28
  (`docs/plans/evidence/word-probes-2026-10-24.md`, `s28a`, #614) for exactly this refusal.

## Violation 1 — `dangling-hyperlink-rid` (8 sites, `word/document.xml`)

**Sites.** `word/document.xml` lines **276, 299, 322, 345, 368, 391, 414, 437**, with
`rId672fcb25`, `rId27ad765f`, `rId0e48fc9d`, `rId5a6cfee7`, `rId79be095b`, `rId0c69145f`,
`rId47bfad64`, `rId13dc01b0` respectively (verified by grep of the unpacked part this
session).

**Offending XML** (line 276; the other seven repeat the identical shape):

```xml
<w:hyperlink w:history="1" r:id="rId672fcb25">
  <w:del w:id="4615b1d9-2023-4eb5-a99a-8da819690f81" w:author="Live Bench"
         w:authorEmail="ai@example.com" w:date="2026-10-10T02:00:00.000Z">
    <w:r><w:rPr><w:rStyle w:val="Hyperlink"/></w:rPr><w:delText>19-139
    </w:delText></w:r>
  </w:del>
</w:hyperlink>
```

**The relationships part** (`word/_rels/document.xml.rels`) defines exactly `rId1`–`rId14`:
styles, stylesWithEffects, settings, webSettings, footnotes, endnotes, headers 1–3, footers
1–3, fontTable, theme. **No hyperlink Type, and none of the eight referenced Ids.** Every
`r:id` in the document resolves except these eight (the others are `rId7`–`rId13`, all
present).

**Schema rule.** `CT_Hyperlink` in `src/scripts/office/schemas/ISO-IEC29500-4_2016/wml.xsd:1296–1304`:

```xsd
<xsd:complexType name="CT_Hyperlink">
  <xsd:group ref="EG_PContent" minOccurs="0" maxOccurs="unbounded"/>
  <xsd:attribute name="tgtFrame" type="s:ST_String" use="optional"/>
  <xsd:attribute name="tooltip" type="s:ST_String" use="optional"/>
  <xsd:attribute name="docLocation" type="s:ST_String" use="optional"/>
  <xsd:attribute name="history" type="s:ST_OnOff" use="optional"/>
  <xsd:attribute name="anchor" type="s:ST_String" use="optional"/>
  <xsd:attribute ref="r:id"/>
</xsd:complexType>
```

and `src/scripts/office/schemas/ISO-IEC29500-4_2016/shared-relationshipReference.xsd:8–11`:

```xsd
<xsd:simpleType name="ST_RelationshipId">
  <xsd:restriction base="xsd:string"/>
</xsd:simpleType>
<xsd:attribute name="id" type="ST_RelationshipId"/>
```

Because `ST_RelationshipId` is a bare `xsd:string`, **XSD validation cannot see this defect**
— the constraint that the id name a Relationship of the part's relationships part
(Transitional or ISO Strict hyperlink Type, `TargetMode="External"`) is a package-level
semantic requirement, enforced by Word and carried in this repo by the
`word-hyperlink-relationship` check. Confirmed this session: `xmllint --noout --schema
wml.xsd word/document.xml` reports 814 validity errors — 407× GUID-valued `w:id` not
`ST_DecimalNumber` and 407× `w:authorEmail` not allowed, all on the tool's own `w:del`/`w:ins`
stamps — and **zero** findings mentioning hyperlinks.

**Proposed fix.** In superdoc-redlines' output, each of these hyperlinks must either
(a) have its `<Relationship Id="rId…" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"
TargetMode="External" Target="…"/>` serialized into `word/_rels/document.xml.rels`, or
(b) drop the `r:id` attribute when no relationship is intended. The proven minimal-change
repair for the existing file is (b) restricted to the eight attributes — Word then opens
the document clean.

**Fix target:** `superdoc-redlines` (generator bug). The generated `rId` values look
minted (hex-suffixed, unlike the template's `rId1`–`rId14`), so the likeliest generator
defect is that hyperlink relationships are allocated ids but never written to the rels
part.

## Where the fixes belong and why

- **superdoc-redlines** — the redline's package is genuinely invalid: eight `r:id`
  references point at relationships that were never serialized. This is the only fix
  target; the file is a tool output, not a hand-broken source document, so no `repair`
  prescription is filed.
- **docx-validate** — no change. The `word-hyperlink-relationship` check exists
  (`src/scripts/office/checks/word-hyperlink.ts:83`), is registered under the `relationships`
  group, and fires correctly on this file; it is not a false positive. Verified this
  session (command and result under Verification).
- Documented, not filed: the 814 XSD deviations on `w:del`/`w:ins` (GUID `w:id` outside
  `ST_DecimalNumber`; `w:authorEmail`, which `CT_RunTrackChange` does not declare) are
  spec-invalid but **Word-tolerated** — the proven experiment opens clean with all 814 still
  present — and docx-validate's word-valid profile reports none of them. superdoc-redlines
  should conform its stamps anyway, but they are not the cause of the refusal.

## Verification

- **Validator run (this session):** `pnpm exec tsx src/scripts/office/validate.ts
  really_broken/tool_redlines/repair_prompt/superdoc-redlines__01a11a49_0767.docx` —
  exactly **8 ERROR lines**, all `ERROR [word/document.xml]` with code
  `word-hyperlink-relationship`, one per dangling id; **no warnings and no other codes**.
- **XSD check (this session):** `xmllint --noout --schema
  src/scripts/office/schemas/ISO-IEC29500-4_2016/wml.xsd <unpacked>/word/document.xml` —
  814 validity errors (407 `ST_DecimalNumber` + 407 `authorEmail`), 0 mentioning hyperlinks.
- **Proven minimal-change experiment** (assessment `docs/really-broken/assessments.jsonl:117`):
  delete only the eight `r:id` attributes from the `w:hyperlink` start tags, every other
  byte unchanged → Word goes from `unreadable-content-warning` to `clean-open`. Repro:
  unpack, strip the attributes, pack, open in Word twice with logs per the repo's
  two-run rule.
- **Regression home for the generator fix:** a fixture built from this file with the eight
  relationships added (Type `…/relationships/hyperlink`, `TargetMode="External"`) and one
  with the ids stripped should both open clean; docx-validate's existing
  `word-hyperlink-relationship` spec already pins both accepted shapes (session 28).
