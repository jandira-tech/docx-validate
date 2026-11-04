# [superdoc-redlines 2380] repair-prompt: invalid w:ins/w:del attributes (GUID ids, w:authorEmail)

## Summary

`tool_redlines/repair_prompt/superdoc-redlines__01a11a49_2380.docx` (sha256
`736eff18f871a5fb5e886d59d9b3d5558c9bd3ecded8a9163d28adcd8f4c5fc3`) is a redline
produced by **superdoc-redlines**: 174 body-level tracked changes, all stamped
`w:author="Live Bench"` / `w:authorEmail="ai@example.com"` / `w:date="2026-10-13T20:00:00.000Z"`.
Unpacked at `really_broken/unpacked/tool_redlines__repair_prompt__superdoc-redlines__01a11a49_2380/`.

Live Word refuses the file with the recovery dialog. This repo's **strict** profile reports
**452 `xsd-error`** findings, which I re-verified independently and which reconcile exactly
with the six violations below; the default **word-valid** profile passes the package clean.
Which defect Word is actually refusing is **not yet proven** — no minimal-change experiment
exists — but all six defects are genuine schema violations, independently confirmed with
`xmllint` against the bundled `wml.xsd`.

## Word behaviour

- **Live result (this run):** `repair-prompt` — "Word found unreadable content in
  superdoc-redlines__01a11a49_2380.docx. Do you want to recover the contents…?" (Yes/No);
  No → "Word experienced an error trying to open the file", no document window. No crash
  report; Word quit cleanly.
- **Manifest:** no row in this repo (`really_broken/manifest.csv` does not exist here).
  The analyst notes record the mirror-corpus manifest row: category `repair_prompt`,
  `d1_failed_in word:corrupt`, `word_screen.log` capturing the same repair prompt twice.
- **Prior assessment:** none — `grep -c superdoc-redlines__01a11a49_2380` returns 0 in both
  `docs/really-broken/assessments.jsonl` and `docs/really-broken/inspection-log.jsonl`.
- **docx-validate runs (mine):** default `pnpm exec tsx src/scripts/office/validate.ts <file>`
  → `All validations PASSED!` (15 non-blocking: `paraid-missing-element` ×6,
  `style-default-missing` ×3, `xml-bom-leading` ×3, `rels-unreferenced` ×1,
  `rel-ids-sequential` ×2). `--strict --json` → `valid: false`, 452 `xsd-error`
  (document.xml 450, settings.xml 1, styles.xml 1) + 3 `xsd-summary`.

## Violation 1 — GUID `w:id` on every `w:ins`/`w:del` (174 sites)

`word/document.xml:90` (first `w:del`), `:108` (first `w:ins`), last at `:2760`. All 174
body-level tracked changes (94 `w:ins` + 80 `w:del`; no other part has any) stamp `w:id`
with a GUID — only 18 distinct GUIDs reused across 174 elements (77 share
`4a5f93fa-771a-4c38-9535-bd0192ec7fa3`).

```xml
<w:del w:id="eaa93de7-9923-473c-9825-62ecd815ca57" w:author="Live Bench"
       w:authorEmail="ai@example.com" w:date="2026-10-13T20:00:00.000Z"
       w:rsidR="00E81F93" w:rsidRPr="00C534E6">
```

Schema (`src/scripts/office/schemas/ISO-IEC29500-4_2016/wml.xsd`): `:843-845`
`<xsd:complexType name="CT_Markup"> <xsd:attribute name="id" type="ST_DecimalNumber" use="required"/>`; `:56-58` `<xsd:simpleType name="ST_DecimalNumber"> <xsd:restriction base="xsd:integer"/>`. Body-level `ins`/`del` are typed `CT_RunTrackChange` (`:3405-3406`, in
`EG_RunLevelElts`); `CT_RunTrackChange` (`:997-1006`) extends `CT_TrackChange`
(`:846-851`) extends `CT_Markup`. xmllint: `document.xml:90: element del: …
attribute '…}id': 'eaa93de7-…' is not a valid value of the atomic type
'…}ST_DecimalNumber'` (×174). Strict walk emits 174 `xsd-error`.

**Fix:** stamp `w:id` with integers (Word's own convention: small unique per-change
integers), not GUIDs. **Fix target: superdoc-redlines.**

## Violation 2 — undeclared `w:authorEmail` on every `w:ins`/`w:del` (174 sites)

Same 174 elements carry `w:authorEmail="ai@example.com"` in the plain `w:` namespace.
The `CT_Markup`/`CT_TrackChange`/`CT_RunTrackChange` closure declares only `id`, `author`,
`date`. `grep -c authorEmail` over `wml.xsd` and all nine bundled `microsoft/*.xsd`
returns 0 — the attribute exists in no bundled schema (Word's own spelling of that datum,
`w14:authorEmail`, is a comments-extension attribute, not a track-change one).
xmllint: `attribute '…}authorEmail': The attribute '…}authorEmail' is not allowed.` (×174).

**Fix:** drop `w:authorEmail` from track changes (no schema-legal slot exists for it).
**Fix target: superdoc-redlines.**

## Violation 3 — undeclared `w:rsidRPr` on `w:ins`/`w:del` (94 sites: 16 `w:ins` + 78 `w:del`)

`w:rsidRPr` is declared only on `CT_HdrFtr` (wml.xsd:1633), `CT_R` (:1787), `CT_P` (:2243)
and `CT_Row` (:2423) — never on `CT_TrackChange`/`CT_RunTrackChange`.

xmllint: `attribute '…}rsidRPr': … is not allowed.` (×94). Strict walk: 94 `xsd-error`.

**Fix:** drop `w:rsidRPr` from `w:ins`/`w:del`. **Fix target: superdoc-redlines.**

## Violation 4 — undeclared `w:rsidR` on `w:ins`/`w:del` (8 sites: 4 `w:ins` + 4 `w:del`)

`w:rsidR` is likewise declared only on `CT_HdrFtr` (wml.xsd:1635), `CT_R` (:1789),
`CT_P` (:2244) and `CT_Row` (:2424). xmllint: `document.xml:90` / `:108` etc. — 8 findings.

**Fix:** drop `w:rsidR` from `w:ins`/`w:del`. **Fix target: superdoc-redlines.**

## Violation 5 — `w:footnotePr` after `w:listSeparator` in settings.xml (1 site)

`word/settings.xml:72-75` places `w:footnotePr` after `w:decimalSymbol` (:70) and
`w:listSeparator` (:71). `CT_Settings`' sequence slots `footnotePr` far earlier — before
`compat` — and ends at `decimalSymbol`/`listSeparator`:

- wml.xsd:2972 `<xsd:element name="footnotePr" type="CT_FtnDocProps" minOccurs="0"/>`
- wml.xsd:2974 `<xsd:element name="compat" …/>`; :2988 `shapeDefaults`;
  :2990 `decimalSymbol`; :2991 `listSeparator` (final member).

xmllint: `settings.xml:72: element footnotePr: This element is not expected.`

**Fix:** move the `w:footnotePr` block to its schema slot (anywhere before `w:compat`/
`w:decimalSymbol`, e.g. where Word itself writes it, near the top with the other doc-level
properties). **Fix target: superdoc-redlines** (serializer child-order bug).

## Violation 6 — `w:tblCellMar` before `w:tblBorders` in styles.xml (1 site)

Table style `w:styleId="Tabellenraster"` (styles.xml:273): its `w:tblPr` (:281) orders
`tblInd` (:282) → `tblCellMar` (:283-288) → `tblBorders` (:289-296). `CT_TblPrBase`'s
sequence (verified enclosing type at wml.xsd:2486-2495) is `tblInd` :2491 →
`tblBorders` :2492 → `shd` :2493 → `tblLayout` :2494 → `tblCellMar` :2495.

xmllint: `styles.xml:289: element tblBorders: This element is not expected.`

**Fix:** emit `w:tblPr` children in schema order (`tblInd` → `tblBorders` → … →
`tblCellMar`). **Fix target: superdoc-redlines** (serializer child-order bug).

## Where the fixes belong and why

All six belong in **superdoc-redlines**. The XML it emits is genuinely invalid, and the
defect pattern is the tool's own stamp (one author, one email, GUID identifiers reused
across 174 changes) plus two serializer ordering slips. **docx-validate needs no change
for detection**: `--strict` reports every instance exactly — my run's 452 findings
classify as 174 GUID-id + 174 authorEmail + 94 rsidRPr + 8 rsidR + 1 footnotePr order +
1 tblPr order, matching `xmllint` on the same parts.

One open validator question, deliberately *not* a fix target yet: the default word-valid
profile passes this package clean (word-valid runs no XSD for DOCX by design) while Word
refuses it. Every blocking word-valid rule in this repo is pinned to a proven Word
behaviour; until the minimal-change experiment below proves which defect Word refuses,
there is nothing sound to pin a blocking rule on. If the experiment proves the track-change
attribute family, a named Word rule for it (like the existing tracked-change checks) is
the follow-up.

## Verification

Cause vs Word is **not proven** — state of evidence: strict/XSD violations proven (two
engines), Word refusal observed once live, no assessment row, no isolation experiment.

1. Reproduce detection: `pnpm exec tsx src/scripts/office/validate.ts
   really_broken/tool_redlines/repair_prompt/superdoc-redlines__01a11a49_2380.docx
   --strict --json` → `valid: false`, 452 `xsd-error`.
2. Independent engine: `xmllint --noout --schema
   src/scripts/office/schemas/ISO-IEC29500-4_2016/wml.xsd <unpacked>/word/document.xml`
   (and `settings.xml`, `styles.xml`) — the same six findings with stored-part line numbers.
3. Minimal-change experiment to prove Word's cause, each variant opened **twice** per this
   repo's two-run Word-probe standard, applied one family at a time on copies:
   a. renumber the 174 GUID `w:id` to unique integers **and** drop `w:authorEmail`,
      `w:rsidR`, `w:rsidRPr` (all 450 document.xml defects) → reopen;
   b. if still refused, instead move `settings.xml`'s `w:footnotePr` before
      `w:decimalSymbol` → reopen; then swap `styles.xml`'s `tblCellMar`/`tblBorders`.
   The variant that first opens clean names the refusing defect.
4. Generator regression: a fresh superdoc-redlines output validated clean by
   `xmllint --schema wml.xsd` (integer ids, only `id`/`author`/`date` attributes, schema-
   ordered `w:settings`/`w:tblPr` children) and opened clean in Word, twice.
