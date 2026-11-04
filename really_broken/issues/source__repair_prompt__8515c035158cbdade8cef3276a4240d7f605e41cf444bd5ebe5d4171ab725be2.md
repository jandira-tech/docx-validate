# [source 8515c035] failed: fractional a:arcTo radii make Word refuse the file

## Summary

- **Document:** `fixtures/repair_prompt/8515c035158cbdade8cef3276a4240d7f605e41cf444bd5ebe5d4171ab725be2.docx` — a source document (lane `repair_prompt`, producer `source-doc`) whose main document draws a custom-geometry rounded-rectangle shape (`a:custGeom` inside `wps:wsp`/`wps:spPr`, `w:drawing`).
- **The lane paths are stale.** `really_broken/fixtures/`, `really_broken/manifest.csv` and `really_broken/unpacked/source__repair_prompt__8515c035…` do not exist (repo-wide `find` returns nothing for `8515c035` outside `docs/really-broken/`), which is why this session's live Word probe reported `failed` without opening anything. The exact bytes survive in-repo as `tests/fixtures/broken/package.drawing-arc-radius.docx`, whose sha256 is `8515c035158cbdade8cef3276a4240d7f605e41cf444bd5ebe5d4171ab725be2` — identical to the specimen id (verified with `shasum -a 256`). Everything below was re-verified against those hash-identical bytes.
- **Outcome:** Word refuses the file (open error). docx-validate's default word-valid profile already blocks it (`word-drawing-arc-radius`, 8 errors, exit 1). The analyst report's `checkCode: ""` claim does not reproduce on the current tree (see "Where the fixes belong").

## Word behaviour

- **This session:** no live Word open — the file does not exist at the lane path (`outcome: "failed"`). Stated plainly: no new Word observation was made here.
- **Prior, hash-bound evidence for these exact bytes:**
  - `tests/fixtures-all.manifest.json` (entry at line 89): `broken/package.drawing-arc-radius.docx` → `word: { outcome: "open-error", sha256: 8515c035…725be2, observedAt: 2026-10-26T01:30:59Z, wordVersion: 16.115 }`.
  - Word session 41 `s41a` (`docs/plans/evidence/word-probes-2026-10-24.md:2517-2518`; two runs, both readings agree): the broken fixture (`8515c035158cbdad`) → **open-error / open-error**; the integer control `working/package.drawing-arc-radius-integer.docx` (`d6faf27e24b0d540`) → **clean-open / clean-open**.
- **Proven minimal-change assessment** (`docs/really-broken/assessments.jsonl:5`, `inspection-log.jsonl:5`, status `proven`): replacing each fractional radius token with the integer `408` — no other `document.xml` character changes — turned `open-error` into `clean-open`; restoring a single `wR` to `407.666666666667` re-broke the open. One non-integer radius is sufficient.

## Violation 1 — `a:arcTo` fractional radii (8 attribute values, `word/document.xml`)

**Offending XML** — lines 55–67 of `word/document.xml`, read from the unpacked hash-identical package (`arcTo` occurs in no other part):

```xml
<a:arcTo wR="407.666666666667" hR="407.666666666667" stAng="16200000" swAng="-5400000"/>  <!-- line 55 -->
<a:arcTo wR="407.666666666667" hR="407.666666666667" stAng="10800000" swAng="-5400000"/>  <!-- line 59 -->
<a:arcTo wR="407.666666666668" hR="407.666666666667" stAng="5400000" swAng="-5400000"/>   <!-- line 63 -->
<a:arcTo wR="407.666666666668" hR="407.666666666667" stAng="0" swAng="-5400000"/>          <!-- line 67 -->
```

**Schema rule** — bundled `src/scripts/office/schemas/ISO-IEC29500-4_2016/dml-main.xsd`, quoted lines read this session:

- `:2037-2043` `CT_Path2DArcTo`: `<xsd:attribute name="wR" type="ST_AdjCoordinate" use="required"/>` (same for `hR`).
- `:1968-1970` `ST_AdjCoordinate` = `<xsd:union memberTypes="ST_Coordinate ST_GeomGuideName"/>`.
- `:189-196` `ST_Coordinate` = `union(ST_CoordinateUnqualified, s:ST_UniversalMeasure)`, where `ST_CoordinateUnqualified` is `<xsd:restriction base="xsd:long">` (integers only) and `ST_UniversalMeasure` (`shared-commonSimpleTypes.xsd:144-148`) is the pattern `-?[0-9]+(\.[0-9]+)?(mm|cm|in|pt|pc|pi)` — a decimal **requires a unit suffix**.
- `:1953-1955` `ST_GeomGuideName` = `<xsd:restriction base="xsd:token"/>` — unrestricted.

So `407.666666666667` matches **no numeric branch** (`xsd:long` rejects the fraction; universal measure demands a suffix) and is XSD-legal only as a *guide name*. But the shape's `a:gdLst` (`document.xml` lines 42–47) defines only `textAreaLeft`, `textAreaRight`, `textAreaTop`, `textAreaBottom` — no guide holds these tokens. Word's runtime parses a numeric radius as a coordinate, requires an integer, and refuses the whole file (open error, s41a, both runs).

**Proposed fix (repair):** in `word/document.xml`, round all eight radius tokens to the integer `408` — `wR="408" hR="408"` on lines 55, 59, 63 and 67. `407.666…` (= 1223/3) rounds to 408, and the same path's adjacent points already use 408 (`<a:pt x="408" y="0"/>`, `<a:pt x="10696" y="408"/>`), so 408 is the value the source generator should have emitted. This is exactly the proven assessment's change and matches the clean control fixture byte-for-byte in intent.

**Fix target:** `repair` — broken source document (producer is `source-doc`, not a redline generator).

## Where the fixes belong and why

- **`repair`** — the document's XML is at fault, and only it. All evidence (s41a open-error ×2 beside an integer twin clean ×2, plus the proven both-directions minimal-change experiment) points at the eight fractional radii and nothing else.
- **Not `docx-validate`** — the analyst's central claim ("checkCode is '' … the CLI is blind to it") does not reproduce. What I ran on the hash-identical fixture:
  - `pnpm exec tsx src/scripts/office/validate.ts tests/fixtures/broken/package.drawing-arc-radius.docx` → exit 1, **8 ERROR `word-drawing-arc-radius` lines**, one per offending attribute (default profile `word-valid`).
  - Same with `--json` → `valid: false`; codes: `word-drawing-arc-radius: 8`, plus pre-existing `rel-ids-sequential: 2` (info), `paraid-missing-element: 1` and `style-default-missing: 1` (warnings).
  - Same with `--profile lenient` → exit 0, `All validations PASSED!` — the only way the analyst's output reproduces. That is by design, not a gap: named Word rules run only under full `word-valid`/`strict` (AGENTS.md note 18), and the XSD walk genuinely cannot flag this defect — `ST_GeomGuideName` accepts any token, and `removeNonOoxmlElements` strips the `wps:` subtree before the engine sees the arcs — which is precisely why the named check `word-drawing-arc-radius` exists (`src/scripts/office/checks/word-drawing-values.ts:99-135`, registered at `checks/index.ts:139`, Word-blocking at `validators/word-blocking.ts:278`).
- **No producing-tool target** — this lane's document is a broken source, so the repair prescription above is the whole fix.

## Verification

1. Apply the repair (all eight radii → `408`) and re-run the default profile: the 8 `word-drawing-arc-radius` errors disappear and the run exits 0 (the paraid/style/rel-ids findings are pre-existing info/warnings in this specimen and unaffected).
2. Minimal-change experiment (already proven, `assessments.jsonl:5`): bytes identical except the eight radius tokens → `clean-open`; restoring one fractional `wR` → `open-error`. Any fresh Word probe must follow the two-run protocol (AGENTS.md "Real-Word probes").
3. Regression coverage already exists in-repo: `tests/checks-word-drawing-values.test.ts`, the hash-bound manifest row for `broken/package.drawing-arc-radius.docx`, and the integer twin `working/package.drawing-arc-radius-integer.docx` pinning the clean control.
