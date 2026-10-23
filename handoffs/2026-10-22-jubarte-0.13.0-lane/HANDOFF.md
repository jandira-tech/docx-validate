# Handoff: docx-validate on the jubarte 0.13.0 redline lane (2026-10-22)

Before the jubarte-redlines 0.13.0 release, we ran docx-validate over the 600
redlines of the release sample. Each redline was checked against its two source
documents to see which error kinds the redline introduced. One docx-validate
check gives a false positive and needs fixing here. The other findings are
jubarte schema hygiene, and Word opens those files.

## What was run

- `lane_validity_dv.mjs` (copied here): it runs `validate()` from `dist/index.mjs`
  on each redline and on both of its sources. It keeps errors only, and skips
  the `xsd-summary` lines. A redline's "own" kinds are those that neither of
  its sources has.
- The redlines are `neurotic_docx_bench/results/redlines_0929_full/jubarte-0.13.0/docx/`.
  The pairs are `results/release_0.13.0_evidence/gen_pairs_sample.csv`, and the
  sources are `corpus/word/` (public on HF: `arthrod/neurotic_docx_bench-fixtures`).
- Full output: `lane_validity_docx_validate.log`. For comparison,
  `lane_validity_openxmlvalidator.log` is the same lane through the Open XML
  SDK's OpenXmlValidator (Office 2019).

| validator | redlines with errors | with a kind neither source has |
|---|---|---|
| docx-validate | 221 | 144 (40 without `ct-undeclared-ext`) |
| OpenXmlValidator (Office 2019) | 84 | 32 |

Microsoft Word opened and exported to PDF 39 of the 40 redlines that have
real kinds of their own. The 40th pairs against `1424386e9b…`. It is one of
the 23 redlines Word 16.115 crashes on (jubarte KNOWN_ISSUES #8), which were
already scored as zero.

## Fix here first: `ct-undeclared-ext` ignores `<Override>`

- **Scale:** this check flags 117 redlines and accounts for all 104 redlines
  whose only "own" kind is this one.
- **What the check reports:** "File with extension 'jpeg' not declared in
  [Content_Types].xml - should add: <Default Extension="jpeg" …/>".
- **What the package actually has:** a declaration through an override:

```xml
<Override PartName="/word/media/P6016719bb4db4c4dcb99b8d36b55d4c1905f2cad5b4cbccd7f6776808ace30f4.jpeg"
          ContentType="image/jpeg"/>
```

- **Why that is valid:** OPC (ECMA-376 Part 2, §10.1.2.4) resolves a part's
  content type from an `Override` with its exact PartName first, and only
  then from a `Default` on its extension. A part covered by an `Override`
  is declared.
- **Required behaviour:** `ct-undeclared-ext` should fire only when a part
  has neither an `Override` matching its PartName nor a `Default` for its
  extension. The PartName comparison is case-insensitive ASCII (OPC part
  name equivalence).
- **Where to look:** search `src/scripts/office/validators/` for
  `ct-undeclared-ext`.
- **Red test:** `fixtures/k01_1.redline.docx` must report no
  `ct-undeclared-ext`. Also keep a broken probe whose part has neither an
  Override nor a Default, and check that it still fires.

## The other kinds: jubarte output, Word opens it

These are the kinds the redline has and neither source has, in the
`kNN_*` fixtures. Counts are numbers of redlines. Each needs a decision
in jubarte, and possibly a docx-validate check against the transitional
schema.

| id | redlines | part | message (start) | what is in the redline |
|---|---|---|---|---|
| k02 | 15 | document.xml | `w:ins` not expected | `w:ins` under `w:p`, `w:rPr`, `w:trPr`, `m:ctrlPr` and `m:r`. **Check docx-validate:** ECMA `CT_CtrlPr` allows `w:ins`/`w:del`, so a hit on `m:ctrlPr` would be a second false positive. The math-run case (`m:r`) is what Word itself writes for a tracked change in an equation (seen at 0.11.2). |
| k03 | 13 | document.xml | `wp:extent` not expected, expected `wp:positionH` | Child order of a floating drawing (`wp:anchor`): the XSD order is simplePos, positionH, positionV, extent. This is jubarte's bug. |
| k04 | 10 | document.xml | `w:del` not expected | Same as k02 (`w:rPr`, `w:p`, `m:r`). |
| k05 | 8 | charts/chart1.xml | `c:ext` missing child | An empty `c:extLst/c:ext`. Jubarte should drop it or keep its child. |
| k06 | 7 | document.xml | `w:shd w:fill=""` | `<w:shd w:val="clear" w:fill=""/>`. An empty fill is not ST_HexColor. This is jubarte's bug: it should omit the attribute. |
| k07 | 6 | document.xml | `w:gridCol w:type` | `<w:gridCol w:w="1995" w:type="dxa"/>`. `CT_TblGridCol` has no `w:type`. This is jubarte's bug. |
| k08, k09 | 5 + 5 | document.xml, styles.xml | `w:bottom` not expected | Order of the children of `w:pBdr`/`w:tblBorders`: the XSD order is top, left/start, bottom, right/end, insideH, insideV/between, bar. This is jubarte's bug. |
| k10 | 2–4 | document.xml | `w:pgMar` requires `gutter` / `header` / `footer` | `w:pgMar` is missing required attributes. This is jubarte's bug: it should write all seven. |
| k11, k13 | 3 + 1 | header2.xml / header4.xml | `w:tblGrid` not expected | A header table written as the source's header under another part name (seen at 0.11.2). |
| k12 | 1 | document.xml | `wp:txbx` missing `w:txbxContent` | An empty text box. |

The kind ids follow the order of `manifest.csv`, which groups by kind. Check
`kind` there rather than relying on the numbering above.

## Fixtures

- **Where they are:** `fixtures/`, two examples per kind (`kNN_1`, `kNN_2`).
  Each example has three files: `.redline.docx` (jubarte 0.13.0),
  `.base.docx` and `.next.docx` (the two sources).
- **`manifest.csv`:** each row gives the kind, the file's sha256, its bench
  path, and, for redlines, whether Word opened and exported it.
- **Unpacking:** use `pnpm exec tsx src/scripts/office/unpack.ts fixtures/<file>`.
- **Where they come from:** all of them come from the public corpus, so no
  private documents are involved. They are not under `tests/fixtures/`
  because the `fixtures-all` suites pin that manifest. To make a test, move
  the one you need there and regenerate the manifest
  (`scripts/update-manifest.ts`).

## Tooling notes

- docx-validate issues carry no line number, so finding which `w:ins` (k02)
  fired means unpacking the file and running xmllint on the part.
- A full-lane run takes several minutes, and a large redline (k02) takes
  more than a minute on its own.
