# [superdoc-redlines 2283] repair-prompt: `[trash]` `.dat` parts have no content type

## Summary

`tool_redlines/repair_prompt/superdoc-redlines__01a11a49_2283.docx` (503,423 bytes, sha256
`65bb63702c01a7c7e403e72952873baaa94de02f9960ea0149e9acfce55f034e` — verified this session with
`shasum -a 256`, matching harvest manifest row 152) is a machine-authored tracked-changes redline:
every `w:del`/`w:ins` carries `w:author="Live Bench" w:authorEmail="ai@example.com"
w:date="2026-10-13T11:20:00.000Z"`. The producing tool, **superdoc-redlines**, stored two of its
own staging parts in the export — `[trash]/0000.dat` (243 bytes) and `[trash]/0001.dat`
(1,639 bytes) — and `[Content_Types].xml` covers neither: the stream's only `Default` entries are
`png`, `rels`, `xml`, and no `Override` names either part.

Live Word refuses the file with the recovery dialog; the repo's default `word-valid` run returns
`valid: false` with exactly these two `ct-uncovered-part` errors and no other error-severity
finding — validator and Word agree.

## Word behaviour

- **Live run (given with this ask):** Word showed *"Word found unreadable content in
  superdoc-redlines__01a11a49_2283.docx. Do you want to recover the contents of this document?"*
  (Yes/No); answering **No** was followed by *"Word experienced an error trying to open the file"*
  and no document window opened. Outcome: `repair-prompt`. No crash report; Word quit cleanly.
- **Manifest row:** `really_broken/manifest.csv` does not exist in this checkout (checked this
  session), so its wording could not be read. The harvest manifest at
  `~/temp/T/neurotic_docx_bench/maybe_future_corpus/really_broken/manifest.csv` line 152 records
  `word:corrupt` — *"Word could not load it in 2 passes: document loaded empty (Word could not
  read it)"* — which agrees with the live refusal.
- **Prior assessment:** none. Greps for this path in `docs/really-broken/assessments.jsonl` and
  `docs/really-broken/inspection-log.jsonl` return no row (exit 1), so **the causal chain is not
  yet proven by an experiment**; the evidence below is the OPC rule plus the live Word refusal.

## Violation 1 — `ct-uncovered-part`: `[trash]/0000.dat` and `[trash]/0001.dat` have no content type (2 sites)

**Offending site.** `[Content_Types].xml` (read in full this session) declares only:

```xml
<Default Extension="png" ContentType="image/png"/>
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
```

…followed by 27 `Override` entries (`word/document.xml` through `docProps/custom.xml`), none for
`/[trash]/0000.dat` or `/[trash]/0001.dat`, and no `Default` for extension `dat`.

Uncovered entries in the archive (verified with `zipinfo -1`, which lists 57 entries):
`[trash]/0000.dat` is entry 22 (243 bytes), `[trash]/0001.dat` is entry 25 (1,639 bytes), beside a
`[trash]/` directory entry at 54. Both blobs are mangled binary — a hexdump shows four `ef bf bd`
bytes (U+FFFD REPLACEMENT CHARACTER in UTF-8) then NUL padding to the end of each file — and
**no relationship mentions them** (grep of `_rels/.rels` and `word/_rels/*.rels` for `trash`
exits 1; the validator's `rels-unreferenced` warnings target exactly these two parts). Every ZIP
entry, the `[trash]` ones included, carries one timestamp (11:28), eight minutes after the
redline's `w:date` — one write by the producing tool, not a later edit.

**Rule.** OPC requires a content type for every part in the package, and `Default`/`Override`
entries are the only coverage mechanisms (ECMA-376 Part 2 content-types model; analyst cites
§9.1.4 / §10.1.2.2). The bundled schema proves the structural half of that rule —
`src/scripts/office/schemas/ecma/fourth-edition/opc-contentTypes.xsd:11-16`:

```xml
<xs:complexType name="CT_Types">
  <xs:choice minOccurs="0" maxOccurs="unbounded">
    <xs:element ref="Default"/>
    <xs:element ref="Override"/>
  </xs:choice>
</xs:complexType>
```

with `CT_Default` (`Extension`, `ContentType`, both required) at lines 18-21 and `CT_Override`
(`ContentType`, `PartName`, both required) at lines 23-26. No XSD line can express the *coverage*
invariant — I confirmed this directly:

```console
$ xmllint --noout --schema src/scripts/office/schemas/ecma/fourth-edition/opc-contentTypes.xsd \
    "really_broken/unpacked/tool_redlines__repair_prompt__superdoc-redlines__01a11a49_2283/[Content_Types].xml"
.../[Content_Types].xml validates        # exit 0
```

The stream is structurally valid; the defect is that two stored parts are matched by nothing in it.

**Reproduction (this session, default word-valid profile):**

```console
$ pnpm exec tsx src/scripts/office/validate.ts \
    really_broken/tool_redlines/repair_prompt/superdoc-redlines__01a11a49_2283.docx --json
valid: false
error   ct-uncovered-part  ×2  ([trash]/0000.dat, [trash]/0001.dat)
warning rels-unreferenced  ×2  (same two parts)
warning paraid-missing-element ×2;  warning xml-syntax ×2 (customXml/item2.xml, item4.xml:
                                  "xmlns: URI <guid> is not absolute" — SharePoint GUID
                                  namespaces, non-blocking);  info rel-ids-sequential ×3
```

The check is `ct-uncovered-part`, emitted at `src/scripts/office/validators/base.ts:1691`.

**Proposed fix — target: `superdoc-redlines`.** The generator must not emit package parts the
content-types stream does not cover. Concretely: stop storing `[trash]/*` in the exported
redline (nothing references them, they are not OOXML-defined parts, and their bytes are mangled
beyond use), or — if the trash payload must ship — add
`<Default Extension="dat" ContentType="application/octet-stream"/>` (or per-part `Override`s)
when writing `[Content_Types].xml`. Omission is the better fix; a covered `.dat` would still be
dead weight Word never reads.

## Where the fixes belong and why

- **`superdoc-redlines` (both sites):** the `[trash]` parts are the tool's own artifacts, written
  in the same ZIP pass as the rest of the package, so the defect originates in the generator's
  export path, not in a broken source document and not in the validator.
- **Not `docx-validate`:** the validator already reports this defect exactly — two
  `ct-uncovered-part` errors, no false positives, no missed error-severity finding in the run —
  so there is no gap to fill and nothing to retract.
- **Not `repair`:** this is not a broken source document; it is the redline the tool ships.

## Verification

- Re-run the named check (done this session, command above): expect `valid: false` with exactly
  two `ct-uncovered-part` errors on the `[trash]` parts.
- **Minimal-change experiment to prove the Word cause (not yet run — no proven assessment exists
  for this path):** rebuild the archive without `[trash]/0000.dat`, `[trash]/0001.dat` and the
  `[trash]/` directory entry, and open it in Word twice with logs per the repo's Word-probe
  protocol. A clean open in both runs proves the uncovered parts are the refusal cause.
- Differential arm: keep the entries and add
  `<Default Extension="dat" ContentType="application/octet-stream"/>`; a clean open would show
  coverage (not the parts' bytes) is what Word objects to.
- If the repair prompt survives removal, the next suspects are the `word/document.xml`
  tracked-change spellings the lenient profile's XSD walk flags (GUID-valued `w:id` on every
  `w:del`, `w:authorEmail`, `w:rsidRPr` — none declared by `CT_RunTrackChange`); the word-valid
  profile runs no XSD for DOCX, so those are unprobed against Word and are recorded here as
  candidates only, not violations.
- A regression fixture for the generator lane: the specimen itself (hash above) already serves as
  the `repair_prompt` calibration for this shape.
