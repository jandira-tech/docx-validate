# Word-first validation: TDD implementation plan

**Goal:** Make docx-validate the single dependency for portable Word-valid checks, common quick checks, built-in XML validation, strict schema validation, and optional verification in Microsoft Word.

**Architecture:** Load a package once, run named checks against a shared context, then apply one validation policy. Separate portable Word checks, schema conformance, and observed Microsoft Word behavior.

**Stack:** TypeScript ESM, Node 26, pnpm, Vitest/V8 coverage, JSZip, existing XML helpers, and libxml2-wasm.

## 1. Contracts and TDD workflow

### Required behavior

| Mode | Required checks | Meaning of success |
|---|---|---|
| Default / `word-valid` | Complete portable Word check set | No known Word-opening defect detected |
| `--strict` | Complete Word checks and applicable schemas | Word checks pass and schema validation completes successfully |
| `--quick` | Common-defect preset | Selected checks pass; report remaining unchecked coverage |
| `--checks <ids>` | Named checks and prerequisites | Selected checks pass; report remaining unchecked coverage |
| `--verify-word` | Requested portable mode plus real Word | Microsoft Word also opens the exact file without warnings, repair, recovery, or crashes |
| Explicit `lenient` | Existing compatibility behavior | Preserve existing callers during migration |

`--strict` cannot be combined with partial check selection. Quick and named checks are mutually exclusive. Reject conflicting CLI profile selections.

Word verification initially supports macOS using the existing probe. Other platforms return an explicit unavailable result when verification is requested.

PPTX keeps its current default behavior; Word-specific semantics apply to DOCX/DOCM.

### Shared interfaces

- Consolidate validation options so path-based validation, `Validate.run(bytes)`, CLI execution, and `jubarte.validate(bytes, options)` forward the same applicable options.
- Preserve `ValidationResult { valid, issues }`; extend the shared result with optional execution metadata rather than inventing another validator result shape.
- Execution metadata identifies the profile, executed checks, omitted checks, schema completion, and whether Word verification was requested or observed.
- `valid` means the requested checks passed. Partial results must display “Selected checks passed” and identify omissions.
- Preserve existing issue codes where their meaning remains correct. New structural findings receive stable codes independent of libxml2 message wording.
- Keep `XsdValidator` injection working end to end. Add a narrow injectable Word-verifier interface for deterministic tests.

### Red → green → refactor

For every behavior below:

1. **Red:** write one failing behavioral assertion.
2. Run it with coverage; confirm failure comes from the missing behavior.
3. **Green:** implement the smallest production change that passes.
4. Run the focused tests and affected regressions with coverage.
5. **Refactor:** remove duplication while retaining passing assertions.
6. Commit the completed slice, staging only its files.

Use:

```bash
pnpm exec vitest run <test-files> --coverage
```

Configure coverage reporting on failed runs. Every run must print line and branch percentages, including a summary such as `Coverage: 92% lines, 87% branches`.

New modules require at least **80% line / 70% branch** coverage. Validity policy, check scheduling, and result aggregation require **90% / 85%**. Preserve or improve existing measured coverage; do not delete assertions to accommodate changed defaults.

Unit tests use pure functions, inline XML, byte arrays, dependency-provided utilities, and centralized in-memory fakes. Filesystem, actual schema-engine, packaged-artifact, and Word tests are integration tests.

## 2. Ordered implementation tasks

### Task 1 — Establish coverage and independent acceptance fixtures

**Tests:** add `tests/validation-contract.integration.test.ts`; retain existing profile and fixture suites.

- [ ] **Red:** construct a schema-clean DOCX, append one undeclared byte after ZIP EOF, and assert both Word-valid and strict reject it with `zip-bytes-after-eocd`. Assert the unmodified control passes both.
- [ ] **Red:** assert every existing Word-specific rejection rule runs under strict.
- [ ] **Red:** assert default DOCX validation matches explicit `word-valid`.
- [ ] **Green:** share ZIP integrity checks across Word-valid and strict; invoke Word compatibility checks under strict.
- [ ] Keep the default flip behind completion of the later structural-check migration.
- [ ] Record baseline line/branch coverage before implementation. Configure V8 text, JSON-summary, HTML, and LCOV reports, including failed runs.

These assertions must distinguish a failed validity check from an unrelated schema or fixture error.

### Task 2 — Introduce the shared context, registry, and scheduler

**Production:** add focused modules under `src/lib/` for validation context, check registry, and execution policy. Keep existing validator classes as adapters.

- [ ] **Red:** an in-memory package containing one defect yields that defect from its named check without reading disk.
- [ ] **Red:** prerequisites execute once; malformed XML prevents dependent DOM checks from crashing.
- [ ] **Red:** two checks sharing a part reuse its bytes and parsed DOM.
- [ ] **Red:** unknown check IDs fail configuration validation rather than silently running nothing.
- [ ] **Green:** define check IDs, descriptions, applicability, prerequisites, and runners. Use one context containing package entries, original archive bytes when available, and lazy parse caches.
- [ ] **Green:** adapt existing checks incrementally; preserve structured results and stable ordering.
- [ ] Verify caches are scoped to one run and cannot leak documents across validations.

For unpacked directories, mark archive-only checks unavailable. Do not present that result as complete packed-file verification.

### Task 3 — Add built-in XML parsing with original-byte diagnostics

**Production:** extend shared XML helpers and the libxml2 adapter; keep DOM construction centralized.

- [ ] **Red:** malformed XML reports the original part path and available line/column.
- [ ] **Red:** supported UTF-8/UTF-16 encodings parse correctly from bytes.
- [ ] **Red:** undeclared namespace prefixes are diagnosed.
- [ ] **Red:** valid BOM-prefixed XML does not fail strict solely because of its BOM.
- [ ] **Red:** malformed content cannot pass through parser recovery.
- [ ] **Green:** use libxml2’s byte parser for validation and preserve structured parser diagnostics. Keep DOM manipulation through existing XML helpers.
- [ ] **Green:** disable external entity loading; resolve schema resources through explicitly configured local resources.
- [ ] Add actual libxml2 integration cases alongside pure diagnostic-mapping unit tests.

Preserve original bytes. Validation must not pretty-print, rewrite, or normalize the input document.

### Task 4 — Correct content-type resolution

**Coverage:** issues #208, #237, #249.

- [ ] **Red:** media covered only by an exact `Override` produces neither `ct-undeclared-ext` nor `ct-uncovered-part`.
- [ ] **Red:** escaped names and OPC-equivalent case variants resolve consistently.
- [ ] **Red:** a genuinely undeclared part still fails.
- [ ] **Green:** extract one deterministic resolver implementing `Override` precedence over extension `Default`.
- [ ] Use the resolver in every coverage check, including XML-root checks; remove inconsistent duplicated lookups.
- [ ] Verify the original affected lane specimens as integration regressions.

### Task 5 — Resolve markup compatibility without manufacturing errors

**Coverage:** issues #239, #247, #251, #254, #255, #258, #261, #262.

- [ ] **Red:** fallback `positionH` and `positionV` survive compatibility processing.
- [ ] **Red:** supported `Choice` wins over fallback; unsupported choices use fallback.
- [ ] **Red:** nested alternatives, inherited declarations, `ProcessContent`, and `MustUnderstand` follow compatibility semantics.
- [ ] **Red:** malformed compatibility markup remains diagnosed.
- [ ] **Green:** implement a deterministic compatibility processor operating on a validation-only clone.
- [ ] Replace blanket removal of foreign namespaces with compatibility-aware processing.
- [ ] Assert the original XML and all unselected branch content remain unchanged in the source package.

Select branches using declared supported namespace capabilities, not prefix spelling.

### Task 6 — Preserve schema-permitted chart extensions

**Coverage:** issues #230, #243, #248, #250, #253, #260.

- [ ] **Red:** a `c:ext` containing a schema-permitted foreign child does not become empty during preprocessing.
- [ ] **Red:** a genuinely empty source `c:ext` still fails the applicable schema.
- [ ] **Green:** preserve wildcard-admitted payloads and let the schema apply its declared lax validation.
- [ ] Verify compatibility preprocessing and extension handling together on the original chart specimens.

Do not solve this by deleting every empty extension after preprocessing; that would conceal genuine defects.

### Task 7 — Replace XSD-message heuristics with structural Word checks

- [ ] **Red:** each currently recognized Word-blocking structure fails even when the XSD engine returns differently worded diagnostics.
- [ ] Cover misplaced runs/properties/sections/table properties, invalid picture structure, missing drawing requirements, and invalid custom-property scalars.
- [ ] Pair every rejection case with a nearby Word-tolerated control.
- [ ] **Green:** implement namespace- and parent-aware checks against the context DOM.
- [ ] Remove `isWordBlockingXsdIssue` message matching only after every currently covered rejection has an independent check.
- [ ] Assert complete Word-valid validation does not compile the full XSD bundle.

Migrate the remaining existing Word rules into the registry, retaining their current positive and negative regression fixtures.

### Task 8 — Implement named checks for all remaining issue families

Create one module and unit-test file per coherent rule family. Run each row as its own red/green slice.

| Reports | Red assertion | Green implementation |
|---|---|---|
| #235, #245, #252, #256, #257, #263 | Illegal direct `m:r/w:ins` is identified; legal math wrappers remain accepted | Check exact namespace and parent relationships; provide permitted-wrapper guidance |
| #213, #259 | Direct `m:r/w:del` is distinguished from legal math deletion marks | Diagnose nesting without changing revision content |
| #210, #215, #220, #221, #225, #229, #238, #242 | Empty or invalid shading color is reported; omitted optional color passes | Validate declared color value space |
| #209, #219, #224, #226, #233, #240 | `gridCol@type` is reported without losing its width | Check permitted grid-column attributes |
| #211, #216, #222, #223, #227, #228, #231, #232, #234, #236, #241, #246 | Cell-margin order is diagnosed in document stories and styles | Apply the bundled schema sequence to `tblCellMar` and `tcMar` |
| #214 | Each missing required page-margin attribute is named | Check required attributes without inventing values |
| #212, #217, #218 | Missing table properties are detected in body and header tables | Check required table children and positions |
| #244 | Wrong textbox-content namespace is reported before preprocessing | Distinguish `wp`, `wps`, and VML textbox models; preserve text |

For historically Word-clean/schema-invalid specimens, assert default-mode advisories and strict-mode rejection. Historical observations retain provenance; they are not newly verified Word results.

### Task 9 — Make strict schema validation complete

- [ ] **Red:** ISO Strict fixtures cannot pass through `xsd-strict-skipped`.
- [ ] **Red:** missing required schemas/imports prevent strict success.
- [ ] **Red:** identical defects in `--original` and modified documents still make strict invalid.
- [ ] **Red:** multiple occurrences of the same error retain distinct locations.
- [ ] **Green:** bundle official Strict schemas and missing imports with source, version, license, and hashes.
- [ ] Dispatch using namespace and part identity; resolve the complete import closure locally.
- [ ] Remove broad message suppression and baseline-error subtraction from strict validity decisions.
- [ ] Preserve all structured schema issues; accurately distinguish original-source and transformed-view locations.
- [ ] Verify both Transitional and Strict schema trees with network access unavailable.

Arbitrary custom XML is not automatically a missing OOXML schema. Respect declared extension and wildcard rules.

### Task 10 — Finish mode policy and expose quick/named checks

- [ ] **Red:** strict success implies complete Word-check success and completed schema validation.
- [ ] **Red:** schema-only hygiene findings do not become default Word-opening errors.
- [ ] **Red:** quick output identifies omitted checks and never prints “All validations passed.”
- [ ] **Red:** conflicting selectors are rejected.
- [ ] **Green:** apply the shared policy and switch the DOCX/DOCM default to `word-valid`.
- [ ] Add `--strict`, `--quick`, `--checks`, and `--list-checks`.
- [ ] Define the quick preset as package integrity, XML syntax, content types, relationships, table structure, and drawing structure.
- [ ] Expose other named groups for story structure, comments, tracked changes, math, and Word settings.
- [ ] Assert quick mode performs no full-schema compilation.

### Task 11 — Complete API and CLI parity

- [ ] **Red:** injected XSD validators reach path, class, and helper entrypoints.
- [ ] **Red:** equivalent bytes/path inputs return equivalent issues and execution metadata.
- [ ] **Red:** `--help` succeeds without initializing XML/XSD engines.
- [ ] **Red:** JSON output contains every issue, including warnings and locations, without human text mixed into stdout.
- [ ] **Green:** consolidate option forwarding and result aggregation; extend `jubarte.validate` with optional options.
- [ ] Add `--json`; keep human rendering in the CLI layer.
- [ ] Verify explicit legacy `lenient` behavior and unchanged PPTX default behavior.

### Task 12 — Add optional real-Word verification

- [ ] **Red:** scripted verifier observations distinguish clean opening, repair warning, open error, crash, password prompt, unknown dialog, and timeout.
- [ ] **Red:** an unrelated open document cannot satisfy verification.
- [ ] **Red:** cleanup closes only the document opened by this run and never saves it.
- [ ] **Green:** extract the existing macOS probe behind an injectable adapter.
- [ ] Open a temporary copy without repair; verify document identity and absence of warnings through a fixed observation period.
- [ ] Use a 45-second total timeout with fake clock/transport control in unit tests.
- [ ] Record file hash, Word version/build, platform, outcome, and observation time.
- [ ] Treat unavailable or inconclusive verification as unsatisfied when explicitly requested.
- [ ] Remove the existing probe’s automatic global force-quit behavior from this path.

Run actual Word tests separately from portable CI.

## 3. Corpus, packaging, and GitHub delivery

### Task 13 — Lock independent expectations and package behavior

- [ ] Add every affected lane specimen to integration coverage, with minimal single-defect controls where possible.
- [ ] Extend the manifest with Word-valid expectations and hash-bound Word provenance.
- [ ] **Red:** changed file hashes invalidate historical observations; unknown/timeouts are excluded from clean/rejected accuracy counts.
- [ ] **Green:** update manifest readers/writers without discarding existing Word evidence.
- [ ] Review expectation changes individually. Never regenerate a manifest solely from current output and treat that output as proof of correctness.
- [ ] Verify installed-package imports, trivial validation, schema asset availability, and CLI help from a package artifact.
- [ ] Update README, AGENTS, CLAUDE, migration notes, and changelog together.

Final gates:

```bash
pnpm exec vitest run --coverage
pnpm exec tsc --noEmit
pnpm run check
pnpm run build
pnpm exec tsx src/scripts/office/validate.ts --help
```

Add the package-artifact smoke and optional macOS Word integration lane. Report their results separately.

### Task 14 — Update GitHub issues with the executable solution

- [ ] Refresh issue state before publication.
- [ ] Update every existing report #208–#263 with its applicable task, failing assertion, minimal solution, mode behavior, coverage command, and acceptance criteria.
- [ ] Preserve original evidence, labels, assignees, and discussion; cross-link shared fixes.
- [ ] Create only missing capability issues for policy/API parity, check discovery, built-in XML diagnostics, schema completeness, Word verification, and corpus gates.
- [ ] Each issue distinguishes planned work, implemented work, passing portable checks, and observed Word verification.

## 4. Completion criteria and boundaries

- Every production behavior has recorded red and green evidence.
- All 56 reports are covered by explicit regression assertions.
- Strict cannot bypass Word checks or silently skip required schemas.
- Portable success is distinguished from observed Microsoft Word success.
- Quick results disclose omissions.
- Source packages remain unchanged by validation.
- No new automatic repair behavior is introduced.
- Coverage thresholds, package smoke checks, and relevant integration tests pass.
- Existing unrelated working-tree changes remain untouched.

This turn produces the implementation plan only. Implementation and GitHub publication remain pending.
