# Changelog

All notable changes to `docx-validate` are documented here. Format follows
[Keep a Changelog](https://keepachangelog.com/en/1.1.0/) and the project
adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

### Added

- Registered `opc-part-name-escaped-unreserved` package check for OPC [M1.8], including packed and directory inputs, full and partial runs, and diagnostics before extraction. Escaped unreserved characters in stored names are errors under strict and Word-valid DOCX/DOCM, and warnings under lenient and Word-valid PPTX/XLSX ([#408](https://github.com/jandira-tech/docx-validate/issues/408)).

- Diagnose equivalent stored part names and multiple content-types stream spellings with the selectable `opc-equivalent-part-names` package check. Full runs report it before extraction. Strict, and word-valid for DOCX/DOCM, report an error, because Word 16.115 offers recovery for every duplicate spelling probed (case, unreserved escape, backslash, two streams), even with identical bytes. Lenient, and word-valid for PPTX/XLSX, warn ([#351](https://github.com/jandira-tech/docx-validate/issues/351)).

- Named Word rules in `WORD_RULE_CHECKS`, prerequisite scheduling with per-part XML gating, `--quick`, `--checks`, `--list-checks`, `--strict`, `--json`, and shared API execution metadata ([#267](https://github.com/jandira-tech/docx-validate/issues/267), [#271](https://github.com/jandira-tech/docx-validate/issues/271), [#272](https://github.com/jandira-tech/docx-validate/issues/272), [#306](https://github.com/jandira-tech/docx-validate/issues/306), [#307](https://github.com/jandira-tech/docx-validate/issues/307)).
- Structural Word checks for misplaced content/properties, incomplete drawings/pictures, unknown elements, namespace-less links, and empty custom-property filetimes replace XSD-message heuristics ([#269](https://github.com/jandira-tech/docx-validate/issues/269)).
- Advisory checks for math revisions, shading colors, grid-column attributes, cell-margin order, required table/page properties, and textbox namespaces ([#235](https://github.com/jandira-tech/docx-validate/issues/235), [#213](https://github.com/jandira-tech/docx-validate/issues/213), [#210](https://github.com/jandira-tech/docx-validate/issues/210), [#209](https://github.com/jandira-tech/docx-validate/issues/209), [#211](https://github.com/jandira-tech/docx-validate/issues/211), [#212](https://github.com/jandira-tech/docx-validate/issues/212), [#214](https://github.com/jandira-tech/docx-validate/issues/214), [#244](https://github.com/jandira-tech/docx-validate/issues/244)).
- Optional `--verify-word` / `verifyWord` verification with macOS Microsoft Word, exact-byte temporary copies, classified outcomes, hash/version/platform/timing metadata, and document-scoped cleanup ([#273](https://github.com/jandira-tech/docx-validate/issues/273)).
- Packaged `docx-validate` bin at `dist/cli.mjs`, a silent library entry, and an installed-package smoke covering imports, validation, schema assets, and CLI help ([#314](https://github.com/jandira-tech/docx-validate/issues/314), [#274](https://github.com/jandira-tech/docx-validate/issues/274)).
- Hash-bound manifest Word observations with preserved history and independent `fixtures-all-word-valid` regression expectations; unknown, timeout, and stale evidence is excluded from accuracy counts ([#274](https://github.com/jandira-tech/docx-validate/issues/274), [#313](https://github.com/jandira-tech/docx-validate/issues/313)).
- Per-module coverage gates through `test:coverage-gate` and CI, plus a `Coverage: N% lines, M% branches` summary on every coverage-enabled Vitest run, including failures ([#266](https://github.com/jandira-tech/docx-validate/issues/266), [#271](https://github.com/jandira-tech/docx-validate/issues/271), [#274](https://github.com/jandira-tech/docx-validate/issues/274)).

### Changed

- **Breaking:** DOCX/DOCM default to `word-valid`; pass `--profile lenient` or `{ profile: "lenient" }` for the previous policy. PPTX retains its lenient default. DOCX/DOCM Word-valid runs structural checks without document XSD validation, with `schemaCompleted: false` and an `xsdWalk` that has not started; strict includes the Word rules and applicable schemas ([#266](https://github.com/jandira-tech/docx-validate/issues/266), [#269](https://github.com/jandira-tech/docx-validate/issues/269), [#271](https://github.com/jandira-tech/docx-validate/issues/271)).

- Document that Word-valid's `xml-encoding` blocking also covers real encodings libxml2 cannot decode (`macintosh`, `x-mac-cyrillic`, `IBM866`, `UTF-7`), which Word 16.115 refuses too ([#410](https://github.com/jandira-tech/docx-validate/issues/410)). The third real-Word session is recorded in `docs/plans/evidence/word-probes-2026-10-24.md`.

### Fixed

- Report an `Override` whose `PartName` is not a valid OPC part name, with the new `ct-override-partname-invalid` content-types check. A dot segment, a query, a fragment, a backslash, a missing leading slash, an empty segment, a percent-escaped separator ([M1.7]) or unreserved character ([M1.8]), or a character outside the RFC 3986 `pchar` set is reported at the stored content-types stream, quoting the PartName. Such a PartName names no part, so the declaration is malformed whatever its content type, and the resolution is left untouched. Word 16.115 offers to recover every such package, even when nothing is stored under that name and the content type is well formed, so the finding is an error, and Word-blocking, under word-valid for DOCX/DOCM; strict reports it as an error, and lenient and the unobserved PPTX/XLSX report a warning ([#414](https://github.com/jandira-tech/docx-validate/issues/414)).

- Stored names whose segments percent-encode `/` or `\` (`%2F`, `%2f`, `%5C`, `%5c`) are now reported by the `opc-part-name-escaped-slash` package check (OPC [M1.7]), in full, quick, and named-check runs, before extraction and on refused package views. Word 16.115 offers recovery for an unreferenced `word%2Fxtra.xml`, refuses `word%5Cxtra.xml`, and opens the well-formed `word/xtra.xml` control clean, so the finding is Word-blocking under word-valid for DOCX/DOCM, an error under strict, and a warning under lenient and word-valid PPTX/XLSX ([#436](https://github.com/jandira-tech/docx-validate/issues/436)).

- Word-valid now rejects a DOCX/DOCM package that Word 16.115 offers to recover: `ct-missing` and `word-main-document-part` are Word-blocking in `isWordBlockingIssue`. A package without `[Content_Types].xml`, an empty ZIP, a package without its main part and one whose `officeDocument` relationship is removed now fail a full word-valid run instead of passing with warnings. Quick and named runs report `word-main-document-part` as an error too; `ct-missing` comes only from the full run's content-types pass, which no partial selection reaches. Strict already rejected them; lenient is unchanged ([#435](https://github.com/jandira-tech/docx-validate/issues/435), [#380](https://github.com/jandira-tech/docx-validate/issues/380)). The observations are recorded in `docs/plans/evidence/word-probes-2026-10-24.md` (sessions 3c and 4).

- Diagnose repeated content-type Default extensions and Override part names in full, quick and named runs. Duplicates are errors under strict and Word-valid DOCX/DOCM, and warnings under lenient and Word-valid PPTX/XLSX ([#431](https://github.com/jandira-tech/docx-validate/issues/431)).

- Validate relationships-part roots and direct children without an XSD walk. Word-valid rejects malformed package relationships and malformed relationships parts under `word/`, and warns on other sidecars; strict rejects all of them. Main-part discovery and other semantic relationship consumers read only direct `Relationship` children of a valid `Relationships` root under Word-valid and strict; legacy lenient traversal is preserved ([#433](https://github.com/jandira-tech/docx-validate/issues/433), [#434](https://github.com/jandira-tech/docx-validate/issues/434)).

- Keep `xml-syntax`, `xml-encoding`, and `xml-undeclared-prefix` as Word-valid errors on core, extended, and custom properties parts named by their package relationships. Word reads those targets and offers recovery for the observed defects; unreferenced `docProps/*.xml` findings stay warnings. Full, quick and named runs resolve the targets by OPC part-name equivalence; strict verdicts are unchanged ([#409](https://github.com/jandira-tech/docx-validate/issues/409)).

- Stored names are read from the ZIP central directory, in order and keeping duplicates, instead of from the loaded entry view, which keeps only the last entry under a given name and rewrites `.`/`..`/empty segments before that. Two entries stored under the identical name are now reported by `opc-equivalent-part-names` (naming the spelling once per entry), and a stored `..` name that resolution would hide behind the entry it resolves to is refused in either ZIP order, instead of the verdict depending on which entry came last. The identical-name shape is Word-blocking, as #351's spellings are: Word 16.115 offers recovery whether both entries are valid or only the second is. The package view still reads the last stored entry, which the finding's message says ([#411](https://github.com/jandira-tech/docx-validate/issues/411)).

- Resolve relationship-id and hyperlink sidecars by OPC part-name equivalence, including `.RELS`, for packed and directory inputs and full, named and quick runs. Escaped names such as `.%72els` remain invalid under Word-valid and strict through the OPC [M1.8] check ([#402](https://github.com/jandira-tech/docx-validate/issues/402)).

- A quick or named-check run now reports an archive whose package view cannot be read — a stored item name that places no part (`word//x.xml`, `c:doc/x.xml`, a `.` or `..` segment) or an unreadable entry — as `package-extract-failed`, with the same code and message the full run reports, instead of `package-open-failed` ([#413](https://github.com/jandira-tech/docx-validate/issues/413)).

- On such an archive the run still executes the selected stored-name checks (`opc-equivalent-part-names`, `opc-part-name-escaped-unreserved`), which read only the stored entry names and report the same findings the full run reports, and every selected check that could not run is listed in `omittedChecks` instead of appearing in neither `executedChecks` nor `omittedChecks` ([#420](https://github.com/jandira-tech/docx-validate/issues/420)).

- Full validation stages directory inputs with literal backslashes in stored file names under package part names, sharing the packed-input extraction guards and duplicate selection. `--auto-repair` writes modified bytes and deletions back to their original stored names; untouched files remain byte-identical and validation without repair leaves the source unchanged ([#374](https://github.com/jandira-tech/docx-validate/issues/374)).

- `validate()` no longer throws on staged-directory I/O: a stored file that cannot be read into the package view fails with `package-extract-failed` naming the file, and a failed `--auto-repair` write-back — a read-only or vanished stored file — is reported as a `repair-write-back-failed` error while the remaining repairs are still written back ([#412](https://github.com/jandira-tech/docx-validate/issues/412)).

- Scope `word-content-type-invalid` to stored parts. In Word 16.115, an Override for a missing part and an unused Default open clean; an invalid content type on an existing part is refused or recovered. The observations, with hashes, are in `docs/plans/evidence/word-probes-2026-10-24.md` ([#401](https://github.com/jandira-tech/docx-validate/issues/401)).

- Resolve the content-types ZIP stream ASCII case-insensitively, without percent decoding, across Word checks, content-type coverage, XSD dispatch, and semantic diagnostics. Findings retain the stored stream name. Of two stored spellings, `[Content_Types].xml` is read, otherwise the first in code-unit order (System.IO.Packaging reads the last in ZIP order), and the second is now diagnosed as a strict error or word-valid/lenient warning ([#351](https://github.com/jandira-tech/docx-validate/issues/351)). In a directory input, a folder named `[Content_Types].xml` is now `ct-missing` rather than `ct-parse` ([#396](https://github.com/jandira-tech/docx-validate/issues/396)).

- The relationship checks (`rels-broken`, `rels-unreferenced`, `rels-target-missing`) resolve targets by OPC part-name equivalence, not through the filesystem, so ASCII case variants (`WORD/document.xml`) and unreserved percent escapes (`word/%64ocument.xml`) name the stored part on every filesystem, and a reserved escape (`word%2Fdocument.xml`) still names no part. `_RELS/.RELS` is read as the package relationships part. Target resolution drops a query or fragment and reads `\` as `/`, as it does for the main part: `styles.xml#frag`, `styles.xml?q=1` and `word\styles.xml` now resolve instead of reporting `rels-broken` ([#361](https://github.com/jandira-tech/docx-validate/issues/361)). Every relationships pass and the XSD walk now recognise a relationships part by its OPC-equivalent name (any ASCII case, unreserved escapes decoded), including `.RELS` and `.%72els`. `.XML`-spelled parts are now schema-validated and seen by the passes that iterate the part list ([#398](https://github.com/jandira-tech/docx-validate/issues/398)). Missing targets and unreferenced parts keep their existing diagnostics ([#346](https://github.com/jandira-tech/docx-validate/issues/346), [#393](https://github.com/jandira-tech/docx-validate/issues/393)).

- Word-valid now fails a part it cannot decode (`xml-encoding`: an unknown declared encoding, or bytes that do not decode in the declared one), with the same path policy as `xml-syntax`: under `word/`, the content-types stream and `_rels/.rels`. Word 16.115 refuses a `word/document.xml` declaring `X-UNKNOWN-ENCODING` and opens the `windows-1252` control clean. The other paths and the invalid-bytes case follow the `xml-syntax` policy by analogy; their probes are queued in `docs/plans/evidence/word-probes-2026-10-24.md` ([#397](https://github.com/jandira-tech/docx-validate/issues/397)).

- Word-valid now fails a package whose `_rels/.rels` does not parse (zero bytes, truncated, or an undeclared prefix), in full and partial runs: `xml-syntax` and `xml-undeclared-prefix` on the package relationships part, matched by OPC part name, are Word-blocking. Word 16.115 offers to recover each such specimen and opens the readable control clean; the observations, with file hashes, are in `docs/plans/evidence/word-probes-2026-10-24.md` ([#388](https://github.com/jandira-tech/docx-validate/issues/388)).

- The full run's XML pass now checks every part whose extension is `.xml` or `.rels` as OPC compares part names (in any case, or with escaped unreserved characters such as `word/comments.%78ml`), as the `xml-syntax` check of quick and named runs does. A malformed `word/COMMENTS.XML` used to pass a full Word-valid run (the comment checks only warn about it) and fail `--quick`; both now fail it. The pass names each finding by its part name, so a directory file stored as `word\foo.xml` is judged as the `word/` part it is ([#390](https://github.com/jandira-tech/docx-validate/issues/390)).

- A lenient named run that selects the comment checks stops where `xml-syntax` rejects a part, as the full lenient run stops at its first malformed part. It no longer reports defects of the parts that parsed, which the full lenient run never reaches; `executedChecks` lists only `xml-syntax` ([#391](https://github.com/jandira-tech/docx-validate/issues/391)).

- Word-valid treats a malformed part under `WORD/` (any case, or with percent-escaped unreserved characters) as under `word/`: OPC part names compare that way. A package whose main document part `WORD/document.xml` does not parse no longer passes. A malformed content-types stream blocks in any case (`[CONTENT_TYPES].XML`), as System.IO.Packaging matches that ZIP item ([#392](https://github.com/jandira-tech/docx-validate/issues/392)).

- Named comment checks also run XML syntax diagnostics and suppress restated parse findings, preserving the full run's profile verdict for malformed parts ([#385](https://github.com/jandira-tech/docx-validate/issues/385)).

- Named `comments` selection now runs comment marker and threading integrity checks while full DOCX validation preserves its findings under every profile. The comment checks now also judge a comment part stored under an OPC-equivalent name (`word/comments.XML`, percent-encoded characters), as every other check does; the legacy rules ignored such parts, so a package that passed can now report comment findings ([#384](https://github.com/jandira-tech/docx-validate/issues/384)).

- Parse XML from original bytes without recovery, preserve available line/column diagnostics, support UTF-8/UTF-16, and diagnose undeclared prefixes and encoding failures with stable codes. BOMs alone never fail validation ([#268](https://github.com/jandira-tech/docx-validate/issues/268), [#290](https://github.com/jandira-tech/docx-validate/issues/290)).
- Resolve content types with exact Override precedence and canonical OPC part-name comparison, including reserved-escape distinctions ([#208](https://github.com/jandira-tech/docx-validate/issues/208), [#297](https://github.com/jandira-tech/docx-validate/issues/297)).
- Select markup-compatibility branches on validation-only clones, preserve schema-permitted extension payloads, and judge Word rules on the content Word consumes. Diagnose invalid tokens, unbound prefixes, and loss of required main roots without duplicate root Ignorable findings ([#239](https://github.com/jandira-tech/docx-validate/issues/239), [#250](https://github.com/jandira-tech/docx-validate/issues/250), [#287](https://github.com/jandira-tech/docx-validate/issues/287), [#291](https://github.com/jandira-tech/docx-validate/issues/291), [#296](https://github.com/jandira-tech/docx-validate/issues/296), [#309](https://github.com/jandira-tech/docx-validate/issues/309)).
- Validate ISO Strict OOXML under strict with bundled schemas and local Dublin Core imports. Preserve repeated schema-error locations, bypass strict message suppression and original-baseline subtraction, and keep infrastructure failures out of baseline subtraction ([#270](https://github.com/jandira-tech/docx-validate/issues/270), [#289](https://github.com/jandira-tech/docx-validate/issues/289)).
- Reject a main document part whose effective root is not a WordprocessingML `document` under Word-valid and strict (`word-document-root`). Word 16.115 refuses to open a foreign root and silently drops all content under a root that markup compatibility ignores ([#317](https://github.com/jandira-tech/docx-validate/issues/317)).
- Report a missing main document part with `word-main-document-part`: strict rejects it, full Word-valid warns pending real-Word evidence, and lenient stays unchanged ([#380](https://github.com/jandira-tech/docx-validate/issues/380)).
- Probe the XSD engine only for runs that walk schemas: a broken engine no longer fails a default (word-valid) DOCX run, and `--profile strict` still fails loudly ([#318](https://github.com/jandira-tech/docx-validate/issues/318)).
- `metadata.executedChecks` lists only the checks a run executed: lenient no longer reports the Word rules it never runs, and a rule skipped by a failed prerequisite is not listed ([#319](https://github.com/jandira-tech/docx-validate/issues/319)).
- `validate --help`, the `xsd-strict-skipped` message and the option docs describe the per-format default and the bundled ISO Strict schemas ([#321](https://github.com/jandira-tech/docx-validate/issues/321)).
- `verifyInWord` reports `crash` only with a Microsoft Word crash report written after the open; a Word process that vanishes without one (an external kill or quit) is `unavailable`. Every `crash` behind the issue was an external SIGKILL ([#305](https://github.com/jandira-tech/docx-validate/issues/305)).
- Race `verifyInWord`'s version and availability calls against its budget, so a hanging transport call ends as `timeout` ([#320](https://github.com/jandira-tech/docx-validate/issues/320)).
- Count only Word's verdicts as rejections: `crash`, `unavailable` and unmapped outcome spellings are unattributable, and verifier outcomes map to the manifest's spellings ([#316](https://github.com/jandira-tech/docx-validate/issues/316)).
- Quick and named-check runs fail an archive that does not open with `package-open-failed`, as complete runs do, instead of passing the selected checks on an empty package. Their `executedChecks` lists the checks that ran ([#332](https://github.com/jandira-tech/docx-validate/issues/332)).
- Quick and named-check runs apply their profile as a complete run does, so they no longer fail a package the complete run passes. Under `lenient` they run neither the Word rules nor the ZIP-envelope checks. Naming one of those checks, or only groups made of them, is refused as a mode conflict, before the target is read, like every other mode conflict. (`--checks package` is not such a selection: the package group also contains `opc-equivalent-part-names` and `opc-part-name-escaped-unreserved`, which a lenient run executes.) Under `word-valid` a DOCX check's error that Word tolerates is a warning, through the same `isWordBlockingIssue` predicate the complete run uses, now in `validators/word-blocking.ts`. A malformed `docProps/app.xml` is then a warning in both runs ([#381](https://github.com/jandira-tech/docx-validate/issues/381)).
- `word-attached-toolbar-utf8` reports every referenced toolbar part that carries the UTF-8 replacement sequence, once each, instead of stopping at the first ([#323](https://github.com/jandira-tech/docx-validate/pull/323)).
- A partial run whose archive does not open lists in `omittedChecks` only the checks it did not select, as every other run does; a selected check the archive kept from running is no longer called an omission ([#337](https://github.com/jandira-tech/docx-validate/pull/337)).
- A complete Word-valid or strict run of a package that opens but fails to extract lists the ZIP envelope checks it ran in `executedChecks`, in registry order, instead of none ([#335](https://github.com/jandira-tech/docx-validate/pull/335)).
- A Word verifier that throws yields an `unavailable`, unsatisfied verification instead of rejecting `validate` ([#333](https://github.com/jandira-tech/docx-validate/issues/333)).
- A partial CLI run prints `Selected checks passed`, never `All validations PASSED!`, even when it omitted no registered check: a PPTX or XLSX `--quick` run used to claim a complete pass on a deck the full run rejects. An `Also not run:` line names the format validator's own rules, where a full run walks schemas the schema walk ([#339](https://github.com/jandira-tech/docx-validate/issues/339)), and for a DOCX given `--original` the redlining validation against it ([#353](https://github.com/jandira-tech/docx-validate/issues/353)).
- `profileWalksSchemas` takes the format: a PPTX walks its schemas under every profile, so a `word-valid` PPTX run now probes the XSD engine at startup, and an XLSX run, which compiles no schema, no longer does ([#341](https://github.com/jandira-tech/docx-validate/issues/341)).
- The `schemaCompleted` doc scopes "always false under `word-valid`" to DOCX/DOCM ([#340](https://github.com/jandira-tech/docx-validate/issues/340)).
- Let an open that overran its budget settle before the cleanup close, so Word is never left holding a document opened after its close ([#328](https://github.com/jandira-tech/docx-validate/issues/328)).
- Stamp `observedAt` and the crash-report open instant from wall time (`wallNow`, default `Date.now`) instead of the deadline clock, and stamp an unpacked directory's `unavailable` verification with the current time instead of the epoch ([#329](https://github.com/jandira-tech/docx-validate/issues/329)).
- Return `unavailable` when a Word dialog is already showing before verification, without opening or closing anything; dismiss dialogs raised by this run during macOS cleanup so later probes remain attributable ([#315](https://github.com/jandira-tech/docx-validate/issues/315)).
- `word-document-root` finds the main part behind an OPC-equivalent officeDocument target (`WORD/document.xml`, `word/%64ocument.xml`), which let a foreign main root pass Word-valid. The schema walk's main-part test compares canonical part names, so an escaped reserved character (`word%2Fdocument.xml`) no longer names `word/document.xml` ([#344](https://github.com/jandira-tech/docx-validate/issues/344)).

- `word-document-root` and the schema walk's main-part test resolve the officeDocument target through one shared resolver (`resolvePartName`), so they agree: a backslash target (`word\document.xml`) now names the main part for the schema walk too, and an ignored main root behind it is an error again rather than info. Of two equivalent entries, only the one the target names is main; `findEquivalentPart` picks it whatever order the names come in ([#349](https://github.com/jandira-tech/docx-validate/issues/349)). The schema walk resolves its main parts once per validator rather than walking the unpacked tree for every ignored root ([#355](https://github.com/jandira-tech/docx-validate/issues/355)), reads `\` as `/` in entry names as it does in targets, so a main part stored as `word/a\b.xml` stays main ([#357](https://github.com/jandira-tech/docx-validate/issues/357)), and `resolvePartName` returns `undefined` for a reference that is empty, names the package root or climbs above it ([#356](https://github.com/jandira-tech/docx-validate/issues/356)). A query or fragment in the target is dropped before it is resolved, as System.IO.Packaging does, so `word/document.xml#frag` still names the main part for both checks ([#361](https://github.com/jandira-tech/docx-validate/issues/361)), and the schema walk reads `_rels/.rels` as the package view does: under a backslash name and in UTF-16 too ([#362](https://github.com/jandira-tech/docx-validate/issues/362)). Both checks find the package relationships by OPC equivalence, so `_RELS/.RELS` names them too ([#364](https://github.com/jandira-tech/docx-validate/issues/364)). When several entries name one part (`_rels\.rels` beside `_rels/.rels`), the archive, the extracted directory and the schema walk keep the same one, whatever the entry order: the entry stored under the part name itself, otherwise the smaller name. Before, the archive kept whichever entry finished decompressing last (in practice the last in ZIP order) and the directory the last in sorted order, so a quick run and a full run could judge different main parts ([#366](https://github.com/jandira-tech/docx-validate/issues/366)). The full run's own checks list the walked files the same way, so they no longer judge an entry the package view drops ([#368](https://github.com/jandira-tech/docx-validate/issues/368)). The archive still reads every entry, so an unreadable entry rejects it even when another entry names the same part, as extraction does ([#369](https://github.com/jandira-tech/docx-validate/issues/369)). A full run extracts the package view itself: each part once, under its part name, with the entry the view keeps. A part stored only as `word\document.xml` is then schema-validated and finds its `.rels` sidecar, and a stray entry beside its part can no longer overwrite it on Windows ([#371](https://github.com/jandira-tech/docx-validate/issues/371)). A ZIP item whose name starts with a slash or names a drive (`/word/x.xml`, `\word\x.xml`, `C:/x.xml`) names no part. The ZIP format forbids such names, and System.IO.Packaging ignores them, so a quick run now refuses the package (`package-open-failed`) as a full run already did (`package-extract-failed`), on every platform ([#372](https://github.com/jandira-tech/docx-validate/issues/372)). When the filesystem holds two parts as one file (a case or Unicode-normalization variant on macOS APFS), the full run stops with `package-extract-failed` instead of judging the merged bytes ([#373](https://github.com/jandira-tech/docx-validate/issues/373)). The full run also stops when such a filesystem merges two parts' folders, because the later part would otherwise be judged under a name the package does not hold ([#376](https://github.com/jandira-tech/docx-validate/issues/376)). Names are judged as stored in the archive, before JSZip rewrites them, so `/../word/document.xml` is refused as a leading-slash name instead of loading as `word/document.xml`. An empty or dot-only segment (`word//x.xml`, `docProps\nested\..\custom.xml`) is refused as well, because OPC forbids it and the two views would otherwise judge different names ([#378](https://github.com/jandira-tech/docx-validate/issues/378)). The schema walk matches part folders and file names ASCII case-insensitively, as OPC compares part names, so parts under `WORD/` are schema-validated, after the same markup-compatibility processing as parts under `word/` ([#377](https://github.com/jandira-tech/docx-validate/issues/377)).
- `word-attached-toolbar-utf8` reports a toolbar part under the package entry its target names, so a backslash or case-variant spelling of a part already reported adds no second finding and no finding names a part that does not exist; a case-variant target alone now finds its part ([#350](https://github.com/jandira-tech/docx-validate/issues/350)). The check builds its equivalence index (`equivalentPartLookup`) once per run instead of scanning every entry for every reference ([#358](https://github.com/jandira-tech/docx-validate/issues/358)), skips an External `attachedToolbars` relationship, and drops a target's query before resolving it, so it no longer reports an entry that the relationship does not name ([#359](https://github.com/jandira-tech/docx-validate/issues/359)).

### Known limitations

- Of equivalent stored part names, the package view still validates one entry, chosen deterministically; the content of the others is not validated. Case/normalization collisions can still stop full extraction on such a filesystem ([#351](https://github.com/jandira-tech/docx-validate/issues/351)).

### Security

- Replace the quadratic XML pretty-printer suffix regex with a linear scan, preserving formatting while preventing repeated tag starts in comments from causing excessive work (CodeQL alert #2).

## [0.53.0-rc.1] — 2026-10-15

Changes since the latest GitHub release, `v0.52.0` (2026-10-06), through 2026-10-15. This prerelease is published under npm's `next` tag; it does not advance `latest`.

### Added

- Word-valid diagnostics for ZIP trailing bytes and EOCD comment overruns; malformed drawing radii, extents, anchors, and empty drawings; malformed text, table rows/cells, and content-control IDs; hyperlink and duplicate singleton-part relationships; undeclared markup-compatibility prefixes; numbering picture-bullet and footnote references; and attached-toolbar corruption. Broken/clean fixture pairs and 144 recorded Word assessments document the evidence ([#192](https://github.com/jandira-tech/docx-validate/pull/192)).
- `validate()` results expose an `xsdWalk` trace showing whether XSD started, which parts reached the engine, and how many parts were eligible ([#191](https://github.com/jandira-tech/docx-validate/pull/191)).
- `pnpm run release:check` validates versions, types, lint/format, coverage tests, build output, and packed entry points/declarations/schemas. `pnpm run release:publish <tag>` dispatches GitHub Actions rather than publishing with local npm credentials.

### Changed

- Word-valid skips XSD when a cheaper check already rejects a document, and otherwise stops at the first surfaced Word-blocking schema error. Lenient and strict retain complete XSD walks. Schema-valid parts use a status-only WASM path; failures retain messages ([#191](https://github.com/jandira-tech/docx-validate/pull/191)).
- Redlining validation uses iterative DFS with mixed-namespace containment and drawing exclusions. Nested insertions emit one diagnostic per offending text node. End-to-end benchmark results are mixed; no fixed speedup is promised ([#194](https://github.com/jandira-tech/docx-validate/pull/194)).
- New Word-openability checks share read-only XML parses within each validation run. ZIP record searches use native signature scanning while retaining concatenated-archive detection.
- Publishing uses the existing `publish.yml` / `npm-publish` OIDC identity on a GitHub-hosted runner, without dependency caches. The workflow verifies the release tag against the package and chooses `next` for prereleases or `latest` for stable releases.

### Fixed

- Accept both npm 11 array metadata and npm 12 package-keyed metadata in the prerelease package guard. The initial GitHub candidate, `v0.53.0-rc.0`, stopped at this check and was never published to npm.
- Guard temporary original-part extraction with the shared path-containment resolver and retain nested paths for schema selection ([#193](https://github.com/jandira-tech/docx-validate/pull/193)). This hardens the boundary; a critical exploit through normal JSZip loading was not demonstrated.
- Use stable DOM snapshots in mutation-sensitive loops and normalize PowerPoint notes relationship targets ([#190](https://github.com/jandira-tech/docx-validate/pull/190)).
- Preserve Word-blocking XSD errors beyond the reporting cap; keep schema/engine failures blocking; guard private WASM pointers, reuse callbacks, and fall back to public validation when the fast path is unavailable.
- Recognize the implicit `xml` prefix, refresh XML parses between validation runs, and preserve ZIP-envelope findings alongside package loading failures.
- Restore type-aware lint startup by declaring `oxlint-tsgolint`; align Vitest and coverage 5.0.1 with Vite+ 1.0.0; give cold-start fixture tests appropriate timeout budgets.

### Known limitations

- Word compatibility for some uncovered package parts and unselected markup-compatibility branches still needs targeted Word A/B evidence; see [the inspection notes](docs/really-broken/README.md). Recorded Word observations were preserved; this release review did not perform new Word opens.

## [0.52.0] — 2026-10-06

Breaking under `0.x` conventions: the native `libxmljs2` addon is replaced by
`libxml2-wasm` and schema-load failures change severity — see `MIGRATION.md`.

### Added

- **`word-error-explanations` public module.** `WORD_ERROR_EXPLANATIONS` / `explainWordError` — plain-language, real-Word-probing-grounded explanations for the 19 issue codes the `word-valid` profile treats as Word-blocking. The CLI prints them as a `↳ …` line under each error when running `--profile word-valid`. Re-exported from the package root.
- **`ct-uncovered-part` OPC coverage check.** Parts matched by no `<Override PartName>` and no `<Default Extension>` in `[Content_Types].xml` are now flagged (previously only declarable XML roots and known media extensions were checked — extensionless media slipped through). `PartName`s are URI-decoded before comparison, and media-extension matching uses own-property checks only (`x.constructor` is not a media type).
- **`getElementsByTagNameAll`** — stack-based DFS walker replacing quadratic live-`NodeList` iteration for wildcard collection; `getElementsByTagNameNSAll` now uses the same O(N) traversal (one shared implementation, two entry points). New `tests/xml-helpers.test.ts` pins the contract, including the native root semantics (Document root ⇒ `documentElement` included; Element root ⇒ self excluded).
- **`MIGRATION.md`** with per-version breaking-change notes.
- CI runs a Node 24 + 26 matrix.

### Changed

- **XSD engine: `libxmljs2` → `libxml2-wasm`.** The native addon ships no Node 26 / ABI-147 prebuild and no longer compiles against Node's V8 headers; the wasm build has identical libxml2 semantics with no native compilation step. Compiled schema validators are cached per schema path (the cache stores in-flight load promises, so concurrent first loads compile exactly one validator; failed loads evict and retry; source documents are disposed even when compilation throws).
- **Schema-load failures are errors, not info.** Unloadable schemas now produce `severity: "error"`, `code: "xsd-schema-load-failed"` (was `info` / `xsd-schema-load-skipped`). A validator must not call a document valid against a schema it could not load. The known-noisy Dublin Core import in `opc-coreProperties.xsd` stays suppressed via `IGNORED_VALIDATION_ERRORS`.
- `BaseSchemaValidator.assertLibxmljsAvailable()` → **async** `assertXsdValidationAvailable()` (wasm initialises asynchronously); error prefix is now `XSD engine unavailable (libxml2-wasm)`.
- **Toolchain: `pnpm` replaces `bun`** (`packageManager: pnpm@12.8.2`); all dependencies updated to latest.
- **`src/index.ts` is hand-maintained.** barrelsby could not reproduce the curated named exports and silently destroyed them on regeneration; the `barrel` / `barrel:check` scripts are gone and the barrel is edited by hand.
- **Format settings moved** from `.oxfmtrc.json` into `vite.config.ts`'s `fmt` block so `vp check` and `pnpm run fmt` cannot desync; `fmt` / `fmt:fix` now go through `vp fmt`.

### Fixed

- **CI workflows actually run again.** `pnpm/action-setup` now executes before `actions/setup-node`'s `cache: pnpm` resolution (the old order failed every pnpm job with "Unable to locate executable file: pnpm" — CI had not run green on `main` since 2026-06-02). The validator-diff workflow resolves pnpm via `package_json_file` (its path-scoped checkouts leave no root `package.json`).
- Validator performance on large documents: del-first `collectDeletedRunText` with Set dedup across both Word namespaces, native-DOM `validateDeletions` (no `xpath` descendant queries), O(N) paragraph counting that skips text-box subtrees during traversal, single-pass redlining stripping.

### Security

- **Zip Slip** blocked at extraction (`resolveSafeZipEntry`: `path.relative` containment, prefix-slip-safe).
- **Path traversal in relationship targets** contained lexically and via realpath-based symlink-escape detection (`containRelationshipTarget`), for both absolute and relative targets.
- **Document IDs use `crypto.randomInt`** instead of predictable `Math.random`-derived identifiers.

### Removed

- `libxmljs2` (runtime + `allowBuilds` entry), `barrelsby` (dev), `.oxfmtrc.json`, the `barrel` / `barrel:check` scripts.

## [0.51.0] — 2026-05-30

### Changed

- Version bump to `0.51.0`.

## [0.5.0] — 2026-05-27

### Added

- **Content-descriptive fixture corpus.** Every undescriptive `.docx` test fixture is now named `<subject>.<comment-or-error>.docx`, derived deterministically from its content. New tooling under `scripts/`: `fixture-fingerprint.ts` (JSZip + `validate()` content fingerprint), `derive-fixture-name.ts` (pure name derivation with `error-first` and `content-first` descriptor modes), and `apply-fixture-names.ts` (rename/sort driver with content-hash dedup, collision disambiguation, and `--into-categories` / `--category` / `--descriptor` flags). All three are unit-tested.
- **`tests/fixtures/eigen/`** — 148 real-world Plate/SuperDoc specimens (deduped from 151) imported from the former unreferenced top-level `fixtures/eigen-extended/`, renamed by content (`content-first`: distinguishing feature, falling back to the validation error code). Provenance in `tests/fixtures/eigen/README.md`. The `fixtures-all-{strict,lenient}` suites now pin all of them (`tests/fixtures-all.manifest.json`: 416 → 564 fixtures).

### Changed

- **`scripts/update-manifest.ts` preserves `word` metadata across regens.** When the Word-probe JSONL is absent (no LibreOffice in CI), the per-fixture `word` outcome now falls back to the previous manifest's value instead of being downgraded to `"unknown"`. Exposed as the testable `resolveWordOutcome` helper.
- Renamed three undescriptive `tests/fixtures/` strays (a random-named root specimen, `Ouch.docx`, and a path-as-filename artifact) to content-derived names.

### Removed

- `Sample Document.repaired.docx` (a tracked, regenerable repair artifact) and the stray npm `package-lock.json` (this is a bun-only repo) from the repo root.
- The dead top-level `fixtures/` directory (nothing referenced it).

## [0.1.3] — 2026-05-04

### Added

- **TypeScript type declarations now ship in the package.** `vp pack` emits `dist/index.d.mts` (27.87 kB / 9.37 kB gzipped) alongside `dist/index.mjs`. `package.json` exports the types via the `exports."."` conditions (`types` listed first per TS bundler-resolution rules) and a top-level `types` fallback. Consumers of every public re-export from `src/index.ts` — `validate`, the validator classes, side helpers, and the `XsdValidationOutcome` / `ParagraphCounts` / `RedliningOptions` interfaces — now get full IntelliSense.
- **`"sideEffects": false`** in `package.json` so bundlers can tree-shake the package aggressively. The barrel only re-exports pure functions, classes, types, and namespace constants — no top-level side effects, so the claim is safe.
- **publint and attw build-time gates** wired into `vite.config.ts` (`pack.publint: true`, `pack.attw: true`). publint fails the build on `package.json` shape violations (`exports` map order, missing files); attw verifies type resolution under `node10` / `node16` / `bundler` modes. The `attw.profile: "esm-only"` field documents that the `cjs-resolves-to-esm` warning is intentional — this is an ESM-only package by design (CLAUDE.md "ES modules only; no CommonJS").
- **Test coverage in CI.** `@vitest/coverage-v8` devDep + `test.coverage` block in `vite.config.ts` (text + html + lcov reporters). `ci.yml` runs `vitest run --coverage` and uploads `coverage/lcov.info` via `codecov/codecov-action@v5` (token sourced from `secrets.CODECOV_TOKEN`).
- **CodeQL security scanning** (`.github/workflows/codeql.yml`). Uses `javascript-typescript` language with `build-mode: none`, scoped to first-party source via `paths-ignore` (skips fixtures, vendored snapshots, build output, docs). Runs on push/PR to `main` only — no schedule.
- README badges: CI status, CodeQL, Codecov, Snyk Known Vulnerabilities, npm version + downloads + types, bundle size (min + gzip), tree-shakeable, dependency count, license, node version, GitHub repo.
- `llms.txt` follows [llmstxt.org](https://llmstxt.org/) format — title + summary + Docs / API / Tests / Optional sections with deep links into every public module.
- Public-surface exports for `XsdValidationOutcome`, `ParagraphCounts`, and `RedliningOptions` (previously declared `interface` without `export`, so consumers couldn't reference them by name).
- `knip.json` with `ignoreDependencies` for plugins knip can't statically resolve through oxlint's `jsPlugins[]` config (`barrelsby`, `eslint-plugin-functional`, `eslint-plugin-jsdoc`, `eslint-plugin-prefer-arrow`).
- `coverage/` added to `.gitignore`.
- `bun run release:publish` — single-step `git push --follow-tags && npm publish --access public`. Uses local npm credentials (`npm login` / `NPM_TOKEN`); the `publish.yml` Trusted-Publishing OIDC flow remains as a backup if you'd rather publish from CI. Named `release:publish` (not `publish`) to avoid the recursive lifecycle trigger when `npm publish` runs the `publish` script.
- Maintainer line in README pointing at [jandira.tech](https://www.jandira.tech) — Jandira Technologies + Cicero context.

### Changed

- `package.json` `description`: trailing typo fix (`"ESM. for the neurotic developer"` → `"ESM, for the neurotic developer."`).
- README header rewritten with badge row; summary tightened.
- typedoc no longer surfaces stale TS errors — unused imports (`mergeResults` in `validators/base.ts`, `serializeXml` in `validators/redlining.ts`) cleared.

### Internal

- cspell dictionary: added `barrelsby`, `llms`.
- `defaultSchemasDir()` in `validators/base.ts` falls back to both layouts (`../schemas` for the source tree, `./schemas` for the published `dist/` bundle) so the bundled package resolves XSDs against the copy emitted by the `copySchemasPlugin` in `vite.config.ts`.

## [0.1.2] — 2026-05-04

### Security

- **LD_PRELOAD shim removed from the published bundle.** The
  `acceptChanges()` LibreOffice flow uses a runtime-compiled C shim that
  `LD_PRELOAD`s into `soffice` to work around `AF_UNIX` restrictions in
  sandboxed VMs. That pattern legitimately matches supply-chain malware
  heuristics in automated scanners (Socket.dev, etc.) — and they're
  correct to flag _unknown_ packages doing this. The shim source remains
  in `src/scripts/office/soffice.ts` for developers who run the CLI
  directly, but `src/index.ts` no longer re-exports `soffice.ts` or
  `accept-changes.ts`, so the bundler tree-shakes both files (and
  `SHIM_SOURCE`, `LD_PRELOAD`, `RTLD_NEXT`, `dlsym`, etc.) out of
  `dist/index.mjs`. `grep -E "LD_PRELOAD|SHIM_SOURCE|RTLD_NEXT|dlsym"
dist/index.mjs` returns zero matches.
- Added `SECURITY.md` with the disclosure policy and a detailed
  explanation of the shim pattern, where it lives, and how to verify the
  published bundle is clean.

### Added

- **`bun run release` checklist wrapper** (`scripts/release.ts`) that
  reads `.release-checklist.md`, prompts `[y/N]` for each unchecked
  item, and aborts on the first "no" before forwarding to `bumpp`. The
  checklist itself lives in `.release-checklist.md` and isn't hard-coded
  — projects edit the file to match their actual workflow. `--yes` skips
  prompts for CI.
- `CHANGELOG.md` (this file).
- `llms.txt` package summary for LLM-friendly tooling.
- CodeQL workflow (`.github/workflows/codeql.yml`).
- Codecov coverage upload step in `ci.yml`.
- README badges for CI / CodeQL / Codecov / npm / bundle size / types.

### Changed

- README "Programmatic use" section updated with a clearly-marked "Not
  in the package surface: LibreOffice helpers" subsection pointing at
  SECURITY.md and the source-checkout path for callers that need
  `acceptChanges()` / `runSoffice()`.
- `llms.txt` synced with the post-shim-removal API surface.

### Internal

- `package.json` `barrel` script now passes `--exclude` patterns for
  `scripts/office/soffice\.ts$` and `scripts/accept-changes\.ts$` so
  regenerating the barrel preserves the surface trim.

## [0.1.1] — earlier

Initial publishable release. See `git log --oneline` for the full set of
commits — the changelog starts here.

[0.52.0]: https://github.com/jandira-tech/docx-validate/compare/v0.51.0...v0.52.0
[0.51.0]: https://github.com/jandira-tech/docx-validate/compare/v0.5.0...v0.51.0
[0.5.0]: https://github.com/jandira-tech/docx-validate/compare/v0.1.3...v0.5.0
[0.1.3]: https://github.com/jandira-tech/docx-validate/compare/v0.1.2...v0.1.3
[0.1.2]: https://github.com/jandira-tech/docx-validate/compare/v0.1.1...v0.1.2
[0.1.1]: https://github.com/jandira-tech/docx-validate/releases/tag/v0.1.1
[0.53.0-rc.1]: https://github.com/jandira-tech/docx-validate/compare/v0.52.0...v0.53.0-rc.1
[Unreleased]: https://github.com/jandira-tech/docx-validate/compare/v0.53.0-rc.1...HEAD
