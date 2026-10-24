# Real-Word probes, 2026-10-24

Observations behind the word-valid policy for #388, #392, #396, #397, #351, #398, #402, #408, #409, #410 and #411, the scope of `word-content-type-invalid` (#401), and the defects session 3 found (#380, #414, #431–#437).

- Word: Microsoft Word for Mac 16.115 (build 16.115.26100424), platform `darwin`.
- Transport, sessions 1 and 2: `verifyInWord` (`src/lib/word-verifier.ts`). It opens each specimen by path, reads Word's window and dialog state, and never saves.
- Transport, session 3: `neurotic_docx_bench/scripts/word_pdf.py` in its default batch mode (`--label wfv --no-check-preset`). It stages each file inside Word's container as `NNNNN__wfv__<file>`, opens it, and exports a PDF when Word reads it. When Word shows an alert, it brings Word forward, answers a repair prompt No and then the OK that follows, and logs the alert's words. A file's outcome comes only from an alert that names its staged file, never from timing (`wordpdf_outcomes.py`, reproduced at the end). The specimens' SHA-256 were checked against the files handed to the driver after each run, and no specimen was rebuilt between build and run.
- Outcomes: `clean-open` means Word opened the document with no dialog. `repair-warning` means Word showed "Word found unreadable content … Do you want to recover the contents of this document?". `open-error` means Word showed "Word experienced an error trying to open the file." In session 3, `inconclusive` means Word loaded no document in two passes and no alert named the file. Word stayed alive throughout, and `~/Library/Logs/DiagnosticReports` holds no Word report from any session. In session 3 the driver itself restarted Word after runs of empty loads, as designed.
- Specimens: built in code from `MINIMAL_PARTS` (`tests/support/minimal-docx.ts`). The `398-dup-theme-*` specimens are built from `tests/fixtures/broken/package.word-duplicate-theme.docx`, which was only read. The probe script that builds, validates and opens them is reproduced at the end of this file. No fixture or manifest was changed.
- Validator: the `word-valid` and `strict` columns record `validate()` **at probe time**, before the policy changes the probes motivated. They are not current verdicts. Session 1 (02:18Z–02:26Z) ran against `wfv/46-plan-audit` (199c2d5b); session 2 (03:10Z–03:19Z) ran against `wfv/51-directory-staging` (ad072db8). Session 3's columns are `validate()` at `wfv/57-central-directory-names` (ab7f311d), run on the exact files Word was given.

## What the observations decide

- **#388: package relationships.** A zero-byte, truncated or undeclared-prefix `_rels/.rels` makes Word offer recovery, and the readable control opens clean. Word-valid passed all three, so `xml-syntax` and `xml-undeclared-prefix` on `_rels/.rels` are now Word-blocking.
- **#392: `WORD/` folder.** A package with every Word part under `WORD/` opens clean. With its main part malformed, Word refuses it exactly as it refuses the lower-case control. That matches the current policy, which matches `word/` by OPC part name.
- **#396: content-types stream spelling.** A stream stored as `[CONTENT_TYPES].XML` opens clean. Malformed, it makes Word offer recovery, as word-valid predicts. That matches the any-case lookup.
- **#397: unknown declared encoding.** `word/document.xml` declaring `X-UNKNOWN-ENCODING` makes Word refuse the file, and the `windows-1252` control opens clean. Word-valid passed it, so this was a false negative; `xml-encoding` now follows the `xml-syntax` path policy. Session 2 confirms the inferred scope:
    - the same declaration in `word/styles.xml`, in the content-types stream and in `_rels/.rels` makes Word offer recovery;
    - a `word/document.xml` declared UTF-8 holding a lone Latin-1 byte makes Word refuse it;
    - word-valid fails all four;
    - in `docProps/core.xml` the declaration also makes Word offer recovery, but word-valid passes it (#409).
- **#351: equivalent part names.** Every duplicate probed makes Word offer recovery, even when both spellings hold identical bytes: case variants, an unreserved escape (`word/%73tyles.xml`), a backslash (`word\document.xml`), and two content-types streams. Word-valid passed the escape and backslash duplicates. A single backslash name is read as `/`: the main part stored only as `word\document.xml` opens clean (#374).
- **#408: escaped item names (OPC [M1.8]).** Any stored name with a percent-escaped unreserved character makes Word offer recovery:
    - the styles part stored only as `word/%73tyles.xml`;
    - the main part stored only as `word/%64ocument.xml`;
    - the sidecar stored as `document.xml.%72els`;
    - an unreferenced extra `word/%78tra.xml`, which settles the cause, since nothing points at it;
    - the styles control (its part missing, its relationship kept) opens clean, so a missing target is not the cause;
    - word-valid passed the first two and the extra item.
- **#398 and #402: `.RELS` sidecar.** A `.RELS`-spelled document sidecar opens clean. With two theme relationships, Word offers recovery for both the `.rels` control and the `.RELS` variant, as word-valid now predicts for both. An external hyperlink resolved through a `.RELS` sidecar opens clean, as the `.rels` control does, but word-valid fails it. That is the #402 false positive: the sidecar lookups must fold case. Through a `.%72els` sidecar, Word offers recovery; that is #408.
- **#409: core properties.** A truncated `docProps/core.xml`, or one declaring an unknown encoding, makes Word offer recovery. Word-valid passes both, because docProps findings are downgraded. The `docProps/app.xml` control opens clean. Session 3 settles the rule (below).
- **#401: `word-content-type-invalid`.** An `invalid-*` type on an existing part makes Word refuse the file or offer recovery: styles, the main part and core properties. The same type on an `Override` for a missing part, or on a `Default` that covers no part, opens clean. The check over-reported; it now reports only declarations that apply to a stored part.

Session 3 adds:

- **#409: properties parts are read through their relationships.** Truncated, undeclared-prefix and unknown-encoding `docProps/core.xml`, `app.xml` and `custom.xml`, each named by its package relationship, make Word offer recovery. The same defects in unreferenced `core.xml`, `app.xml` and `custom.xml` open clean. A truncated core part stored as `props/meta.xml` and named by the core-properties relationship makes Word offer recovery, and its well-formed control opens clean. So Word reads what the relationship names, wherever it is stored. Word-valid passes every one of these (lane 409 implements the rule).
- **#410: real encodings libxml2 cannot decode.** `word/document.xml` declaring `macintosh`, `x-mac-cyrillic`, `IBM866` or `UTF-7` makes Word refuse the file. The `windows-1252` control with non-ASCII text opens clean. Word-valid already fails all four, so the policy stands; the README now says it covers these labels.
- **#411: identical stored names and dot segments.** Two entries stored under the identical name `word/document.xml` make Word offer recovery, whichever entry is malformed, and so do two `..` item names. At ab7f311d word-valid fails all five.
- **#408: every unreserved escape probed.** `%41` (upper-case letter), `%30` (digit), `%2D` and `%2d` (hyphen), `%2E`, `%5F` and `%7E` in an unreferenced extra item all make Word offer recovery. The `%20` control (a reserved character) opens clean. An extra `[Content_Types].%78ml` item beside the real stream makes Word offer recovery too, so the check is right to report a stream spelling (review X F1).
- **#436: escaped `/` and `\` (OPC [M1.7]).** An unreferenced `word%2Fxtra.xml` makes Word offer recovery, and `word%5Cxtra.xml` makes it refuse the file. Word-valid passes both.
- **#414: `Override` PartNames that are not part names.** The session 3b styles specimens are confounded (#432, below). Session 3c adds an `Override` with `ContentType="invalid-x"` for a part that does not exist. With the plain `PartName` `/word/nothere.xml`, it opens clean. With `..`, `#fragment` or a backslash, Word offers recovery, and with `?query` it refuses the file. Word-valid passes all four.
- **#432: a styles part typed by the generic Default.** Removing the styles `Override`, so that `word/styles.xml` falls to `<Default Extension="xml" ContentType="application/xml"/>`, makes Word offer recovery. Nothing reports it.
- **#431: duplicate content-type declarations.** Two `Default`s for one extension (identical, or differing in case) and two `Override`s for one part (identical, conflicting, or differing in case) all make Word offer recovery. Word-valid and strict pass all five.
- **#433 and #434: package relationships structure.** A foreign element anywhere in `_rels/.rels`, whether beside the relationships or wrapping the main one, makes Word offer recovery, and so does a root other than `Relationships`. Word-valid passes all three; strict fails them through the XSD walk only.
- **#435: no content-types stream.** Without `[Content_Types].xml`, Word offers recovery. Word-valid reports `ct-missing` as a warning only.
- **#380: no main part.** An empty ZIP and a package without `word/document.xml` make Word offer recovery. Word-valid passes both. A package whose main relationship is removed is inconclusive.
- **#437: an unreferenced malformed part under `word/`.** A truncated, unreferenced `word/xtra.xml` opens clean, but word-valid fails it. This is a false positive of the `word/` folder rule; reachability, as for the properties parts, would fit both observations.

## Not observed

Session 1. Another session began driving Word by 02:28Z: its first log line is at 02:28:14Z, and its map file was written at 02:27:55Z. Its preflight closes every open document, so rows observed from 02:26Z on were not treated as evidence. Session 2 observed all of them again (below).

Session 2. Every record now carries a guard. Before and after each open, the probe lists the processes driving Word that it did not start itself: an `osascript` executable, or the other session's `word_pdf`/`word_redline` scripts. It also notes whether that session's log grew. A record with any such process goes to an inconclusive file instead of the table.

- The first run, at 03:08Z, matched on command-line text, not on executables: its own shell and two review agents whose prompts forbid osascript. All ten of its records were set aside, and the guard was fixed. The 03:10Z run repeated the same ten specimens with a clean guard and identical outcomes.
- At 03:19:33Z the other session started a Word batch mid-run, and the last two records were set aside.

Session 3 observed every specimen listed for it after session 2, except as noted here.

- The first session 3 batch, at 04:38Z, ran in the default mode but wedged at item 8: Word was slow to refuse the encoding specimens, and the driver's replays queued opens of staged files it had already deleted. Its rows were not used. A serial re-run with `--no-one-osascript` was also set aside, because that mode never brings Word forward, so Word's alerts for files in the background were never drawn. Sessions 3b and 3c below are complete default-mode runs.
- `rels-unknown-sibling-control` was inconclusive in session 3b and was observed in session 3c.
- Sessions 3b and 3c were started from a background shell. macOS therefore asked for AppleScript automation authorization for the driver's helper processes, and a person granted each prompt by hand. Every alert was read and attributed, so no outcome depends on this, but later sessions run in the foreground, as the driver's documentation requires.
- No other Word driver ran during sessions 3b (05:02Z–05:18Z) or 3c (05:19Z–05:27Z). Every driver takes the same flock, and the other session's next Word batch began at 05:29Z.

To be observed in the next Word session:

- `main-relationship-missing` (#380), inconclusive in session 3c.
- The #414 `PartName`s with a well-formed content type, such as `application/xml`, instead of `invalid-x`.
- An escape in a folder segment, such as `%77ord/document.xml` (#408).
- Other parts typed by the generic Default: `settings`, `fontTable`, `numbering`, `theme`, `webSettings`, and the main part (#432).
- Foreign elements in a part-level relationships part, such as `word/_rels/document.xml.rels` (#433).
- An unreferenced malformed part outside `word/`, such as `customXml/item9.xml` (#437).

## Observations

### Session 1: 02:18Z–02:26Z, validator at 199c2d5b

| Specimen                                 | Issue   | SHA-256 (first 16) | Word           | word-valid at probe time                                                                                                    | strict at probe time |
| ---------------------------------------- | ------- | ------------------ | -------------- | --------------------------------------------------------------------------------------------------------------------------- | -------------------- |
| `control-minimal`                        | control | `e3fd9a9607c643b1` | clean-open     | valid                                                                                                                       | valid                |
| `388-rels-zero-byte`                     | #388    | `518a110444455811` | repair-warning | valid                                                                                                                       | invalid              |
| `388-rels-truncated`                     | #388    | `a85ad66282671910` | repair-warning | valid                                                                                                                       | invalid              |
| `388-rels-undeclared-prefix`             | #388    | `19383610c2a9f89a` | repair-warning | valid                                                                                                                       | invalid              |
| `392-WORD-folder-clean`                  | #392    | `36d4625db191d3e9` | clean-open     | valid                                                                                                                       | valid                |
| `392-WORD-folder-malformed-main`         | #392    | `b5c3d2178d4696be` | open-error     | invalid (`xml-undeclared-prefix@WORD/document.xml`, `xml-syntax@WORD/document.xml`)                                         | invalid              |
| `392-word-folder-malformed-main-control` | #392    | `95d4f0bfb82a2c4a` | open-error     | invalid (`xml-undeclared-prefix@word/document.xml`, `xml-syntax@word/document.xml`)                                         | invalid              |
| `396-CT-upper-clean`                     | #396    | `744bbb14ad7f0f3e` | clean-open     | valid                                                                                                                       | valid                |
| `396-CT-upper-invalid-type`              | #396    | `ec74a991d4b8b9dc` | clean-open     | invalid (`word-content-type-invalid@[CONTENT_TYPES].XML`)                                                                   | invalid              |
| `396-CT-exact-invalid-type-control`      | #396    | `9857e3aaebaa53d7` | clean-open     | invalid (`word-content-type-invalid@[Content_Types].xml`)                                                                   | invalid              |
| `396-CT-upper-malformed`                 | #396    | `949e38e644ebd7db` | repair-warning | invalid (`xml-syntax@[CONTENT_TYPES].XML`)                                                                                  | invalid              |
| `397-encoding-unknown`                   | #397    | `a45331da8b2129ea` | open-error     | valid                                                                                                                       | invalid              |
| `397-encoding-windows-1252-control`      | #397    | `0e8165d11d1898b5` | clean-open     | valid                                                                                                                       | valid                |
| `351-dup-case-identical`                 | #351    | `48a8b62b9c83535d` | repair-warning | invalid (`package-extract-failed@-`)                                                                                        | invalid              |
| `351-dup-case-lower-good-upper-bad`      | #351    | `2d58ed5a3241a634` | repair-warning | invalid (`package-extract-failed@-`)                                                                                        | invalid              |
| `351-dup-case-upper-bad-first`           | #351    | `8ef60be49ff4ac55` | repair-warning | invalid (`package-extract-failed@-`)                                                                                        | invalid              |
| `351-dup-stream-good-then-bad`           | #351    | `970eca474f9eba1c` | repair-warning | invalid (`package-extract-failed@-`)                                                                                        | invalid              |
| `351-dup-stream-bad-then-good`           | #351    | `b503d9e06055091e` | repair-warning | invalid (`package-extract-failed@-`)                                                                                        | invalid              |
| `398-RELS-sidecar-clean`                 | #398    | `ac61b47790dc028b` | clean-open     | valid                                                                                                                       | valid                |
| `398-dup-theme-control`                  | #398    | `52b18dea69dafe56` | repair-warning | invalid (`word-duplicate-theme-relationship@word/_rels/document.xml.rels`)                                                  | invalid              |
| `398-dup-theme-RELS`                     | #398    | `f1a5a570c9effd78` | repair-warning | invalid (`word-duplicate-theme-relationship@word/_rels/document.xml.RELS`, `word-hyperlink-relationship@word/document.xml`) | invalid              |
| `cti-override-existing-styles`           | #401    | `dde1b233d9672e0f` | repair-warning | invalid (`word-content-type-invalid@[Content_Types].xml`)                                                                   | invalid              |
| `cti-override-existing-main`             | #401    | `bded55854c3edf59` | repair-warning | invalid (`word-content-type-invalid@[Content_Types].xml`)                                                                   | invalid              |
| `cti-override-existing-core`             | #401    | `f4081445c08647f1` | open-error     | invalid (`word-content-type-invalid@[Content_Types].xml`)                                                                   | invalid              |
| `cti-override-existing-core-control`     | #401    | `86df51ceafbcb5ba` | clean-open     | valid                                                                                                                       | valid                |
| `cti-default-unused-extension`           | #401    | `d19b131b7b46dbba` | clean-open     | invalid (`word-content-type-invalid@[Content_Types].xml`)                                                                   | invalid              |
| `351-dup-escape-identical`               | #351    | `f5cd33350d5e730e` | repair-warning | valid                                                                                                                       | invalid              |
| `351-dup-backslash-identical`            | #351    | `348701aaaa331f41` | repair-warning | valid                                                                                                                       | valid                |

### Session 2: 03:10Z–03:19Z, validator at ad072db8

| Specimen                             | Issue | SHA-256 (first 16) | Word           | word-valid at probe time                                                                            | strict at probe time |
| ------------------------------------ | ----- | ------------------ | -------------- | --------------------------------------------------------------------------------------------------- | -------------------- |
| `351-escape-only-styles`             | #351  | `ca196e7ef4a292f2` | repair-warning | valid                                                                                               | valid                |
| `351-backslash-only-main`            | #351  | `869f616f1ba1640e` | clean-open     | valid                                                                                               | valid                |
| `398-hyperlink-rels-control`         | #398  | `21f622f77bf187c2` | clean-open     | valid                                                                                               | valid                |
| `398-hyperlink-%72els`               | #402  | `48a74a89fdd0177e` | repair-warning | invalid (`rels-missing-sidecar@word/document.xml`, `word-hyperlink-relationship@word/document.xml`) | invalid              |
| `398-hyperlink-RELS`                 | #398  | `440b780d13f48f40` | clean-open     | invalid (`word-hyperlink-relationship@word/document.xml`)                                           | invalid              |
| `397-encoding-unknown-styles`        | #397  | `d2c61a613863872e` | repair-warning | invalid (`xml-encoding@word/styles.xml`)                                                            | invalid              |
| `397-encoding-unknown-content-types` | #397  | `a2be2959ad591d9a` | repair-warning | invalid (`xml-encoding@[Content_Types].xml`)                                                        | invalid              |
| `397-encoding-unknown-package-rels`  | #397  | `c8a06454f5c52111` | repair-warning | invalid (`xml-encoding@_rels/.rels`)                                                                | invalid              |
| `397-encoding-unknown-core`          | #397  | `bde4363f2fed0946` | repair-warning | valid                                                                                               | invalid              |
| `397-invalid-utf8-bytes`             | #397  | `87b043160608a0aa` | open-error     | invalid (`xml-encoding@word/document.xml`)                                                          | invalid              |
| `esc-unreferenced-extra-item`        | #408  | `02392f4064145df1` | repair-warning | valid                                                                                               | invalid              |
| `esc-styles-missing-control`         | #408  | `79dddb1d4ffcd1ca` | clean-open     | valid                                                                                               | invalid              |
| `esc-only-main`                      | #408  | `845abca15ef137fe` | repair-warning | valid                                                                                               | valid                |
| `docprops-core-truncated`            | #409  | `0f729853110d309f` | repair-warning | valid                                                                                               | invalid              |
| `docprops-app-control`               | #409  | `727b58abd042452e` | clean-open     | valid                                                                                               | valid                |

Full hashes, dialog text and observation times are in the JSONL record of each probe, reproduced below.

<details><summary>Raw records</summary>

```jsonl
{"name": "control-minimal", "sha256": "e3fd9a9607c643b1615001a916501be367de53faeb4fd925ea38a4e8c034555b", "observedAt": "2026-10-24T02:18:56.522Z", "outcome": "clean-open", "details": "processAlive=true; documents=1; names=docx-verify-e3fd9a9607c6-22ee5d2f.docx; buttons=missing value,missing value,missing value"}
{"name": "388-rels-zero-byte", "sha256": "518a110444455811628385302e5c5f69036f8543327108eda929ac46bdccdcc8", "observedAt": "2026-10-24T02:19:10.007Z", "outcome": "repair-warning", "details": "processAlive=true; documents=0; dialog=Microsoft Word Word found unreadable content in docx-verify-518a11044445-3a7272c0.docx. Do you want to recover the contents of this document? If you trust the s; buttons=Yes,No"}
{"name": "388-rels-truncated", "sha256": "a85ad662826719107f763185b905b4609e0e03b65393a0f1fc3ad24d955099c9", "observedAt": "2026-10-24T02:19:13.170Z", "outcome": "repair-warning", "details": "processAlive=true; documents=0; dialog=Microsoft Word Word found unreadable content in docx-verify-a85ad6628267-6e2605ed.docx. Do you want to recover the contents of this document? If you trust the s; buttons=Yes,No"}
{"name": "388-rels-undeclared-prefix", "sha256": "19383610c2a9f89af29cd2ce549f38d09c2a3cec4510ab6111aeb97492230cbb", "observedAt": "2026-10-24T02:19:16.391Z", "outcome": "repair-warning", "details": "processAlive=true; documents=0; dialog=Microsoft Word Word found unreadable content in docx-verify-19383610c2a9-39293b8c.docx. Do you want to recover the contents of this document? If you trust the s; buttons=Yes,No"}
{"name": "392-WORD-folder-clean", "sha256": "36d4625db191d3e9d58ea1d2f108c3829bb1120e415137574bf105a6d642dfb0", "observedAt": "2026-10-24T02:19:19.571Z", "outcome": "clean-open", "details": "processAlive=true; documents=1; names=docx-verify-36d4625db191-17e063f1.docx; buttons=missing value,missing value,missing value"}
{"name": "392-WORD-folder-malformed-main", "sha256": "b5c3d2178d4696be642c02a63acc76ddbecbb89e5949fff78d77d6e34f077b89", "observedAt": "2026-10-24T02:19:25.387Z", "outcome": "open-error", "details": "processAlive=true; documents=0; dialog=Microsoft Word Word experienced an error trying to open the file.; buttons=OK"}
{"name": "392-word-folder-malformed-main-control", "sha256": "95d4f0bfb82a2c4ad256fd0692fef37730d9e00601bc8df2e6c50bb758913072", "observedAt": "2026-10-24T02:19:27.963Z", "outcome": "open-error", "details": "processAlive=true; documents=0; dialog=Microsoft Word Word experienced an error trying to open the file.; buttons=OK"}
{"name": "396-CT-upper-clean", "sha256": "744bbb14ad7f0f3ea38903283a49203448d176fb81e7a42fb4f3e5923200b2d2", "observedAt": "2026-10-24T02:19:30.409Z", "outcome": "clean-open", "details": "processAlive=true; documents=1; names=docx-verify-744bbb14ad7f-533bc5b2.docx; buttons=missing value,missing value,missing value"}
{"name": "396-CT-upper-invalid-type", "sha256": "ec74a991d4b8b9dcd0e18603f3fe10130f4e05058985046034fd3757fb18abc6", "observedAt": "2026-10-24T02:19:36.354Z", "outcome": "clean-open", "details": "processAlive=true; documents=1; names=docx-verify-ec74a991d4b8-68a1da7f.docx; buttons=missing value,missing value,missing value"}
{"name": "396-CT-exact-invalid-type-control", "sha256": "9857e3aaebaa53d7938dd1a048d43965b628faa6a40604a0319dc4baf1892840", "observedAt": "2026-10-24T02:19:42.537Z", "outcome": "clean-open", "details": "processAlive=true; documents=1; names=docx-verify-9857e3aaebaa-48b7cbb6.docx; buttons=missing value,missing value,missing value"}
{"name": "396-CT-upper-malformed", "sha256": "949e38e644ebd7dbae84d399c5682ad07adea9feaf661f4f51e206782d3e4705", "observedAt": "2026-10-24T02:19:48.322Z", "outcome": "repair-warning", "details": "processAlive=true; documents=0; dialog=Microsoft Word Word found unreadable content in docx-verify-949e38e644eb-491711fa.docx. Do you want to recover the contents of this document? If you trust the s; buttons=Yes,No"}
{"name": "397-encoding-unknown", "sha256": "a45331da8b2129ea843e382a493a38605597921880738ba558e9f12fce5ce0aa", "observedAt": "2026-10-24T02:19:51.391Z", "outcome": "open-error", "details": "processAlive=true; documents=0; dialog=Microsoft Word Word experienced an error trying to open the file.; buttons=OK"}
{"name": "397-encoding-windows-1252-control", "sha256": "0e8165d11d1898b5d286c31c832235942657751da8b9f4444aa89193c056fb8b", "observedAt": "2026-10-24T02:19:53.893Z", "outcome": "clean-open", "details": "processAlive=true; documents=1; names=docx-verify-0e8165d11d18-5f5e06b4.docx; buttons=missing value,missing value,missing value"}
{"name": "351-dup-case-identical", "sha256": "48a8b62b9c83535d2d93e2ef9b0bcbc7f40f703f9bd63c9dbabb534b21e343e9", "observedAt": "2026-10-24T02:20:00.197Z", "outcome": "repair-warning", "details": "processAlive=true; documents=0; dialog=Microsoft Word Word found unreadable content in docx-verify-48a8b62b9c83-3ba31e8f.docx. Do you want to recover the contents of this document? If you trust the s; buttons=Yes,No"}
{"name": "351-dup-case-lower-good-upper-bad", "sha256": "2d58ed5a3241a6341c59340a0723c64c00508566c3a7bc6b56165f26383144ca", "observedAt": "2026-10-24T02:20:04.079Z", "outcome": "repair-warning", "details": "processAlive=true; documents=0; dialog=Microsoft Word Word found unreadable content in docx-verify-2d58ed5a3241-15e1301a.docx. Do you want to recover the contents of this document? If you trust the s; buttons=Yes,No"}
{"name": "351-dup-case-upper-bad-first", "sha256": "8ef60be49ff4ac5507c1be56e26e9312f2c3965103b4a933084f4a419526a387", "observedAt": "2026-10-24T02:20:07.779Z", "outcome": "repair-warning", "details": "processAlive=true; documents=0; dialog=Microsoft Word Word found unreadable content in docx-verify-8ef60be49ff4-1ee9161e.docx. Do you want to recover the contents of this document? If you trust the s; buttons=Yes,No"}
{"name": "351-dup-stream-good-then-bad", "sha256": "970eca474f9eba1c30fc10fb4343195616cbfc0863c8e7247a3752ea7a89b518", "observedAt": "2026-10-24T02:20:11.571Z", "outcome": "repair-warning", "details": "processAlive=true; documents=0; dialog=Microsoft Word Word found unreadable content in docx-verify-970eca474f9e-69aa90bc.docx. Do you want to recover the contents of this document? If you trust the s; buttons=Yes,No"}
{"name": "351-dup-stream-bad-then-good", "sha256": "b503d9e06055091e24edec6fe109722fbc26b37eca71911b727fa5199eb0b294", "observedAt": "2026-10-24T02:20:15.230Z", "outcome": "repair-warning", "details": "processAlive=true; documents=0; dialog=Microsoft Word Word found unreadable content in docx-verify-b503d9e06055-4b6e42d7.docx. Do you want to recover the contents of this document? If you trust the s; buttons=Yes,No"}
{"name": "398-RELS-sidecar-clean", "sha256": "ac61b47790dc028bfdd35606ee0ed593c843c98df57c107a2ca626d201743928", "observedAt": "2026-10-24T02:20:18.973Z", "outcome": "clean-open", "details": "processAlive=true; documents=1; names=docx-verify-ac61b47790dc-a167f9e.docx; buttons=missing value,missing value,missing value"}
{"name": "398-dup-theme-control", "sha256": "52b18dea69dafe56f1242778155d03e1dc419395f21dac99436ff06fc3305758", "observedAt": "2026-10-24T02:20:24.995Z", "outcome": "repair-warning", "details": "processAlive=true; documents=0; dialog=Microsoft Word Word found unreadable content in docx-verify-52b18dea69da-315ee282.docx. Do you want to recover the contents of this document? If you trust the s; buttons=Yes,No"}
{"name": "398-dup-theme-RELS", "sha256": "f1a5a570c9effd789dd70cf7264b12bed75d4b930325b156f74a625d34617dd0", "observedAt": "2026-10-24T02:20:28.959Z", "outcome": "repair-warning", "details": "processAlive=true; documents=0; dialog=Microsoft Word Word found unreadable content in docx-verify-f1a5a570c9ef-d43130e.docx. Do you want to recover the contents of this document? If you trust the so; buttons=Yes,No"}
{"name": "cti-override-existing-styles", "sha256": "dde1b233d9672e0f4038cd6439d2109cc2daf31e737a634c628681624a3d9366", "observedAt": "2026-10-24T02:25:27.343Z", "outcome": "repair-warning", "details": "processAlive=true; documents=0; dialog=Microsoft Word Word found unreadable content in docx-verify-dde1b233d967-152ec719.docx. Do you want to recover the contents of this document? If you trust the s; buttons=Yes,No"}
{"name": "cti-override-existing-main", "sha256": "bded55854c3edf59deb21a1e9942d9a901c7585d2101bae26e5673cf04d6d465", "observedAt": "2026-10-24T02:25:32.196Z", "outcome": "repair-warning", "details": "processAlive=true; documents=0; dialog=Microsoft Word Word found unreadable content in docx-verify-bded55854c3e-2f5ded2d.docx. Do you want to recover the contents of this document? If you trust the s; buttons=Yes,No"}
{"name": "cti-override-existing-core", "sha256": "f4081445c08647f1ffa572eb8b305e8bbf656df9505728bc795851c8a0337426", "observedAt": "2026-10-24T02:25:36.391Z", "outcome": "open-error", "details": "processAlive=true; documents=0; dialog=Microsoft Word Word experienced an error trying to open the file.; buttons=OK"}
{"name": "cti-override-existing-core-control", "sha256": "86df51ceafbcb5ba7dbbe52523c8dfd47749e105f22efe99c62e0e4daf152b5a", "observedAt": "2026-10-24T02:25:39.665Z", "outcome": "clean-open", "details": "processAlive=true; documents=1; names=docx-verify-86df51ceafbc-1154a840.docx; buttons=missing value,missing value,missing value"}
{"name": "cti-default-unused-extension", "sha256": "d19b131b7b46dbba48f21f4997752bfcbc97722bfde8b714b4f8f2e9d7cf4374", "observedAt": "2026-10-24T02:25:46.833Z", "outcome": "clean-open", "details": "processAlive=true; documents=1; names=docx-verify-d19b131b7b46-75ea9926.docx; buttons=missing value,missing value,missing value"}
{"name": "351-dup-escape-identical", "sha256": "f5cd33350d5e730e8d7729493407a69834a98a8cc5ea78d258700a462ec1234c", "observedAt": "2026-10-24T02:25:53.635Z", "outcome": "repair-warning", "details": "processAlive=true; documents=0; dialog=Microsoft Word Word found unreadable content in docx-verify-f5cd33350d5e-1ee1b408.docx. Do you want to recover the contents of this document? If you trust the s; buttons=Yes,No"}
{"name": "351-dup-backslash-identical", "sha256": "348701aaaa331f414752bb12024e1227f17c0ae9695c15b0785f91b935d1881b", "observedAt": "2026-10-24T02:25:58.711Z", "outcome": "repair-warning", "details": "processAlive=true; documents=0; dialog=Microsoft Word Word found unreadable content in docx-verify-348701aaaa33-46da418.docx. Do you want to recover the contents of this document? If you trust the so; buttons=Yes,No"}
{"name": "351-escape-only-styles", "sha256": "ca196e7ef4a292f215eec905800c44a95308695b4c846da3c8de202bc42961d3", "observedAt": "2026-10-24T03:10:04.556Z", "outcome": "repair-warning", "details": "processAlive=true; documents=0; dialog=Microsoft Word Word found unreadable content in docx-verify-ca196e7ef4a2-340c7c4b.docx. Do you want to recover the contents of this document? If you trust the s; buttons=Yes,No"}
{"name": "351-backslash-only-main", "sha256": "869f616f1ba1640e7d87b83017fcfec7e25afc8fc8f584031e1c38bc6c3f6d38", "observedAt": "2026-10-24T03:10:12.763Z", "outcome": "clean-open", "details": "processAlive=true; documents=1; names=docx-verify-869f616f1ba1-74030a55.docx; buttons=missing value,missing value,missing value"}
{"name": "398-hyperlink-rels-control", "sha256": "21f622f77bf187c2ed548011da2c47aa1112bdd346fd4c65952818a1543b078e", "observedAt": "2026-10-24T03:10:20.131Z", "outcome": "clean-open", "details": "processAlive=true; documents=1; names=docx-verify-21f622f77bf1-60b154d8.docx; buttons=missing value,missing value,missing value"}
{"name": "398-hyperlink-%72els", "sha256": "48a74a89fdd0177e2a178168e47928058da5f14e6c36c431e417ff6be40ecb02", "observedAt": "2026-10-24T03:10:27.011Z", "outcome": "repair-warning", "details": "processAlive=true; documents=0; dialog=Microsoft Word Word found unreadable content in docx-verify-48a74a89fdd0-29d9260f.docx. Do you want to recover the contents of this document? If you trust the s; buttons=Yes,No"}
{"name": "398-hyperlink-RELS", "sha256": "440b780d13f48f405af64eb498d30fba0cb997425bc837923216316657114499", "observedAt": "2026-10-24T03:10:30.750Z", "outcome": "clean-open", "details": "processAlive=true; documents=1; names=docx-verify-440b780d13f4-76dd9f42.docx; buttons=missing value,missing value,missing value"}
{"name": "397-encoding-unknown-styles", "sha256": "d2c61a613863872e925e6ad9bff68fa6504b40dacf914d14d022da8cc899dd91", "observedAt": "2026-10-24T03:10:38.075Z", "outcome": "repair-warning", "details": "processAlive=true; documents=0; dialog=Microsoft Word Word found unreadable content in docx-verify-d2c61a613863-25dfa1c5.docx. Do you want to recover the contents of this document? If you trust the s; buttons=Yes,No"}
{"name": "397-encoding-unknown-content-types", "sha256": "a2be2959ad591d9a05198194d558d46633ab45fc3cec5255e96aa0e3a62e21d6", "observedAt": "2026-10-24T03:10:41.809Z", "outcome": "repair-warning", "details": "processAlive=true; documents=0; dialog=Microsoft Word Word found unreadable content in docx-verify-a2be2959ad59-5c626fc1.docx. Do you want to recover the contents of this document? If you trust the s; buttons=Yes,No"}
{"name": "397-encoding-unknown-package-rels", "sha256": "c8a06454f5c52111b62db510785a73580b46cde2901bfa35cfaade2c3d563239", "observedAt": "2026-10-24T03:10:45.804Z", "outcome": "repair-warning", "details": "processAlive=true; documents=0; dialog=Microsoft Word Word found unreadable content in docx-verify-c8a06454f5c5-7c6602f9.docx. Do you want to recover the contents of this document? If you trust the s; buttons=Yes,No"}
{"name": "397-encoding-unknown-core", "sha256": "bde4363f2fed0946d3c27b124d6b8c4d73332c190f269345a0ebecac541168a6", "observedAt": "2026-10-24T03:10:49.836Z", "outcome": "repair-warning", "details": "processAlive=true; documents=0; dialog=Microsoft Word Word found unreadable content in docx-verify-bde4363f2fed-2869ea96.docx. Do you want to recover the contents of this document? If you trust the s; buttons=Yes,No"}
{"name": "397-invalid-utf8-bytes", "sha256": "87b043160608a0aaab2944629c3bb96f204ab7b1e7f923f82066851eb02f037a", "observedAt": "2026-10-24T03:10:53.694Z", "outcome": "open-error", "details": "processAlive=true; documents=0; dialog=Microsoft Word Word experienced an error trying to open the file.; buttons=OK"}
{"name": "esc-unreferenced-extra-item", "sha256": "02392f4064145df1ab5015e564bed05841e805071ce9a306cca94f71cc975c56", "observedAt": "2026-10-24T03:11:55.380Z", "outcome": "repair-warning", "details": "processAlive=true; documents=0; dialog=Microsoft Word Word found unreadable content in docx-verify-02392f406414-5510bdc.docx. Do you want to recover the contents of this document? If you trust the so; buttons=Yes,No"}
{"name": "esc-styles-missing-control", "sha256": "79dddb1d4ffcd1caf0f6238942fcc5af4f6a5f9d97e5499fd3b064f78e66a818", "observedAt": "2026-10-24T03:12:04.007Z", "outcome": "clean-open", "details": "processAlive=true; documents=1; names=docx-verify-79dddb1d4ffc-4ce497ad.docx; buttons=missing value,missing value,missing value"}
{"name": "esc-only-main", "sha256": "845abca15ef137feff7301f1809b31a5a63f03c50854d6404d5e0a90236d856f", "observedAt": "2026-10-24T03:12:13.526Z", "outcome": "repair-warning", "details": "processAlive=true; documents=0; dialog=Microsoft Word Word found unreadable content in docx-verify-845abca15ef1-6a85ef09.docx. Do you want to recover the contents of this document? If you trust the s; buttons=Yes,No"}
{"name": "docprops-core-truncated", "sha256": "0f729853110d309f0857d4a4784ff7d98f9d23f1b581bf1526c56b536119e94b", "observedAt": "2026-10-24T03:19:10.366Z", "outcome": "repair-warning", "details": "processAlive=true; documents=0; dialog=Microsoft Word Word found unreadable content in docx-verify-0f729853110d-654190a6.docx. Do you want to recover the contents of this document? If you trust the s; buttons=Yes,No"}
{"name": "docprops-app-control", "sha256": "727b58abd042452e1587525f3932f401bfc4b5f9160d526779718b1d31a0b436", "observedAt": "2026-10-24T03:19:22.290Z", "outcome": "clean-open", "details": "processAlive=true; documents=1; names=docx-verify-727b58abd042-4fd3d74d.docx; buttons=missing value,missing value,missing value"}
```

</details>

### Session 3b: 05:02Z–05:18Z, validator columns at ab7f311d

The first eight rows rebuild session 1 and 2 specimens to calibrate the driver against `verifyInWord`: `control-minimal`, `397-encoding-windows-1252-control`, `351-dup-case-identical`, `cti-override-existing-core-control`, `398-hyperlink-RELS`, `docprops-core-truncated`, `esc-unreferenced-extra-item` and `397-encoding-unknown-core`. All eight reproduce their earlier outcomes. `s3-control-minimal-start` and `s3-control-minimal-end` bracket the new specimens.

| Specimen                              | Issue      | SHA-256 (first 16) | Word           | word-valid at ab7f311d                                                                                                             | strict at ab7f311d                                                                                                                 |
| ------------------------------------- | ---------- | ------------------ | -------------- | ---------------------------------------------------------------------------------------------------------------------------------- | ---------------------------------------------------------------------------------------------------------------------------------- |
| `control-minimal`                     | control    | `0160dfc1db4c5ce1` | clean-open     | valid                                                                                                                              | valid                                                                                                                              |
| `397-encoding-windows-1252-control`   | #397       | `3ccb84b7f1be0b88` | clean-open     | valid                                                                                                                              | valid                                                                                                                              |
| `351-dup-case-identical`              | #351       | `4400bc4cd6b8d23a` | repair-warning | invalid (`opc-equivalent-part-names@word/document.xml`, `package-extract-failed`)                                                  | invalid (`opc-equivalent-part-names@word/document.xml`, `package-extract-failed`)                                                  |
| `cti-override-existing-core-control`  | #401       | `ae55ebebdd97a971` | clean-open     | valid                                                                                                                              | valid                                                                                                                              |
| `398-hyperlink-RELS`                  | #398       | `18b773c0d5d7ec6f` | clean-open     | valid                                                                                                                              | valid                                                                                                                              |
| `docprops-core-truncated`             | #409       | `c8eb3d968d217f0c` | repair-warning | valid                                                                                                                              | invalid (`xml-syntax@docProps/core.xml`, `xml-undeclared-prefix@docProps/core.xml`)                                                |
| `esc-unreferenced-extra-item`         | #408       | `d8ad7b6fa1e05937` | repair-warning | invalid (`opc-part-name-escaped-unreserved@word/%78tra.xml`)                                                                       | invalid (`rels-unreferenced@word/%78tra.xml`, `opc-part-name-escaped-unreserved@word/%78tra.xml`)                                  |
| `397-encoding-unknown-core`           | #397       | `031ee748acdee96d` | repair-warning | valid                                                                                                                              | invalid (`xml-encoding@docProps/core.xml`)                                                                                         |
| `s3-control-minimal-start`            | control    | `3cdc3192bc053374` | clean-open     | valid                                                                                                                              | valid                                                                                                                              |
| `docprops-app-truncated`              | #409       | `96ef7c3a11c2d06a` | repair-warning | valid                                                                                                                              | invalid (`xml-syntax@docProps/app.xml`)                                                                                            |
| `docprops-app-encoding-unknown`       | #409       | `3a3d526f35194415` | repair-warning | valid                                                                                                                              | invalid (`xml-encoding@docProps/app.xml`)                                                                                          |
| `docprops-app-undeclared-prefix`      | #409       | `ee29efbb9af0ee63` | repair-warning | valid                                                                                                                              | invalid (`xml-undeclared-prefix@docProps/app.xml`)                                                                                 |
| `docprops-core-undeclared-prefix`     | #409       | `9692d5b7434cea96` | repair-warning | valid                                                                                                                              | invalid (`xml-undeclared-prefix@docProps/core.xml`)                                                                                |
| `docprops-custom-control`             | #409       | `a82099eacbd6b8ed` | clean-open     | valid                                                                                                                              | valid                                                                                                                              |
| `docprops-custom-truncated`           | #409       | `2619e0980a11bbf6` | repair-warning | valid                                                                                                                              | invalid (`xml-syntax@docProps/custom.xml`)                                                                                         |
| `docprops-custom-undeclared-prefix`   | #409       | `de4ca7ce3c73f670` | repair-warning | valid                                                                                                                              | invalid (`xml-undeclared-prefix@docProps/custom.xml`)                                                                              |
| `docprops-custom-encoding-unknown`    | #409       | `1d03dad3ff19c572` | repair-warning | valid                                                                                                                              | invalid (`xml-encoding@docProps/custom.xml`)                                                                                       |
| `410-macintosh`                       | #410       | `9f5a2e272d2e34cb` | open-error     | invalid (`xml-encoding@word/document.xml`)                                                                                         | invalid (`xml-encoding@word/document.xml`)                                                                                         |
| `410-x-mac-cyrillic`                  | #410       | `81d8f59f16ff8e5d` | open-error     | invalid (`xml-encoding@word/document.xml`)                                                                                         | invalid (`xml-encoding@word/document.xml`)                                                                                         |
| `410-IBM866`                          | #410       | `b48036e2d2f7d93c` | open-error     | invalid (`xml-encoding@word/document.xml`)                                                                                         | invalid (`xml-encoding@word/document.xml`)                                                                                         |
| `410-UTF-7`                           | #410       | `6bf6863458ce8038` | open-error     | invalid (`xml-encoding@word/document.xml`)                                                                                         | invalid (`xml-encoding@word/document.xml`)                                                                                         |
| `410-windows-1252-nonascii-control`   | #410       | `e004e8ebb0082fbd` | clean-open     | valid                                                                                                                              | valid                                                                                                                              |
| `411-identical-valid-valid`           | #411       | `ea01967dbdd15786` | repair-warning | invalid (`opc-equivalent-part-names@word/document.xml`)                                                                            | invalid (`opc-equivalent-part-names@word/document.xml`)                                                                            |
| `411-identical-valid-then-malformed`  | #411       | `3ac694002959888c` | repair-warning | invalid (`xml-undeclared-prefix@word/document.xml`, `xml-syntax@word/document.xml`, `opc-equivalent-part-names@word/document.xml`) | invalid (`xml-undeclared-prefix@word/document.xml`, `xml-syntax@word/document.xml`, `opc-equivalent-part-names@word/document.xml`) |
| `411-identical-malformed-then-valid`  | #411       | `4ada736ae140496e` | repair-warning | invalid (`opc-equivalent-part-names@word/document.xml`)                                                                            | invalid (`opc-equivalent-part-names@word/document.xml`)                                                                            |
| `411-dotdot-unreferenced-item`        | #411       | `97f752f1b03e1749` | repair-warning | invalid (`package-extract-failed`)                                                                                                 | invalid (`package-extract-failed`)                                                                                                 |
| `411-dotdot-then-plain-styles`        | #411       | `e68121f4cf107657` | repair-warning | invalid (`package-extract-failed`)                                                                                                 | invalid (`package-extract-failed`)                                                                                                 |
| `414-styles-override-removed-control` | #432       | `e0595fa4e2e464ec` | repair-warning | valid                                                                                                                              | valid                                                                                                                              |
| `414-styles-override-dotdot`          | #414       | `1fa751332e19b8f0` | repair-warning | valid                                                                                                                              | valid                                                                                                                              |
| `414-styles-override-query`           | #414       | `75810dc3e308696a` | repair-warning | valid                                                                                                                              | valid                                                                                                                              |
| `414-styles-override-fragment`        | #414       | `7a16f99036a62688` | repair-warning | valid                                                                                                                              | valid                                                                                                                              |
| `414-styles-override-backslash`       | #414       | `4e5e85d63f2ee9c3` | repair-warning | valid                                                                                                                              | valid                                                                                                                              |
| `414-invalid-x-dotdot-main`           | #414       | `f7bae552bbd4fd8a` | repair-warning | valid                                                                                                                              | invalid (`xsd-summary@[Content_Types].xml`, `xsd-error@[Content_Types].xml`)                                                       |
| `414-invalid-x-query-main`            | #414       | `0ceb5371ecc1714b` | repair-warning | valid                                                                                                                              | invalid (`xsd-summary@[Content_Types].xml`, `xsd-error@[Content_Types].xml`)                                                       |
| `414-invalid-x-fragment-main`         | #414       | `00fb33975b0027c9` | repair-warning | valid                                                                                                                              | invalid (`xsd-summary@[Content_Types].xml`, `xsd-error@[Content_Types].xml`)                                                       |
| `414-invalid-x-backslash-main`        | #414       | `f6451c46a8323938` | repair-warning | valid                                                                                                                              | invalid (`xsd-summary@[Content_Types].xml`, `xsd-error@[Content_Types].xml`)                                                       |
| `dupdecl-default-identical`           | #431       | `222d7ee2fc43b1f6` | repair-warning | valid                                                                                                                              | valid                                                                                                                              |
| `dupdecl-default-case`                | #431       | `fe596b4f8f32a55c` | repair-warning | valid                                                                                                                              | valid                                                                                                                              |
| `dupdecl-override-identical`          | #431       | `92c289efb80c1b41` | repair-warning | valid                                                                                                                              | valid                                                                                                                              |
| `dupdecl-override-conflicting`        | #431       | `1fe83fb76ef60ec2` | repair-warning | valid                                                                                                                              | valid                                                                                                                              |
| `dupdecl-override-case`               | #431       | `2f9f7277561b364b` | repair-warning | valid                                                                                                                              | valid                                                                                                                              |
| `esc-plain-unreferenced-control`      | #408       | `6e63642feef861ab` | clean-open     | valid                                                                                                                              | invalid (`rels-unreferenced@word/xtra.xml`)                                                                                        |
| `esc-%41-upper-letter`                | #408       | `fae72ee8b638fa8a` | repair-warning | invalid (`opc-part-name-escaped-unreserved@word/x%41.xml`)                                                                         | invalid (`rels-unreferenced@word/x%41.xml`, `opc-part-name-escaped-unreserved@word/x%41.xml`)                                      |
| `esc-%30-digit`                       | #408       | `17d763961bc3f58c` | repair-warning | invalid (`opc-part-name-escaped-unreserved@word/x%30.xml`)                                                                         | invalid (`rels-unreferenced@word/x%30.xml`, `opc-part-name-escaped-unreserved@word/x%30.xml`)                                      |
| `esc-%2D-hyphen`                      | #408       | `84a0e0f1843139fd` | repair-warning | invalid (`opc-part-name-escaped-unreserved@word/x%2D.xml`)                                                                         | invalid (`rels-unreferenced@word/x%2D.xml`, `opc-part-name-escaped-unreserved@word/x%2D.xml`)                                      |
| `esc-%2d-hyphen-lowerhex`             | #408       | `37529d0e79b352aa` | repair-warning | invalid (`opc-part-name-escaped-unreserved@word/x%2d.xml`)                                                                         | invalid (`rels-unreferenced@word/x%2d.xml`, `opc-part-name-escaped-unreserved@word/x%2d.xml`)                                      |
| `esc-%2E-period`                      | #408       | `fcbdbbf79faa0bcd` | repair-warning | invalid (`opc-part-name-escaped-unreserved@word/x%2E.xml`)                                                                         | invalid (`rels-unreferenced@word/x%2E.xml`, `opc-part-name-escaped-unreserved@word/x%2E.xml`)                                      |
| `esc-%5F-underscore`                  | #408       | `7d0278c4bf9fd3fa` | repair-warning | invalid (`opc-part-name-escaped-unreserved@word/x%5F.xml`)                                                                         | invalid (`rels-unreferenced@word/x%5F.xml`, `opc-part-name-escaped-unreserved@word/x%5F.xml`)                                      |
| `esc-%7E-tilde`                       | #408       | `4436f009f2500d57` | repair-warning | invalid (`opc-part-name-escaped-unreserved@word/x%7E.xml`)                                                                         | invalid (`rels-unreferenced@word/x%7E.xml`, `opc-part-name-escaped-unreserved@word/x%7E.xml`)                                      |
| `esc-%20-space-control`               | #408       | `c7598087941b0c35` | clean-open     | valid                                                                                                                              | invalid (`rels-unreferenced@word/x%20y.xml`)                                                                                       |
| `s3-control-minimal-end`              | control    | `00d28798c1f66dc7` | clean-open     | valid                                                                                                                              | valid                                                                                                                              |
| `rels-main-nested-in-unknown-element` | #433, #434 | `176e6d8f133812b3` | repair-warning | valid                                                                                                                              | invalid (`xsd-summary@_rels/.rels`, `xsd-error@_rels/.rels`)                                                                       |
| `rels-root-not-relationships`         | #433       | `1a5c659713cf7f9b` | repair-warning | valid                                                                                                                              | invalid (`xsd-summary@_rels/.rels`, `xsd-error@_rels/.rels`)                                                                       |
| `rels-unknown-sibling-control`        | #433       | `27cb9b0032c7b10c` | inconclusive   | valid                                                                                                                              | invalid (`xsd-summary@_rels/.rels`, `xsd-error@_rels/.rels`)                                                                       |

### Session 3c: 05:19Z–05:27Z, validator columns at ab7f311d

| Specimen                                         | Issue      | SHA-256 (first 16) | Word           | word-valid at ab7f311d                                                      | strict at ab7f311d                                                                                                          |
| ------------------------------------------------ | ---------- | ------------------ | -------------- | --------------------------------------------------------------------------- | --------------------------------------------------------------------------------------------------------------------------- |
| `s3c-control-minimal`                            | control    | `1a414a1109291a18` | clean-open     | valid                                                                       | valid                                                                                                                       |
| `ct-stream-missing-control`                      | #435       | `a4c6ecaf2f422b09` | repair-warning | valid                                                                       | invalid (`ct-missing`)                                                                                                      |
| `ct-stream-escaped-only`                         | #408, #435 | `42d70cff0d0d0e1a` | repair-warning | invalid (`opc-part-name-escaped-unreserved@[Content_Types].%78ml`)          | invalid (`rels-unreferenced@[Content_Types].%78ml`, `ct-missing`, `opc-part-name-escaped-unreserved@[Content_Types].%78ml`) |
| `ct-stream-escaped-extra`                        | #408       | `898d49c3184b8458` | repair-warning | invalid (`opc-part-name-escaped-unreserved@[Content_Types].%78ml`)          | invalid (`rels-unreferenced@[Content_Types].%78ml`, `opc-part-name-escaped-unreserved@[Content_Types].%78ml`)               |
| `esc-pct2F-slash-unreferenced`                   | #436       | `1003d9e0b112fcab` | repair-warning | valid                                                                       | invalid (`rels-unreferenced@word%2Fxtra.xml`)                                                                               |
| `esc-pct5C-backslash-unreferenced`               | #436       | `76f9efc51af63338` | open-error     | valid                                                                       | invalid (`rels-unreferenced@word%5Cxtra.xml`)                                                                               |
| `docprops-custom-unreferenced-undeclared-prefix` | #409       | `0a49797ea81f27d6` | clean-open     | valid                                                                       | invalid (`xml-undeclared-prefix@docProps/custom.xml`, `rels-unreferenced@docProps/custom.xml`)                              |
| `docprops-custom-unreferenced-truncated`         | #409       | `050e2698f564c761` | clean-open     | valid                                                                       | invalid (`xml-syntax@docProps/custom.xml`, `rels-unreferenced@docProps/custom.xml`)                                         |
| `docprops-app-unreferenced-truncated`            | #409       | `57b78d2872e4569a` | clean-open     | valid                                                                       | invalid (`xml-syntax@docProps/app.xml`, `rels-unreferenced@docProps/app.xml`)                                               |
| `word-unreferenced-malformed-part`               | #437       | `8968d3be05c61519` | clean-open     | invalid (`xml-syntax@word/xtra.xml`, `xml-undeclared-prefix@word/xtra.xml`) | invalid (`xml-syntax@word/xtra.xml`, `xml-undeclared-prefix@word/xtra.xml`, `rels-unreferenced@word/xtra.xml`)              |
| `word-unreferenced-wellformed-control`           | #437       | `7b8e0e0d4ae48b38` | clean-open     | valid                                                                       | invalid (`rels-unreferenced@word/xtra.xml`)                                                                                 |
| `docprops-core-unreferenced-truncated`           | #409       | `cefef105ee09a3c0` | clean-open     | valid                                                                       | invalid (`xml-syntax@docProps/core.xml`, `xml-undeclared-prefix@docProps/core.xml`, `rels-unreferenced@docProps/core.xml`)  |
| `docprops-core-renamed-truncated`                | #409       | `e0290e52e694e983` | repair-warning | valid                                                                       | invalid (`xml-syntax@props/meta.xml`, `xml-undeclared-prefix@props/meta.xml`)                                               |
| `docprops-core-renamed-control`                  | #409       | `de166f9b02fa683d` | clean-open     | valid                                                                       | valid                                                                                                                       |
| `empty-archive`                                  | #380       | `8739c76e681f9009` | repair-warning | valid                                                                       | invalid (`ct-missing`, `word-main-document-part`)                                                                           |
| `main-part-missing`                              | #380       | `035e63734f8226d0` | repair-warning | valid                                                                       | invalid (`rels-broken@_rels/.rels`, `rels-target-missing@_rels/.rels`, `word-main-document-part@_rels/.rels`)               |
| `main-relationship-missing`                      | #380       | `6cff8f9d94dc4883` | inconclusive   | valid                                                                       | invalid (`rels-unreferenced@word/document.xml`, `rels-empty-part@_rels/.rels`, `word-main-document-part@_rels/.rels`)       |
| `414-invalid-x-plain-missing-control`            | #414       | `dc5dafa7e6a85c49` | clean-open     | valid                                                                       | invalid (`xsd-summary@[Content_Types].xml`, `xsd-error@[Content_Types].xml`)                                                |
| `414-invalid-x-dotdot-missing`                   | #414       | `2bd531b03fb63a74` | repair-warning | valid                                                                       | invalid (`xsd-summary@[Content_Types].xml`, `xsd-error@[Content_Types].xml`)                                                |
| `414-invalid-x-query-missing`                    | #414       | `8c41377382c63da1` | open-error     | valid                                                                       | invalid (`xsd-summary@[Content_Types].xml`, `xsd-error@[Content_Types].xml`)                                                |
| `414-invalid-x-fragment-missing`                 | #414       | `ae39b41b00a82878` | repair-warning | valid                                                                       | invalid (`xsd-summary@[Content_Types].xml`, `xsd-error@[Content_Types].xml`)                                                |
| `414-invalid-x-backslash-missing`                | #414       | `8fc0c3015afe29b2` | repair-warning | valid                                                                       | invalid (`xsd-summary@[Content_Types].xml`, `xsd-error@[Content_Types].xml`)                                                |
| `rels-unknown-sibling-control`                   | #433       | `43205d2359aaac36` | repair-warning | valid                                                                       | invalid (`xsd-summary@_rels/.rels`, `xsd-error@_rels/.rels`)                                                                |

Full hashes, the driver's per-file result and every alert that named the file are in the session 3 records:

<details><summary>Raw records, session 3</summary>

```jsonl
{"name": "control-minimal", "sha256": "0160dfc1db4c5ce1e1f87f3b7547cae3bbe64240705ada3e54be8897de00ac1a", "observedAt": "2026-10-24T05:18:09+00:00", "outcome": "clean-open", "item": "ok", "error": "", "dialogs": []}
{"name": "397-encoding-windows-1252-control", "sha256": "3ccb84b7f1be0b8885f311c6bb9aaa00e477e53780d16043ca6ea85cf766c633", "observedAt": "2026-10-24T05:18:09+00:00", "outcome": "clean-open", "item": "ok", "error": "", "dialogs": []}
{"name": "351-dup-case-identical", "sha256": "4400bc4cd6b8d23ab6e049f1596436e744f9630cd9e3152f630f36dc63e1cb42", "observedAt": "2026-10-24T05:02:46.447Z", "outcome": "repair-warning", "item": "failed", "error": "repair prompt declined; Word could not read it", "dialogs": ["'Word found unreadable content in 00000__wfv__351-dup-case-identical.docx. Do you want to recover the contents of this document? If you trust the source of this document, click Yes.'", "\"Word found unreadable content in 00000__wfv__351-dup-case-identical.docx. Do you want to recover the contents of this document? If you trust the source of this document, click Yes.\"", "\"Word experienced an error trying to open the file. / Try these suggestions. / * Check the file permissions for the document or drive. / * Make sure there is sufficient free memory and disk space. / * Open the file with the Text Recovery converter."]}
{"name": "cti-override-existing-core-control", "sha256": "ae55ebebdd97a9712fcd7274520ebe8560b6acddd4443e07489780307d49181c", "observedAt": "2026-10-24T05:18:09+00:00", "outcome": "clean-open", "item": "ok", "error": "", "dialogs": []}
{"name": "398-hyperlink-RELS", "sha256": "18b773c0d5d7ec6ff1b7127a461bc1a1852fbc644f0ce4df5e36079f6d4bde22", "observedAt": "2026-10-24T05:18:09+00:00", "outcome": "clean-open", "item": "ok", "error": "", "dialogs": []}
{"name": "docprops-core-truncated", "sha256": "c8eb3d968d217f0c4f5a67a9770f246a81477f8c26093a48c6385a4f3aeda444", "observedAt": "2026-10-24T05:09:34.682Z", "outcome": "repair-warning", "item": "failed", "error": "repair prompt declined; Word could not read it", "dialogs": ["'Word found unreadable content in 00028__wfv__docprops-core-truncated.docx. Do you want to recover the contents of this document? If you trust the source of this document, click Yes.'", "\"Word found unreadable content in 00028__wfv__docprops-core-truncated.docx. Do you want to recover the contents of this document? If you trust the source of this document, click Yes.\"", "\"Word experienced an error trying to open the file. / Try these suggestions. / * Check the file permissions for the document or drive. / * Make sure there is sufficient free memory and disk space. / * Open the file with the Text Recovery converter."]}
{"name": "esc-unreferenced-extra-item", "sha256": "d8ad7b6fa1e059379c1b0b0646261752d8184363fd2a8ea60f532a56960c9d53", "observedAt": "2026-10-24T05:16:54.733Z", "outcome": "repair-warning", "item": "failed", "error": "repair prompt declined; Word could not read it: open produced 0 new documents, none named 00048__wfv__esc-unreferenced-extra-item.docx; refusing to guess", "dialogs": ["'Word found unreadable content in 00048__wfv__esc-unreferenced-extra-item.docx. Do you want to recover the contents of this document? If you trust the source of this document, click Yes.'", "\"Word found unreadable content in 00048__wfv__esc-unreferenced-extra-item.docx. Do you want to recover the contents of this document? If you trust the source of this document, click Yes.\""]}
{"name": "397-encoding-unknown-core", "sha256": "031ee748acdee96d3007fc1d4be0fadfba9504d567509e86a9f630593366054e", "observedAt": "2026-10-24T05:02:49.147Z", "outcome": "repair-warning", "item": "failed", "error": "repair prompt declined; Word could not read it: open produced 0 new documents, none named 00001__wfv__397-encoding-unknown-core.docx; refusing to guess", "dialogs": ["'Word found unreadable content in 00001__wfv__397-encoding-unknown-core.docx. Do you want to recover the contents of this document? If you trust the source of this document, click Yes.'", "\"Word found unreadable content in 00001__wfv__397-encoding-unknown-core.docx. Do you want to recover the contents of this document? If you trust the source of this document, click Yes.\"", "\"Word experienced an error trying to open the file. / Try these suggestions. / * Check the file permissions for the document or drive. / * Make sure there is sufficient free memory and disk space. / * Open the file with the Text Recovery converter."]}
{"name": "s3-control-minimal-start", "sha256": "3cdc3192bc053374381bfd80959775fa792556e3f36fa7c884da45b1fba6d41b", "observedAt": "2026-10-24T05:18:09+00:00", "outcome": "clean-open", "item": "ok", "error": "", "dialogs": []}
{"name": "docprops-app-truncated", "sha256": "96ef7c3a11c2d06add3f28257aeb145dcfdfd297e7b43b48de010da48e252a60", "observedAt": "2026-10-24T05:09:26.187Z", "outcome": "repair-warning", "item": "failed", "error": "repair prompt declined; Word could not read it", "dialogs": ["'Word found unreadable content in 00026__wfv__docprops-app-truncated.docx. Do you want to recover the contents of this document? If you trust the source of this document, click Yes.'", "\"Word found unreadable content in 00026__wfv__docprops-app-truncated.docx. Do you want to recover the contents of this document? If you trust the source of this document, click Yes.\"", "\"Word experienced an error trying to open the file. / Try these suggestions. / * Check the file permissions for the document or drive. / * Make sure there is sufficient free memory and disk space. / * Open the file with the Text Recovery converter."]}
{"name": "docprops-app-encoding-unknown", "sha256": "3a3d526f3519441538b7e89522c459cc5b1a1ac9465825dede6e1ad32dd7e848", "observedAt": "2026-10-24T05:09:21.748Z", "outcome": "repair-warning", "item": "failed", "error": "repair prompt declined; Word could not read it", "dialogs": ["'Word found unreadable content in 00025__wfv__docprops-app-encoding-unknown.docx. Do you want to recover the contents of this document? If you trust the source of this document, click Yes.'", "\"Word found unreadable content in 00025__wfv__docprops-app-encoding-unknown.docx. Do you want to recover the contents of this document? If you trust the source of this document, click Yes.\"", "\"Word experienced an error trying to open the file. / Try these suggestions. / * Check the file permissions for the document or drive. / * Make sure there is sufficient free memory and disk space. / * Open the file with the Text Recovery converter."]}
{"name": "docprops-app-undeclared-prefix", "sha256": "ee29efbb9af0ee63317e2f662bf312bc409b2ea77d3dd626966f23498bd49758", "observedAt": "2026-10-24T05:09:30.416Z", "outcome": "repair-warning", "item": "failed", "error": "repair prompt declined; Word could not read it", "dialogs": ["'Word found unreadable content in 00027__wfv__docprops-app-undeclared-prefix.docx. Do you want to recover the contents of this document? If you trust the source of this document, click Yes.'", "\"Word found unreadable content in 00027__wfv__docprops-app-undeclared-prefix.docx. Do you want to recover the contents of this document? If you trust the source of this document, click Yes.\"", "\"Word experienced an error trying to open the file. / Try these suggestions. / * Check the file permissions for the document or drive. / * Make sure there is sufficient free memory and disk space. / * Open the file with the Text Recovery converter."]}
{"name": "docprops-core-undeclared-prefix", "sha256": "9692d5b7434cea969c889bdc12d72b20e1e4cb6a74aa0c836cf566ebc86fb561", "observedAt": "2026-10-24T05:09:38.866Z", "outcome": "repair-warning", "item": "failed", "error": "repair prompt declined; Word could not read it", "dialogs": ["'Word found unreadable content in 00029__wfv__docprops-core-undeclared-prefix.docx. Do you want to recover the contents of this document? If you trust the source of this document, click Yes.'", "\"Word found unreadable content in 00029__wfv__docprops-core-undeclared-prefix.docx. Do you want to recover the contents of this document? If you trust the source of this document, click Yes.\"", "\"Word experienced an error trying to open the file. / Try these suggestions. / * Check the file permissions for the document or drive. / * Make sure there is sufficient free memory and disk space. / * Open the file with the Text Recovery converter."]}
{"name": "docprops-custom-control", "sha256": "a82099eacbd6b8ed10a40115643c17e3881ff4458a72e2e2f6e3e812c949bb56", "observedAt": "2026-10-24T05:18:09+00:00", "outcome": "clean-open", "item": "ok", "error": "", "dialogs": []}
{"name": "docprops-custom-truncated", "sha256": "2619e0980a11bbf6b59e2b9feeb8eb28626a0e96d4ddc9bc873b8f319a59521e", "observedAt": "2026-10-24T05:09:47.518Z", "outcome": "repair-warning", "item": "failed", "error": "repair prompt declined; Word could not read it", "dialogs": ["'Word found unreadable content in 00032__wfv__docprops-custom-truncated.docx. Do you want to recover the contents of this document? If you trust the source of this document, click Yes.'", "\"Word found unreadable content in 00032__wfv__docprops-custom-truncated.docx. Do you want to recover the contents of this document? If you trust the source of this document, click Yes.\"", "\"Word experienced an error trying to open the file. / Try these suggestions. / * Check the file permissions for the document or drive. / * Make sure there is sufficient free memory and disk space. / * Open the file with the Text Recovery converter."]}
{"name": "docprops-custom-undeclared-prefix", "sha256": "de4ca7ce3c73f6706526db5e66e4c36272295a89d071dc2c8271b531f475e4fd", "observedAt": "2026-10-24T05:09:51.735Z", "outcome": "repair-warning", "item": "failed", "error": "repair prompt declined; Word could not read it", "dialogs": ["'Word found unreadable content in 00033__wfv__docprops-custom-undeclared-prefix.docx. Do you want to recover the contents of this document? If you trust the source of this document, click Yes.'", "\"Word found unreadable content in 00033__wfv__docprops-custom-undeclared-prefix.docx. Do you want to recover the contents of this document? If you trust the source of this document, click Yes.\"", "\"Word experienced an error trying to open the file. / Try these suggestions. / * Check the file permissions for the document or drive. / * Make sure there is sufficient free memory and disk space. / * Open the file with the Text Recovery converter."]}
{"name": "docprops-custom-encoding-unknown", "sha256": "1d03dad3ff19c572be1997564fb4b08e6ede0e796e82624124756089ecf6ab5c", "observedAt": "2026-10-24T05:09:43.317Z", "outcome": "repair-warning", "item": "failed", "error": "repair prompt declined; Word could not read it", "dialogs": ["'Word found unreadable content in 00031__wfv__docprops-custom-encoding-unknown.docx. Do you want to recover the contents of this document? If you trust the source of this document, click Yes.'", "\"Word found unreadable content in 00031__wfv__docprops-custom-encoding-unknown.docx. Do you want to recover the contents of this document? If you trust the source of this document, click Yes.\"", "\"Word experienced an error trying to open the file. / Try these suggestions. / * Check the file permissions for the document or drive. / * Make sure there is sufficient free memory and disk space. / * Open the file with the Text Recovery converter."]}
{"name": "410-macintosh", "sha256": "9f5a2e272d2e34cb55b14af966ba75d72cafd90595825dd3c3efabfb981105d7", "observedAt": "2026-10-24T05:03:15.432Z", "outcome": "open-error", "item": "failed", "error": "Word could not load it in 2 passes: open produced 0 new documents, none named 00006__wfv__410-macintosh.docx; refusing to guess", "dialogs": ["'Word experienced an error trying to open the file.\\nTry these suggestions.\\n* Check the file permissions for the document or drive.\\n* Make sure there is sufficient free memory and disk space.\\n* Open the file with the Text Recovery conver", "\"Word experienced an error trying to open the file. / Try these suggestions. / * Check the file permissions for the document or drive. / * Make sure there is sufficient free memory and disk space. / * Open the file with the Text Recovery converter."]}
{"name": "410-x-mac-cyrillic", "sha256": "81d8f59f16ff8e5d0d4ddb78556c080beea84718855b72827f8ef4c8d5da6eab", "observedAt": "2026-10-24T05:15:32.619Z", "outcome": "open-error", "item": "failed", "error": "Word could not load it in 2 passes: open produced 0 new documents, none named 00008__wfv__410-x-mac-cyrillic.docx; refusing to guess", "dialogs": ["\"Word experienced an error trying to open the file. / Try these suggestions. / * Check the file permissions for the document or drive. / * Make sure there is sufficient free memory and disk space. / * Open the file with the Text Recovery converter."]}
{"name": "410-IBM866", "sha256": "b48036e2d2f7d93cbbae030e440f8b6e0e6263a8afc8fa32d57e4cd51abf6ebf", "observedAt": "2026-10-24T05:03:06.975Z", "outcome": "open-error", "item": "failed", "error": "Word could not load it in 2 passes: open produced 0 new documents, none named 00004__wfv__410-IBM866.docx; refusing to guess", "dialogs": ["'Word experienced an error trying to open the file.\\nTry these suggestions.\\n* Check the file permissions for the document or drive.\\n* Make sure there is sufficient free memory and disk space.\\n* Open the file with the Text Recovery conver", "\"Word experienced an error trying to open the file. / Try these suggestions. / * Check the file permissions for the document or drive. / * Make sure there is sufficient free memory and disk space. / * Open the file with the Text Recovery converter."]}
{"name": "410-UTF-7", "sha256": "6bf6863458ce8038a6139145bcd8f8e70ba5a2caf44906f3424c0da9956b4d2f", "observedAt": "2026-10-24T05:03:11.114Z", "outcome": "open-error", "item": "failed", "error": "Word could not load it in 2 passes: open produced 0 new documents, none named 00005__wfv__410-UTF-7.docx; refusing to guess", "dialogs": ["'Word experienced an error trying to open the file.\\nTry these suggestions.\\n* Check the file permissions for the document or drive.\\n* Make sure there is sufficient free memory and disk space.\\n* Open the file with the Text Recovery conver", "\"Word experienced an error trying to open the file. / Try these suggestions. / * Check the file permissions for the document or drive. / * Make sure there is sufficient free memory and disk space. / * Open the file with the Text Recovery converter."]}
{"name": "410-windows-1252-nonascii-control", "sha256": "e004e8ebb0082fbd1057891c719da2bfb677d78531af0c32f4e454360287441c", "observedAt": "2026-10-24T05:18:09+00:00", "outcome": "clean-open", "item": "ok", "error": "", "dialogs": []}
{"name": "411-identical-valid-valid", "sha256": "ea01967dbdd157861518144f062b569b30a5a00f14db4e4a57c96d415dd9ce12", "observedAt": "2026-10-24T05:04:51.295Z", "outcome": "repair-warning", "item": "failed", "error": "repair prompt declined; Word could not read it", "dialogs": ["'Word found unreadable content in 00013__wfv__411-identical-valid-valid.docx. Do you want to recover the contents of this document? If you trust the source of this document, click Yes.'", "\"Word found unreadable content in 00013__wfv__411-identical-valid-valid.docx. Do you want to recover the contents of this document? If you trust the source of this document, click Yes.\"", "\"Word experienced an error trying to open the file. / Try these suggestions. / * Check the file permissions for the document or drive. / * Make sure there is sufficient free memory and disk space. / * Open the file with the Text Recovery converter."]}
{"name": "411-identical-valid-then-malformed", "sha256": "3ac694002959888c7edcabe5a0b2809f4128da71e3f2443b372eb8dc5fc440cb", "observedAt": "2026-10-24T05:04:45.803Z", "outcome": "repair-warning", "item": "failed", "error": "repair prompt declined; Word could not read it", "dialogs": ["\"Word found unreadable content in 00012__wfv__411-identical-valid-then-malformed.docx. Do you want to recover the contents of this document? If you trust the source of this document, click Yes.\""]}
{"name": "411-identical-malformed-then-valid", "sha256": "4ada736ae140496eb2dad6d806719a8580e77de6f999d30f4136dec1b44cb37d", "observedAt": "2026-10-24T05:03:53.717Z", "outcome": "repair-warning", "item": "failed", "error": "repair prompt declined; Word could not read it: open produced 0 new documents, none named 00011__wfv__411-identical-malformed-then-valid.docx; refusing to guess", "dialogs": ["'Word found unreadable content in 00011__wfv__411-identical-malformed-then-valid.docx. Do you want to recover the contents of this document? If you trust the source of this document, click Yes.'", "\"Word found unreadable content in 00011__wfv__411-identical-malformed-then-valid.docx. Do you want to recover the contents of this document? If you trust the source of this document, click Yes.\"", "\"Word experienced an error trying to open the file. / Try these suggestions. / * Check the file permissions for the document or drive. / * Make sure there is sufficient free memory and disk space. / * Open the file with the Text Recovery converter."]}
{"name": "411-dotdot-unreferenced-item", "sha256": "97f752f1b03e1749d1c8950d7f83aa2e58aaa9c76292603ee4e6c165621a774b", "observedAt": "2026-10-24T05:03:41.732Z", "outcome": "repair-warning", "item": "failed", "error": "repair prompt declined; Word could not read it: open produced 0 new documents, none named 00010__wfv__411-dotdot-unreferenced-item.docx; refusing to guess", "dialogs": ["'Word found unreadable content in 00010__wfv__411-dotdot-unreferenced-item.docx. Do you want to recover the contents of this document? If you trust the source of this document, click Yes.'", "\"Word found unreadable content in 00010__wfv__411-dotdot-unreferenced-item.docx. Do you want to recover the contents of this document? If you trust the source of this document, click Yes.\"", "\"Word experienced an error trying to open the file. / Try these suggestions. / * Check the file permissions for the document or drive. / * Make sure there is sufficient free memory and disk space. / * Open the file with the Text Recovery converter."]}
{"name": "411-dotdot-then-plain-styles", "sha256": "e68121f4cf1076570a08982b11976d1cdbbb590a55148fc6c887e702a7a49bad", "observedAt": "2026-10-24T05:03:31.782Z", "outcome": "repair-warning", "item": "failed", "error": "repair prompt declined; Word could not read it: open produced 0 new documents, none named 00009__wfv__411-dotdot-then-plain-styles.docx; refusing to guess", "dialogs": ["'Word found unreadable content in 00009__wfv__411-dotdot-then-plain-styles.docx. Do you want to recover the contents of this document? If you trust the source of this document, click Yes.'", "\"Word found unreadable content in 00009__wfv__411-dotdot-then-plain-styles.docx. Do you want to recover the contents of this document? If you trust the source of this document, click Yes.\"", "\"Word experienced an error trying to open the file. / Try these suggestions. / * Check the file permissions for the document or drive. / * Make sure there is sufficient free memory and disk space. / * Open the file with the Text Recovery converter."]}
{"name": "414-styles-override-removed-control", "sha256": "e0595fa4e2e464ec4a3288b474b728226f587d933ac90e416fa36c5b7df7b9df", "observedAt": "2026-10-24T05:09:17.692Z", "outcome": "repair-warning", "item": "failed", "error": "repair prompt declined; Word could not read it: open produced 0 new documents, none named 00022__wfv__414-styles-override-removed-control.docx; refusing to guess", "dialogs": ["\"Word found unreadable content in 00022__wfv__414-styles-override-removed-control.docx. Do you want to recover the contents of this document? If you trust the source of this document, click Yes.\""]}
{"name": "414-styles-override-dotdot", "sha256": "1fa751332e19b8f00e34ecf1007caaf479a50ba162fd324332c8c69645a50838", "observedAt": "2026-10-24T05:07:22.820Z", "outcome": "repair-warning", "item": "failed", "error": "repair prompt declined; Word could not read it: open produced 0 new documents, none named 00019__wfv__414-styles-override-dotdot.docx; refusing to guess", "dialogs": ["\"Word found unreadable content in 00019__wfv__414-styles-override-dotdot.docx. Do you want to recover the contents of this document? If you trust the source of this document, click Yes.\""]}
{"name": "414-styles-override-query", "sha256": "75810dc3e308696a365b4131c39152371d0eb7d4eff33ca2dfe31113994a7139", "observedAt": "2026-10-24T05:08:24.569Z", "outcome": "repair-warning", "item": "failed", "error": "repair prompt declined; Word could not read it: open produced 0 new documents, none named 00021__wfv__414-styles-override-query.docx; refusing to guess", "dialogs": ["'Word found unreadable content in 00021__wfv__414-styles-override-query.docx. Do you want to recover the contents of this document? If you trust the source of this document, click Yes.'", "\"Word found unreadable content in 00021__wfv__414-styles-override-query.docx. Do you want to recover the contents of this document? If you trust the source of this document, click Yes.\"", "\"Word experienced an error trying to open the file. / Try these suggestions. / * Check the file permissions for the document or drive. / * Make sure there is sufficient free memory and disk space. / * Open the file with the Text Recovery converter."]}
{"name": "414-styles-override-fragment", "sha256": "7a16f99036a62688c7426afc8b360a0a7bb19d6f915964b5b1fad024d1ea4753", "observedAt": "2026-10-24T05:13:33.732Z", "outcome": "repair-warning", "item": "failed", "error": "repair prompt declined; Word could not read it: open produced 0 new documents, none named 00020__wfv__414-styles-override-fragment.docx; refusing to guess", "dialogs": ["\"Word found unreadable content in 00020__wfv__414-styles-override-fragment.docx. Do you want to recover the contents of this document? If you trust the source of this document, click Yes.\""]}
{"name": "414-styles-override-backslash", "sha256": "4e5e85d63f2ee9c34982a3dbfd1d4253531e079d16357b2acbdbc0b666c97750", "observedAt": "2026-10-24T05:06:30.793Z", "outcome": "repair-warning", "item": "failed", "error": "repair prompt declined; Word could not read it: open produced 0 new documents, none named 00018__wfv__414-styles-override-backslash.docx; refusing to guess", "dialogs": ["'Word found unreadable content in 00018__wfv__414-styles-override-backslash.docx. Do you want to recover the contents of this document? If you trust the source of this document, click Yes.'", "\"Word found unreadable content in 00018__wfv__414-styles-override-backslash.docx. Do you want to recover the contents of this document? If you trust the source of this document, click Yes.\"", "\"Word experienced an error trying to open the file. / Try these suggestions. / * Check the file permissions for the document or drive. / * Make sure there is sufficient free memory and disk space. / * Open the file with the Text Recovery converter."]}
{"name": "414-invalid-x-dotdot-main", "sha256": "f7bae552bbd4fd8a29a6f758dd0f8200e5931ae8f95abd72f9f2028a11b5a996", "observedAt": "2026-10-24T05:04:59.679Z", "outcome": "repair-warning", "item": "failed", "error": "repair prompt declined; Word could not read it: open produced 0 new documents, none named 00015__wfv__414-invalid-x-dotdot-main.docx; refusing to guess", "dialogs": ["'Word found unreadable content in 00015__wfv__414-invalid-x-dotdot-main.docx. Do you want to recover the contents of this document? If you trust the source of this document, click Yes.'", "\"Word found unreadable content in 00015__wfv__414-invalid-x-dotdot-main.docx. Do you want to recover the contents of this document? If you trust the source of this document, click Yes.\"", "\"Word experienced an error trying to open the file. / Try these suggestions. / * Check the file permissions for the document or drive. / * Make sure there is sufficient free memory and disk space. / * Open the file with the Text Recovery converter."]}
{"name": "414-invalid-x-query-main", "sha256": "0ceb5371ecc1714b06bc9f670e23559bd5fcc3c977cac7e2ca4a42311a7f6ded", "observedAt": "2026-10-24T05:06:19.728Z", "outcome": "repair-warning", "item": "failed", "error": "repair prompt declined; Word could not read it", "dialogs": ["'Word found unreadable content in 00017__wfv__414-invalid-x-query-main.docx. Do you want to recover the contents of this document? If you trust the source of this document, click Yes.'", "\"Word found unreadable content in 00017__wfv__414-invalid-x-query-main.docx. Do you want to recover the contents of this document? If you trust the source of this document, click Yes.\"", "\"Word experienced an error trying to open the file. / Try these suggestions. / * Check the file permissions for the document or drive. / * Make sure there is sufficient free memory and disk space. / * Open the file with the Text Recovery converter."]}
{"name": "414-invalid-x-fragment-main", "sha256": "00fb33975b0027c9e8f1ee38a72ee026c73f330b751ba700b251873d62f22d72", "observedAt": "2026-10-24T05:05:46.744Z", "outcome": "repair-warning", "item": "failed", "error": "repair prompt declined; Word could not read it: open produced 0 new documents, none named 00016__wfv__414-invalid-x-fragment-main.docx; refusing to guess", "dialogs": ["'Word found unreadable content in 00016__wfv__414-invalid-x-fragment-main.docx. Do you want to recover the contents of this document? If you trust the source of this document, click Yes.'", "\"Word found unreadable content in 00016__wfv__414-invalid-x-fragment-main.docx. Do you want to recover the contents of this document? If you trust the source of this document, click Yes.\"", "\"Word experienced an error trying to open the file. / Try these suggestions. / * Check the file permissions for the document or drive. / * Make sure there is sufficient free memory and disk space. / * Open the file with the Text Recovery converter."]}
{"name": "414-invalid-x-backslash-main", "sha256": "f6451c46a832393860a5ea916b28de226ade0cc46cd47b68d7913b0a01b8a5a4", "observedAt": "2026-10-24T05:04:55.509Z", "outcome": "repair-warning", "item": "failed", "error": "repair prompt declined; Word could not read it", "dialogs": ["'Word found unreadable content in 00014__wfv__414-invalid-x-backslash-main.docx. Do you want to recover the contents of this document? If you trust the source of this document, click Yes.'", "\"Word found unreadable content in 00014__wfv__414-invalid-x-backslash-main.docx. Do you want to recover the contents of this document? If you trust the source of this document, click Yes.\"", "\"Word experienced an error trying to open the file. / Try these suggestions. / * Check the file permissions for the document or drive. / * Make sure there is sufficient free memory and disk space. / * Open the file with the Text Recovery converter."]}
{"name": "dupdecl-default-identical", "sha256": "222d7ee2fc43b1f6fa373efa5f8918f02a1a53a6234052e5945327c2be957873", "observedAt": "2026-10-24T05:10:00.151Z", "outcome": "repair-warning", "item": "failed", "error": "repair prompt declined; Word could not read it", "dialogs": ["'Word found unreadable content in 00035__wfv__dupdecl-default-identical.docx. Do you want to recover the contents of this document? If you trust the source of this document, click Yes.'", "\"Word found unreadable content in 00035__wfv__dupdecl-default-identical.docx. Do you want to recover the contents of this document? If you trust the source of this document, click Yes.\"", "\"Word experienced an error trying to open the file. / Try these suggestions. / * Check the file permissions for the document or drive. / * Make sure there is sufficient free memory and disk space. / * Open the file with the Text Recovery converter."]}
{"name": "dupdecl-default-case", "sha256": "fe596b4f8f32a55c1108ef7d3e0fcdca54bf5a5558ba249114f12b90fe2dc656", "observedAt": "2026-10-24T05:09:55.950Z", "outcome": "repair-warning", "item": "failed", "error": "repair prompt declined; Word could not read it", "dialogs": ["'Word found unreadable content in 00034__wfv__dupdecl-default-case.docx. Do you want to recover the contents of this document? If you trust the source of this document, click Yes.'", "\"Word found unreadable content in 00034__wfv__dupdecl-default-case.docx. Do you want to recover the contents of this document? If you trust the source of this document, click Yes.\"", "\"Word experienced an error trying to open the file. / Try these suggestions. / * Check the file permissions for the document or drive. / * Make sure there is sufficient free memory and disk space. / * Open the file with the Text Recovery converter."]}
{"name": "dupdecl-override-identical", "sha256": "92c289efb80c1b4171d3da146063eea94a8fa742b1ce979d98db6f0a9d01b110", "observedAt": "2026-10-24T05:10:12.787Z", "outcome": "repair-warning", "item": "failed", "error": "repair prompt declined; Word could not read it", "dialogs": ["'Word found unreadable content in 00038__wfv__dupdecl-override-identical.docx. Do you want to recover the contents of this document? If you trust the source of this document, click Yes.'", "\"Word found unreadable content in 00038__wfv__dupdecl-override-identical.docx. Do you want to recover the contents of this document? If you trust the source of this document, click Yes.\"", "\"Word experienced an error trying to open the file. / Try these suggestions. / * Check the file permissions for the document or drive. / * Make sure there is sufficient free memory and disk space. / * Open the file with the Text Recovery converter."]}
{"name": "dupdecl-override-conflicting", "sha256": "1fe83fb76ef60ec209e7148ecb8f45eb5bd395abbc35095374c19763163328e6", "observedAt": "2026-10-24T05:10:08.582Z", "outcome": "repair-warning", "item": "failed", "error": "repair prompt declined; Word could not read it", "dialogs": ["'Word found unreadable content in 00037__wfv__dupdecl-override-conflicting.docx. Do you want to recover the contents of this document? If you trust the source of this document, click Yes.'", "\"Word found unreadable content in 00037__wfv__dupdecl-override-conflicting.docx. Do you want to recover the contents of this document? If you trust the source of this document, click Yes.\"", "\"Word experienced an error trying to open the file. / Try these suggestions. / * Check the file permissions for the document or drive. / * Make sure there is sufficient free memory and disk space. / * Open the file with the Text Recovery converter."]}
{"name": "dupdecl-override-case", "sha256": "2f9f7277561b364b8ef0153456571f844b0eaf05dba51c9ca3f8346e67046aac", "observedAt": "2026-10-24T05:10:04.365Z", "outcome": "repair-warning", "item": "failed", "error": "repair prompt declined; Word could not read it", "dialogs": ["'Word found unreadable content in 00036__wfv__dupdecl-override-case.docx. Do you want to recover the contents of this document? If you trust the source of this document, click Yes.'", "\"Word found unreadable content in 00036__wfv__dupdecl-override-case.docx. Do you want to recover the contents of this document? If you trust the source of this document, click Yes.\"", "\"Word experienced an error trying to open the file. / Try these suggestions. / * Check the file permissions for the document or drive. / * Make sure there is sufficient free memory and disk space. / * Open the file with the Text Recovery converter."]}
{"name": "esc-plain-unreferenced-control", "sha256": "6e63642feef861ab8d78183d4d9fc19387179c4605ca42af6a25fb642b04ec41", "observedAt": "2026-10-24T05:18:09+00:00", "outcome": "clean-open", "item": "ok", "error": "", "dialogs": []}
{"name": "esc-%41-upper-letter", "sha256": "fae72ee8b638fa8a70e690ba666383a1bde0f9605b27c84996d99ac14106ec8e", "observedAt": "2026-10-24T05:10:34.149Z", "outcome": "repair-warning", "item": "failed", "error": "repair prompt declined; Word could not read it", "dialogs": ["'Word found unreadable content in 00044__wfv__esc-pct41-upper-letter.docx. Do you want to recover the contents of this document? If you trust the source of this document, click Yes.'", "\"Word found unreadable content in 00044__wfv__esc-pct41-upper-letter.docx. Do you want to recover the contents of this document? If you trust the source of this document, click Yes.\"", "\"Word experienced an error trying to open the file. / Try these suggestions. / * Check the file permissions for the document or drive. / * Make sure there is sufficient free memory and disk space. / * Open the file with the Text Recovery converter."]}
{"name": "esc-%30-digit", "sha256": "17d763961bc3f58c062f92c6a00005b8bee15200a3f6ce251d3310d5db7dc134", "observedAt": "2026-10-24T05:10:29.952Z", "outcome": "repair-warning", "item": "failed", "error": "repair prompt declined; Word could not read it", "dialogs": ["'Word found unreadable content in 00043__wfv__esc-pct30-digit.docx. Do you want to recover the contents of this document? If you trust the source of this document, click Yes.'", "\"Word found unreadable content in 00043__wfv__esc-pct30-digit.docx. Do you want to recover the contents of this document? If you trust the source of this document, click Yes.\"", "\"Word experienced an error trying to open the file. / Try these suggestions. / * Check the file permissions for the document or drive. / * Make sure there is sufficient free memory and disk space. / * Open the file with the Text Recovery converter."]}
{"name": "esc-%2D-hyphen", "sha256": "84a0e0f1843139fd42dfcf81ce7f84f5f2882d445cdb4814c6883516b30d99a5", "observedAt": "2026-10-24T05:10:17.302Z", "outcome": "repair-warning", "item": "failed", "error": "repair prompt declined; Word could not read it", "dialogs": ["'Word found unreadable content in 00040__wfv__esc-pct2D-hyphen.docx. Do you want to recover the contents of this document? If you trust the source of this document, click Yes.'", "\"Word found unreadable content in 00040__wfv__esc-pct2D-hyphen.docx. Do you want to recover the contents of this document? If you trust the source of this document, click Yes.\"", "\"Word experienced an error trying to open the file. / Try these suggestions. / * Check the file permissions for the document or drive. / * Make sure there is sufficient free memory and disk space. / * Open the file with the Text Recovery converter."]}
{"name": "esc-%2d-hyphen-lowerhex", "sha256": "37529d0e79b352aa3d1b870d9157b18b437c83e640d21cfbf001ce01c0a71f98", "observedAt": "2026-10-24T05:10:25.767Z", "outcome": "repair-warning", "item": "failed", "error": "repair prompt declined; Word could not read it", "dialogs": ["'Word found unreadable content in 00042__wfv__esc-pct2d-hyphen-lowerhex.docx. Do you want to recover the contents of this document? If you trust the source of this document, click Yes.'", "\"Word found unreadable content in 00042__wfv__esc-pct2d-hyphen-lowerhex.docx. Do you want to recover the contents of this document? If you trust the source of this document, click Yes.\"", "\"Word experienced an error trying to open the file. / Try these suggestions. / * Check the file permissions for the document or drive. / * Make sure there is sufficient free memory and disk space. / * Open the file with the Text Recovery converter."]}
{"name": "esc-%2E-period", "sha256": "fcbdbbf79faa0bcd6da0f0534059fdc0797cb4314c7f44eb9493499de0c13d9f", "observedAt": "2026-10-24T05:10:21.519Z", "outcome": "repair-warning", "item": "failed", "error": "repair prompt declined; Word could not read it", "dialogs": ["'Word found unreadable content in 00041__wfv__esc-pct2E-period.docx. Do you want to recover the contents of this document? If you trust the source of this document, click Yes.'", "\"Word found unreadable content in 00041__wfv__esc-pct2E-period.docx. Do you want to recover the contents of this document? If you trust the source of this document, click Yes.\"", "\"Word experienced an error trying to open the file. / Try these suggestions. / * Check the file permissions for the document or drive. / * Make sure there is sufficient free memory and disk space. / * Open the file with the Text Recovery converter."]}
{"name": "esc-%5F-underscore", "sha256": "7d0278c4bf9fd3fa41dadb456cd2e31091f92135e4bfc678bb6f137646da52de", "observedAt": "2026-10-24T05:10:38.399Z", "outcome": "repair-warning", "item": "failed", "error": "repair prompt declined; Word could not read it: open produced 0 new documents, none named 00045__wfv__esc-pct5F-underscore.docx; refusing to guess", "dialogs": ["'Word found unreadable content in 00045__wfv__esc-pct5F-underscore.docx. Do you want to recover the contents of this document? If you trust the source of this document, click Yes.'", "\"Word found unreadable content in 00045__wfv__esc-pct5F-underscore.docx. Do you want to recover the contents of this document? If you trust the source of this document, click Yes.\"", "\"Word experienced an error trying to open the file. / Try these suggestions. / * Check the file permissions for the document or drive. / * Make sure there is sufficient free memory and disk space. / * Open the file with the Text Recovery converter."]}
{"name": "esc-%7E-tilde", "sha256": "4436f009f2500d574870504a3c6abdffab59ea97bc5d50416f7784a0ff1c847c", "observedAt": "2026-10-24T05:11:33.011Z", "outcome": "repair-warning", "item": "failed", "error": "repair prompt declined; Word could not read it: open produced 0 new documents, none named 00046__wfv__esc-pct7E-tilde.docx; refusing to guess", "dialogs": ["'Word found unreadable content in 00046__wfv__esc-pct7E-tilde.docx. Do you want to recover the contents of this document? If you trust the source of this document, click Yes.'", "\"Word found unreadable content in 00046__wfv__esc-pct7E-tilde.docx. Do you want to recover the contents of this document? If you trust the source of this document, click Yes.\"", "\"Word experienced an error trying to open the file. / Try these suggestions. / * Check the file permissions for the document or drive. / * Make sure there is sufficient free memory and disk space. / * Open the file with the Text Recovery converter."]}
{"name": "esc-%20-space-control", "sha256": "c7598087941b0c35b84765110d5289c53c86b0675fcdfe31db21ba23def5d1f9", "observedAt": "2026-10-24T05:18:09+00:00", "outcome": "clean-open", "item": "ok", "error": "", "dialogs": []}
{"name": "s3-control-minimal-end", "sha256": "00d28798c1f66dc7dfd1f484e065d729fe95fdc34a4491be031f3750b59ef3d6", "observedAt": "2026-10-24T05:18:09+00:00", "outcome": "clean-open", "item": "ok", "error": "", "dialogs": []}
{"name": "rels-main-nested-in-unknown-element", "sha256": "176e6d8f133812b3de0d01186fcf6f950b700a993d5b6cec7ec3c523b65ececc", "observedAt": "2026-10-24T05:11:58.058Z", "outcome": "repair-warning", "item": "failed", "error": "repair prompt declined; Word could not read it", "dialogs": ["\"Word found unreadable content in 00049__wfv__rels-main-nested-in-unknown-element.docx. Do you want to recover the contents of this document? If you trust the source of this document, click Yes.\""]}
{"name": "rels-root-not-relationships", "sha256": "1a5c659713cf7f9b7ba3aef608df5c3ce6e583de2747c4b58a06437f53ef52e7", "observedAt": "2026-10-24T05:12:38.604Z", "outcome": "repair-warning", "item": "failed", "error": "repair prompt declined; Word could not read it: open produced 0 new documents, none named 00050__wfv__rels-root-not-relationships.docx; refusing to guess", "dialogs": ["'Word found unreadable content in 00050__wfv__rels-root-not-relationships.docx. Do you want to recover the contents of this document? If you trust the source of this document, click Yes.'", "\"Word found unreadable content in 00050__wfv__rels-root-not-relationships.docx. Do you want to recover the contents of this document? If you trust the source of this document, click Yes.\"", "\"Word experienced an error trying to open the file. / Try these suggestions. / * Check the file permissions for the document or drive. / * Make sure there is sufficient free memory and disk space. / * Open the file with the Text Recovery converter."]}
{"name": "rels-unknown-sibling-control", "sha256": "27cb9b0032c7b10c88fc12f5eb47a7b8a9b8845b30ce867f3bd9d33725f98241", "observedAt": "2026-10-24T05:18:09+00:00", "outcome": "other", "item": "failed", "error": "Word could not load it in 2 passes: open produced 1 new documents, none named 00051__wfv__rels-unknown-sibling-control.docx; refusing to guess", "dialogs": []}
{"name": "s3c-control-minimal", "sha256": "1a414a1109291a1823ea0a0bb5be8cd5659634f1614aa961e3bc0cf239cad4a1", "observedAt": "2026-10-24T05:26:57+00:00", "outcome": "clean-open", "item": "ok", "error": "", "dialogs": []}
{"name": "ct-stream-missing-control", "sha256": "a4c6ecaf2f422b09f82284c73c9f288b2488c67bf00e880a82394edc15a012e9", "observedAt": "2026-10-24T05:22:07.264Z", "outcome": "repair-warning", "item": "failed", "error": "Word could not load it in 2 passes: open produced 0 new documents, none named 00007__wfv__ct-stream-missing-control.docx; refusing to guess", "dialogs": ["'Word experienced an error trying to open the file.\\nTry these suggestions.\\n* Check the file permissions for the document or drive.\\n* Make sure there is sufficient free memory and disk space.\\n* Open the file with the Text Recovery conver", "\"Word experienced an error trying to open the file. / Try these suggestions. / * Check the file permissions for the document or drive. / * Make sure there is sufficient free memory and disk space. / * Open the file with the Text Recovery converter.", "'Word found unreadable content in 00007__wfv__ct-stream-missing-control.docx. Do you want to recover the contents of this document? If you trust the source of this document, click Yes.'", "\"Word found unreadable content in 00007__wfv__ct-stream-missing-control.docx. Do you want to recover the contents of this document? If you trust the source of this document, click Yes.\""]}
{"name": "ct-stream-escaped-only", "sha256": "42d70cff0d0d0e1a44ea46f743cd749ff62cc6b527736670c26e6a878f23acf0", "observedAt": "2026-10-24T05:21:51.014Z", "outcome": "repair-warning", "item": "failed", "error": "repair prompt declined; Word could not read it: open produced 0 new documents, none named 00006__wfv__ct-stream-escaped-only.docx; refusing to guess", "dialogs": ["'Word found unreadable content in 00006__wfv__ct-stream-escaped-only.docx. Do you want to recover the contents of this document? If you trust the source of this document, click Yes.'", "\"Word found unreadable content in 00006__wfv__ct-stream-escaped-only.docx. Do you want to recover the contents of this document? If you trust the source of this document, click Yes.\""]}
{"name": "ct-stream-escaped-extra", "sha256": "898d49c3184b8458d67dcb5a880191c4fcae67be06bfb36d89eda6579391f211", "observedAt": "2026-10-24T05:21:03.646Z", "outcome": "repair-warning", "item": "failed", "error": "repair prompt declined; Word could not read it: open produced 0 new documents, none named 00005__wfv__ct-stream-escaped-extra.docx; refusing to guess", "dialogs": ["'Word found unreadable content in 00005__wfv__ct-stream-escaped-extra.docx. Do you want to recover the contents of this document? If you trust the source of this document, click Yes.'", "\"Word found unreadable content in 00005__wfv__ct-stream-escaped-extra.docx. Do you want to recover the contents of this document? If you trust the source of this document, click Yes.\"", "\"Word experienced an error trying to open the file. / Try these suggestions. / * Check the file permissions for the document or drive. / * Make sure there is sufficient free memory and disk space. / * Open the file with the Text Recovery converter."]}
{"name": "esc-pct2F-slash-unreferenced", "sha256": "1003d9e0b112fcabc3d0f00533e7080bcf497d8d54749a7cd3627efd980e446d", "observedAt": "2026-10-24T05:23:30.476Z", "outcome": "repair-warning", "item": "failed", "error": "repair prompt declined; Word could not read it: open produced 0 new documents, none named 00015__wfv__esc-pct2F-slash-unreferenced.docx; refusing to guess", "dialogs": ["\"Word found unreadable content in 00015__wfv__esc-pct2F-slash-unreferenced.docx. Do you want to recover the contents of this document? If you trust the source of this document, click Yes.\""]}
{"name": "esc-pct5C-backslash-unreferenced", "sha256": "76f9efc51af63338b83fe1cfcc23098aaaeb5f8f6bfae30f6315b7f84ebf535a", "observedAt": "2026-10-24T05:25:16.693Z", "outcome": "open-error", "item": "failed", "error": "Word could not load it in 2 passes: open produced 0 new documents, none named 00016__wfv__esc-pct5C-backslash-unreferenced.docx; refusing to guess", "dialogs": ["'Word experienced an error trying to open the file.\\nTry these suggestions.\\n* Check the file permissions for the document or drive.\\n* Make sure there is sufficient free memory and disk space.\\n* Open the file with the Text Recovery conver", "\"Word experienced an error trying to open the file. / Try these suggestions. / * Check the file permissions for the document or drive. / * Make sure there is sufficient free memory and disk space. / * Open the file with the Text Recovery converter."]}
{"name": "docprops-custom-unreferenced-undeclared-prefix", "sha256": "0a49797ea81f27d6dfad5601bbd0cf1fc6e3b8ccafbd66d124d59b5e2663b548", "observedAt": "2026-10-24T05:26:57+00:00", "outcome": "clean-open", "item": "ok", "error": "", "dialogs": []}
{"name": "docprops-custom-unreferenced-truncated", "sha256": "050e2698f564c7613ff607cbb7f0e3ff8d0bd276d2804895b25112f4d28a906b", "observedAt": "2026-10-24T05:26:57+00:00", "outcome": "clean-open", "item": "ok", "error": "", "dialogs": []}
{"name": "docprops-app-unreferenced-truncated", "sha256": "57b78d2872e4569a66fc96367672372a0a63e17c36ad6368871d4c81c158d0bb", "observedAt": "2026-10-24T05:26:57+00:00", "outcome": "clean-open", "item": "ok", "error": "", "dialogs": []}
{"name": "word-unreferenced-malformed-part", "sha256": "8968d3be05c61519b5636bb63aa5fb1107897af6aa3dd4c98d9a282fb462bd31", "observedAt": "2026-10-24T05:26:57+00:00", "outcome": "clean-open", "item": "ok", "error": "", "dialogs": []}
{"name": "word-unreferenced-wellformed-control", "sha256": "7b8e0e0d4ae48b381266fab8968c6c24697dd6072b94a84aab928be96bb379ea", "observedAt": "2026-10-24T05:26:57+00:00", "outcome": "clean-open", "item": "ok", "error": "", "dialogs": []}
{"name": "docprops-core-unreferenced-truncated", "sha256": "cefef105ee09a3c0dee534f9cf65b12214805ec81b9c1689b44baf4ef61a5e2a", "observedAt": "2026-10-24T05:26:57+00:00", "outcome": "clean-open", "item": "ok", "error": "", "dialogs": []}
{"name": "docprops-core-renamed-truncated", "sha256": "e0290e52e694e983777f6affaca16d6152ca7471fe3617031a673c713a799522", "observedAt": "2026-10-24T05:22:11.750Z", "outcome": "repair-warning", "item": "failed", "error": "repair prompt declined; Word could not read it: open produced 0 new documents, none named 00010__wfv__docprops-core-renamed-truncated.docx; refusing to guess", "dialogs": ["'Word found unreadable content in 00010__wfv__docprops-core-renamed-truncated.docx. Do you want to recover the contents of this document? If you trust the source of this document, click Yes.'", "\"Word found unreadable content in 00010__wfv__docprops-core-renamed-truncated.docx. Do you want to recover the contents of this document? If you trust the source of this document, click Yes.\"", "\"Word experienced an error trying to open the file. / Try these suggestions. / * Check the file permissions for the document or drive. / * Make sure there is sufficient free memory and disk space. / * Open the file with the Text Recovery converter."]}
{"name": "docprops-core-renamed-control", "sha256": "de166f9b02fa683d138d0eb981a182681e37bcceba73d77feebfb6215051f65e", "observedAt": "2026-10-24T05:26:57+00:00", "outcome": "clean-open", "item": "ok", "error": "", "dialogs": []}
{"name": "empty-archive", "sha256": "8739c76e681f900923b900c9df0ef75cf421d39cabb54650c4b9ad19b6a76d85", "observedAt": "2026-10-24T05:22:33.454Z", "outcome": "repair-warning", "item": "failed", "error": "repair prompt declined; Word could not read it: open produced 0 new documents, none named 00014__wfv__empty-archive.docx; refusing to guess", "dialogs": ["\"Word found unreadable content in 00014__wfv__empty-archive.docx. Do you want to recover the contents of this document? If you trust the source of this document, click Yes.\""]}
{"name": "main-part-missing", "sha256": "035e63734f8226d0c0dc5022ac68287e13f24b82d78b0e86ff775faa5e97ae8a", "observedAt": "2026-10-24T05:25:33.299Z", "outcome": "repair-warning", "item": "failed", "error": "repair prompt declined; Word could not read it: open produced 0 new documents, none named 00017__wfv__main-part-missing.docx; refusing to guess", "dialogs": ["'Word found unreadable content in 00017__wfv__main-part-missing.docx. Do you want to recover the contents of this document? If you trust the source of this document, click Yes.'", "\"Word found unreadable content in 00017__wfv__main-part-missing.docx. Do you want to recover the contents of this document? If you trust the source of this document, click Yes.\"", "\"Word experienced an error trying to open the file. / Try these suggestions. / * Check the file permissions for the document or drive. / * Make sure there is sufficient free memory and disk space. / * Open the file with the Text Recovery converter."]}
{"name": "main-relationship-missing", "sha256": "6cff8f9d94dc4883be54782e873a345840856291f15a33fb6f31d3c351917890", "observedAt": "2026-10-24T05:26:57+00:00", "outcome": "other", "item": "failed", "error": "Word could not load it in 2 passes: open produced 0 new documents, none named 00018__wfv__main-relationship-missing.docx; refusing to guess", "dialogs": []}
{"name": "414-invalid-x-plain-missing-control", "sha256": "dc5dafa7e6a85c49308d2a5571e5d8042fff7f6b6d528a2959a127b55bc175ae", "observedAt": "2026-10-24T05:26:57+00:00", "outcome": "clean-open", "item": "ok", "error": "", "dialogs": []}
{"name": "414-invalid-x-dotdot-missing", "sha256": "2bd531b03fb63a740f04a71d312914e1ab7d4eb01bd7e0ca88bb74a5bf48b9de", "observedAt": "2026-10-24T05:26:52.613Z", "outcome": "repair-warning", "item": "failed", "error": "repair prompt declined; Word could not read it", "dialogs": ["'Word found unreadable content in 00001__wfv__414-invalid-x-dotdot-missing.docx. Do you want to recover the contents of this document? If you trust the source of this document, click Yes.'", "\"Word found unreadable content in 00001__wfv__414-invalid-x-dotdot-missing.docx. Do you want to recover the contents of this document? If you trust the source of this document, click Yes.\"", "\"Word experienced an error trying to open the file. / Try these suggestions. / * Check the file permissions for the document or drive. / * Make sure there is sufficient free memory and disk space. / * Open the file with the Text Recovery converter."]}
{"name": "414-invalid-x-query-missing", "sha256": "8c41377382c63da1d8dd90a9d730620e896f220b3f083083618d3e93c7d3f5e1", "observedAt": "2026-10-24T05:20:59.363Z", "outcome": "open-error", "item": "failed", "error": "Word could not load it in 2 passes: open produced 0 new documents, none named 00004__wfv__414-invalid-x-query-missing.docx; refusing to guess", "dialogs": ["'Word experienced an error trying to open the file.\\nTry these suggestions.\\n* Check the file permissions for the document or drive.\\n* Make sure there is sufficient free memory and disk space.\\n* Open the file with the Text Recovery conver", "\"Word experienced an error trying to open the file. / Try these suggestions. / * Check the file permissions for the document or drive. / * Make sure there is sufficient free memory and disk space. / * Open the file with the Text Recovery converter."]}
{"name": "414-invalid-x-fragment-missing", "sha256": "ae39b41b00a82878ad7785ad3b46ba8aac1351d748316ecc5a15092d36a5ac61", "observedAt": "2026-10-24T05:20:39.296Z", "outcome": "repair-warning", "item": "failed", "error": "Word could not load it in 2 passes: open produced 0 new documents, none named 00002__wfv__414-invalid-x-fragment-missing.docx; refusing to guess", "dialogs": ["'Word found unreadable content in 00002__wfv__414-invalid-x-fragment-missing.docx. Do you want to recover the contents of this document? If you trust the source of this document, click Yes.'", "\"Word found unreadable content in 00002__wfv__414-invalid-x-fragment-missing.docx. Do you want to recover the contents of this document? If you trust the source of this document, click Yes.\"", "\"Word experienced an error trying to open the file. / Try these suggestions. / * Check the file permissions for the document or drive. / * Make sure there is sufficient free memory and disk space. / * Open the file with the Text Recovery converter.", "'Word experienced an error trying to open the file.\\nTry these suggestions.\\n* Check the file permissions for the document or drive.\\n* Make sure there is sufficient free memory and disk space.\\n* Open the file with the Text Recovery conver"]}
{"name": "414-invalid-x-backslash-missing", "sha256": "8fc0c3015afe29b28a7acbb5e14da9195b110d9220f330cb22d6143788c81e00", "observedAt": "2026-10-24T05:19:49.448Z", "outcome": "repair-warning", "item": "failed", "error": "repair prompt declined; Word could not read it: open produced 0 new documents, none named 00000__wfv__414-invalid-x-backslash-missing.docx; refusing to guess", "dialogs": ["'Word found unreadable content in 00000__wfv__414-invalid-x-backslash-missing.docx. Do you want to recover the contents of this document? If you trust the source of this document, click Yes.'", "\"Word found unreadable content in 00000__wfv__414-invalid-x-backslash-missing.docx. Do you want to recover the contents of this document? If you trust the source of this document, click Yes.\"", "\"Word experienced an error trying to open the file. / Try these suggestions. / * Check the file permissions for the document or drive. / * Make sure there is sufficient free memory and disk space. / * Open the file with the Text Recovery converter."]}
{"name": "rels-unknown-sibling-control", "sha256": "43205d2359aaac364bd348acd4205514d8572e334162a08ed113d02fa9eb826b", "observedAt": "2026-10-24T05:26:22.486Z", "outcome": "repair-warning", "item": "failed", "error": "repair prompt declined; Word could not read it", "dialogs": ["'Word found unreadable content in 00019__wfv__rels-unknown-sibling-control.docx. Do you want to recover the contents of this document? If you trust the source of this document, click Yes.'", "\"Word found unreadable content in 00019__wfv__rels-unknown-sibling-control.docx. Do you want to recover the contents of this document? If you trust the source of this document, click Yes.\"", "\"Word experienced an error trying to open the file. / Try these suggestions. / * Check the file permissions for the document or drive. / * Make sure there is sufficient free memory and disk space. / * Open the file with the Text Recovery converter."]}
```

</details>

<details><summary>Probe script</summary>

Run from the repository root as a copy: `node --import tsx ./word-probe.tmp.mts [specimen …]`.

```ts
// Real-Word probes for #388, #392, #396, #397, #351, #398. Run from the stack worktree root as a copy.
import { createHash } from "node:crypto";
import { execSync } from "node:child_process";
import { appendFileSync, existsSync, readFileSync, statSync, writeFileSync } from "node:fs";
import path from "node:path";
import JSZip from "jszip";
import { MINIMAL_PARTS, buildDocx } from "./tests/support/minimal-docx";
import { validate } from "./src/scripts/office/validate";
import { verifyInWord } from "./src/lib/word-verifier";

const OUT = "/Users/arthrod/temp/T/wfv-scratch/word-probes";
const RESULTS = path.join(OUT, "results.jsonl");
// Interference guard: another session drives Word through word_pdf.py and logs to REDLINE_LOG.
const REDLINE_LOG = "/Users/arthrod/temp/T/fix013w/oracle/redline.log";
const INCONCLUSIVE = path.join(OUT, "results.inconclusive.jsonl");
const logSize = () => {
    try {
        return statSync(REDLINE_LOG).size;
    } catch {
        return -1;
    }
};
const foreignWordDrivers = () => {
    // Match the executable (ucomm), not command-line text: agents' prompts and shells mention osascript.
    const rows = execSync("ps -Ao pid=,ppid=,ucomm=")
        .toString()
        .split("\n")
        .map((l) => l.trim().match(/^(\d+)\s+(\d+)\s+(.*)$/))
        .filter((m): m is RegExpMatchArray => m !== null)
        .map((m) => ({ pid: Number(m[1]), ppid: Number(m[2]), exe: m[3] }));
    const cmds = new Map(
        execSync("ps -Ao pid=,command=")
            .toString()
            .split("\n")
            .map((l) => l.trim().match(/^(\d+)\s+(.*)$/))
            .filter((m): m is RegExpMatchArray => m !== null)
            .map((m) => [Number(m[1]), m[2]]),
    );
    const parent = new Map(rows.map((r) => [r.pid, r.ppid]));
    const mine = (pid: number) => {
        for (let p = pid, i = 0; p > 1 && i < 64; p = parent.get(p) ?? 1, i++) if (p === process.pid) return true;
        return false;
    };
    const driver = (r: { pid: number; exe: string }) =>
        r.exe === "osascript" || (/^python/i.test(r.exe) && /word_pdf|word_redline/.test(cmds.get(r.pid) ?? ""));
    return rows.filter((r) => driver(r) && !mine(r.pid)).map((r) => `${r.pid} ${r.exe} ${(cmds.get(r.pid) ?? "").slice(0, 80)}`);
};
const done = new Set(
    existsSync(RESULTS)
        ? readFileSync(RESULTS, "utf8")
              .trim()
              .split("\n")
              .filter(Boolean)
              .map((l) => JSON.parse(l).name)
        : [],
);

const P = MINIMAL_PARTS as Record<string, string>;
const CT = P["[Content_Types].xml"];
const without = (o: Record<string, string | Uint8Array>, k: string) => {
    const { [k]: _, ...r } = o;
    return r;
};
const upperWord = (parts: Record<string, string | Uint8Array>) =>
    Object.fromEntries(Object.entries(parts).map(([k, v]) => [k.startsWith("word/") ? "WORD/" + k.slice(5) : k, v]));
const invalidCt = CT.replace("</Types>", `<Override PartName="/word/x.xml" ContentType="invalid-type"/></Types>`);
const rawZip = async (entries: [string, string | Uint8Array][]) => {
    const z = new JSZip();
    for (const [k, v] of entries) z.file(k, v);
    return z.generateAsync({ type: "nodebuffer", compression: "DEFLATE" });
};
const renameInFixture = async (fixture: string, from: string, to: string) => {
    const z = await JSZip.loadAsync(readFileSync(fixture));
    const out = new JSZip();
    for (const [name, f] of Object.entries(z.files)) {
        if (f.dir) continue;
        out.file(name === from ? to : name, await f.async("uint8array"));
    }
    return out.generateAsync({ type: "nodebuffer", compression: "DEFLATE" });
};
const DUP_THEME = "tests/fixtures/broken/package.word-duplicate-theme.docx";
const badDoc = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:document`;
const enc = (e: string) => P["word/document.xml"].replace('encoding="UTF-8"', `encoding="${e}"`);

const R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const ctSwap = (from: string, to: string) => CT.replace(from, to);
const STYLES_CT = "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml";
const MAIN_CT = "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml";
const CORE_CT = "application/vnd.openxmlformats-package.core-properties+xml";
const CORE_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/"><dc:title>t</dc:title></cp:coreProperties>`;
const APP_CT = "application/vnd.openxmlformats-officedocument.extended-properties+xml";
const APP_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"><Application>Microsoft Office Word</Application></Properties>`;
const withApp = (appXml: string) => ({
    ...P,
    "docProps/app.xml": appXml,
    "_rels/.rels": P["_rels/.rels"].replace(
        "</Relationships>",
        `<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/></Relationships>`,
    ),
    "[Content_Types].xml": CT.replace("</Types>", `<Override PartName="/docProps/app.xml" ContentType="${APP_CT}"/></Types>`),
});
const withCore = (coreType: string) => ({
    ...P,
    "docProps/core.xml": CORE_XML,
    "_rels/.rels": P["_rels/.rels"].replace(
        "</Relationships>",
        `<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/></Relationships>`,
    ),
    "[Content_Types].xml": CT.replace("</Types>", `<Override PartName="/docProps/core.xml" ContentType="${coreType}"/></Types>`),
});
const docH = P["word/document.xml"]
    .replace("<w:document ", `<w:document xmlns:r="${R_NS}" `)
    .replace("<w:r><w:t>Hello</w:t></w:r>", `<w:hyperlink r:id="rId2"><w:r><w:t>Hello</w:t></w:r></w:hyperlink>`);
const relsH = P["word/_rels/document.xml.rels"].replace(
    "</Relationships>",
    `<Relationship Id="rId2" Type="${R_NS}/hyperlink" Target="https://example.com/" TargetMode="External"/></Relationships>`,
);
const renamed = (from: string, to: string, parts: Record<string, string | Uint8Array> = P) =>
    Object.fromEntries(Object.entries(parts).map(([k, v]) => [k === from ? to : k, v]));

const specimens: [string, string, () => Promise<Buffer>][] = [
    ["control-minimal", "control", () => buildDocx(P)],
    // #388
    ["388-rels-zero-byte", "#388", () => buildDocx({ ...P, "_rels/.rels": "" })],
    ["388-rels-truncated", "#388", () => buildDocx({ ...P, "_rels/.rels": P["_rels/.rels"].slice(0, 120) })],
    [
        "388-rels-undeclared-prefix",
        "#388",
        () =>
            buildDocx({
                ...P,
                "_rels/.rels": P["_rels/.rels"]
                    .replace("<Relationships ", "<x:Relationships ")
                    .replace("</Relationships>", "</x:Relationships>"),
            }),
    ],
    // #392
    ["392-WORD-folder-clean", "#392", () => buildDocx(upperWord(P))],
    ["392-WORD-folder-malformed-main", "#392", () => buildDocx(upperWord({ ...P, "word/document.xml": badDoc }))],
    ["392-word-folder-malformed-main-control", "#392", () => buildDocx({ ...P, "word/document.xml": badDoc })],
    // #396
    ["396-CT-upper-clean", "#396", () => buildDocx({ ...without(P, "[Content_Types].xml"), "[CONTENT_TYPES].XML": CT })],
    ["396-CT-upper-invalid-type", "#396", () => buildDocx({ ...without(P, "[Content_Types].xml"), "[CONTENT_TYPES].XML": invalidCt })],
    ["396-CT-exact-invalid-type-control", "#396", () => buildDocx({ ...P, "[Content_Types].xml": invalidCt })],
    ["396-CT-upper-malformed", "#396", () => buildDocx({ ...without(P, "[Content_Types].xml"), "[CONTENT_TYPES].XML": CT.slice(0, 150) })],
    // #397
    ["397-encoding-unknown", "#397", () => buildDocx({ ...P, "word/document.xml": enc("X-UNKNOWN-ENCODING") })],
    ["397-encoding-windows-1252-control", "#397", () => buildDocx({ ...P, "word/document.xml": enc("windows-1252") })],
    // #351
    ["351-dup-case-identical", "#351", () => rawZip([...Object.entries(P), ["WORD/document.xml", P["word/document.xml"]]])],
    ["351-dup-case-lower-good-upper-bad", "#351", () => rawZip([...Object.entries(P), ["WORD/document.xml", badDoc]])],
    ["351-dup-case-upper-bad-first", "#351", () => rawZip([["WORD/document.xml", badDoc], ...Object.entries(P)])],
    ["351-dup-stream-good-then-bad", "#351", () => rawZip([...Object.entries(P), ["[CONTENT_TYPES].XML", invalidCt]])],
    ["351-dup-stream-bad-then-good", "#351", () => rawZip([["[CONTENT_TYPES].XML", invalidCt], ...Object.entries(P)])],
    // #398
    [
        "398-RELS-sidecar-clean",
        "#398",
        () =>
            buildDocx({ ...without(P, "word/_rels/document.xml.rels"), "word/_rels/document.xml.RELS": P["word/_rels/document.xml.rels"] }),
    ],
    ["398-dup-theme-control", "#398", async () => readFileSync(DUP_THEME)],
    ["398-dup-theme-RELS", "#398", () => renameInFixture(DUP_THEME, "word/_rels/document.xml.rels", "word/_rels/document.xml.RELS")],
    // word-content-type-invalid scope (Override for an existing part vs a missing one; unused Default)
    [
        "cti-override-existing-styles",
        "cti",
        () => buildDocx({ ...P, "[Content_Types].xml": ctSwap(`ContentType="${STYLES_CT}"`, `ContentType="invalid-${STYLES_CT}"`) }),
    ],
    [
        "cti-override-existing-main",
        "cti",
        () => buildDocx({ ...P, "[Content_Types].xml": ctSwap(`ContentType="${MAIN_CT}"`, `ContentType="invalid-${MAIN_CT}"`) }),
    ],
    ["cti-override-existing-core", "cti", () => buildDocx(withCore(`invalid-${CORE_CT}`))],
    ["cti-override-existing-core-control", "cti", () => buildDocx(withCore(CORE_CT))],
    [
        "cti-default-unused-extension",
        "cti",
        () =>
            buildDocx({
                ...P,
                "[Content_Types].xml": CT.replace(
                    '<Default Extension="rels"',
                    `<Default Extension="foo" ContentType="invalid-type"/><Default Extension="rels"`,
                ),
            }),
    ],
    // #351 escape / backslash duplicates, and the single-spelling controls
    ["351-dup-escape-identical", "#351", () => rawZip([...Object.entries(P), ["word/%73tyles.xml", P["word/styles.xml"]]])],
    ["351-dup-backslash-identical", "#351", () => rawZip([...Object.entries(P), ["word\\document.xml", P["word/document.xml"]]])],
    ["351-escape-only-styles", "#351", () => rawZip(Object.entries(renamed("word/styles.xml", "word/%73tyles.xml")))],
    ["351-backslash-only-main", "#351", () => rawZip(Object.entries(renamed("word/document.xml", "word\\document.xml")))],
    // #398 follow-up: hyperlink resolution through a .RELS sidecar
    ["398-hyperlink-rels-control", "#398", () => buildDocx({ ...P, "word/document.xml": docH, "word/_rels/document.xml.rels": relsH })],
    [
        "398-hyperlink-%72els",
        "#402",
        () =>
            buildDocx({
                ...without({ ...P, "word/document.xml": docH }, "word/_rels/document.xml.rels"),
                "word/_rels/document.xml.%72els": relsH,
            }),
    ],
    [
        "398-hyperlink-RELS",
        "#398",
        () =>
            buildDocx({
                ...without({ ...P, "word/document.xml": docH }, "word/_rels/document.xml.rels"),
                "word/_rels/document.xml.RELS": relsH,
            }),
    ],
    // docProps scope for xml-syntax / xml-encoding (#397 follow-up)
    ["docprops-core-truncated", "docprops", () => buildDocx({ ...withCore(CORE_CT), "docProps/core.xml": CORE_XML.slice(0, 140) })],
    ["docprops-app-control", "docprops", () => buildDocx(withApp(APP_XML))],
    ["docprops-app-truncated", "docprops", () => buildDocx(withApp(APP_XML.slice(0, 140)))],
    [
        "docprops-app-encoding-unknown",
        "docprops",
        () => buildDocx(withApp(APP_XML.replace('encoding="UTF-8"', 'encoding="X-UNKNOWN-ENCODING"'))),
    ],
    // Escaped ZIP item names: does Word decode %XX of an unreserved character, or refuse the name itself?
    ["esc-unreferenced-extra-item", "esc", () => rawZip([...Object.entries(P), ["word/%78tra.xml", P["word/styles.xml"]]])],
    ["esc-styles-missing-control", "esc", () => rawZip(Object.entries(without(P, "word/styles.xml")))],
    ["esc-only-main", "esc", () => rawZip(Object.entries(renamed("word/document.xml", "word/%64ocument.xml")))],
    // #397 scope: an unknown declared encoding outside the main part
    [
        "397-encoding-unknown-styles",
        "#397",
        () => buildDocx({ ...P, "word/styles.xml": P["word/styles.xml"].replace('encoding="UTF-8"', 'encoding="X-UNKNOWN-ENCODING"') }),
    ],
    [
        "397-encoding-unknown-content-types",
        "#397",
        () => buildDocx({ ...P, "[Content_Types].xml": CT.replace('encoding="UTF-8"', 'encoding="X-UNKNOWN-ENCODING"') }),
    ],
    [
        "397-encoding-unknown-package-rels",
        "#397",
        () => buildDocx({ ...P, "_rels/.rels": P["_rels/.rels"].replace('encoding="UTF-8"', 'encoding="X-UNKNOWN-ENCODING"') }),
    ],
    [
        "397-encoding-unknown-core",
        "#397",
        () =>
            buildDocx(
                Object.fromEntries(
                    Object.entries(withCore(CORE_CT)).map(([k, v]) => [
                        k,
                        k === "docProps/core.xml" ? (v as string).replace('encoding="UTF-8"', 'encoding="X-UNKNOWN-ENCODING"') : v,
                    ]),
                ),
            ),
    ],
    // #397: bytes that do not decode in the declared encoding (UTF-8 declared, a lone Latin-1 byte)
    [
        "397-invalid-utf8-bytes",
        "#397",
        async () =>
            buildDocx({
                ...P,
                "word/document.xml": Buffer.concat([
                    Buffer.from(P["word/document.xml"].split("Hello")[0]),
                    Buffer.from([0x48, 0xe9, 0x6c]),
                    Buffer.from(P["word/document.xml"].split("Hello")[1]),
                ]),
            }),
    ],
];

const only = process.argv.slice(2);
for (const [name, issue, make] of specimens) {
    if (only.length && !only.includes(name)) continue;
    if (done.has(name)) {
        console.log(`skip ${name}`);
        continue;
    }
    const bytes = await make();
    const file = path.join(OUT, "specimens", `${name}.docx`);
    writeFileSync(file, bytes);
    const sha = createHash("sha256").update(bytes).digest("hex");
    const verdict = async (profile: "word-valid" | "strict") => {
        const r = await validate(file, { profile });
        return { valid: r.valid, errors: r.issues.filter((i) => i.severity === "error").map((i) => `${i.code}@${i.path ?? "-"}`) };
    };
    const wv = await verdict("word-valid");
    const st = await verdict("strict");
    const before = { log: logSize(), drivers: foreignWordDrivers() };
    const word = await verifyInWord(new Uint8Array(bytes));
    const after = { log: logSize(), drivers: foreignWordDrivers() };
    const interference = before.log !== after.log || before.drivers.length > 0 || after.drivers.length > 0 ? { before, after } : undefined;
    const rec = {
        name,
        issue,
        sha256: sha,
        wordValid: wv,
        strict: st,
        word: {
            outcome: word.outcome,
            details: word.details,
            version: word.wordVersion,
            build: word.wordBuild,
            platform: word.platform,
            observedAt: word.observedAt,
            fileSha256: word.fileSha256,
        },
    };
    appendFileSync(interference ? INCONCLUSIVE : RESULTS, JSON.stringify(interference ? { ...rec, interference } : rec) + "\n");
    console.log(
        `${interference ? "INCONCLUSIVE " : ""}${name}: word=${word.outcome} | word-valid=${wv.valid} ${wv.errors.join(",")} | strict=${st.valid}`,
    );
}
```

</details>

<details><summary>Session 3 specimen builder</summary>

The same probe script, extended. Session 3 used its build mode only: `BUILD_DIR=<dir> BUILD_FROM=<first specimen> node --import tsx ./probe.tmp.mts`, run as a copy at the repository root. It writes the specimens and `specimens.jsonl` (hash and verdicts), and opens nothing. JSZip stamps entry times, so a rebuild yields the same shapes under different hashes. The hashes above identify the files Word was given.

<!-- prettier-ignore -->
```ts
// Real-Word probes for #388, #392, #396, #397, #351, #398. Run from the stack worktree root as a copy.
import { createHash } from "node:crypto";
import { execSync } from "node:child_process";
import { appendFileSync, existsSync, readdirSync, readFileSync, statSync, writeFileSync } from "node:fs";
import path from "node:path";
import JSZip from "jszip";
import { MINIMAL_PARTS, buildDocx } from "./tests/support/minimal-docx";
import { validate } from "./src/scripts/office/validate";
import { verifyInWord } from "./src/lib/word-verifier";

const OUT = "/Users/arthrod/temp/T/wfv-scratch/word-probes";
const RESULTS = path.join(OUT, "results.jsonl");
// Interference guard: another session drives Word through word_pdf.py and logs to fix013w/oracle*/redline.log
// (the directory name changed between its batches, so watch every one that exists).
const FIX013W = "/Users/arthrod/temp/T/fix013w";
const INCONCLUSIVE = path.join(OUT, "results.inconclusive.jsonl");
const logSize = () => {
  let dirs: string[] = [];
  try { dirs = readdirSync(FIX013W).filter((d) => d.startsWith("oracle")); } catch { /* none */ }
  return dirs.map((d) => { const f = path.join(FIX013W, d, "redline.log"); try { return `${d}:${statSync(f).size}`; } catch { return `${d}:-1`; } }).join(",");
};
const foreignWordDrivers = () => {
  // Match the executable (ucomm), not command-line text: agents' prompts and shells mention osascript.
  const rows = execSync("ps -Ao pid=,ppid=,ucomm=").toString().split("\n").map((l) => l.trim().match(/^(\d+)\s+(\d+)\s+(.*)$/)).filter((m): m is RegExpMatchArray => m !== null).map((m) => ({ pid: Number(m[1]), ppid: Number(m[2]), exe: m[3] }));
  const cmds = new Map(execSync("ps -Ao pid=,command=").toString().split("\n").map((l) => l.trim().match(/^(\d+)\s+(.*)$/)).filter((m): m is RegExpMatchArray => m !== null).map((m) => [Number(m[1]), m[2]]));
  const parent = new Map(rows.map((r) => [r.pid, r.ppid]));
  const mine = (pid: number) => { for (let p = pid, i = 0; p > 1 && i < 64; p = parent.get(p) ?? 1, i++) if (p === process.pid) return true; return false; };
  const driver = (r: { pid: number; exe: string }) => r.exe === "osascript" || (/^python/i.test(r.exe) && /word_pdf|word_redline/.test(cmds.get(r.pid) ?? ""));
  return rows.filter((r) => driver(r) && !mine(r.pid)).map((r) => `${r.pid} ${r.exe} ${(cmds.get(r.pid) ?? "").slice(0, 80)}`);
};
const done = new Set(existsSync(RESULTS) ? readFileSync(RESULTS, "utf8").trim().split("\n").filter(Boolean).map((l) => JSON.parse(l).name) : []);

const P = MINIMAL_PARTS as Record<string, string>;
const CT = P["[Content_Types].xml"];
const without = (o: Record<string, string | Uint8Array>, k: string) => { const { [k]: _, ...r } = o; return r; };
const upperWord = (parts: Record<string, string | Uint8Array>) =>
  Object.fromEntries(Object.entries(parts).map(([k, v]) => [k.startsWith("word/") ? "WORD/" + k.slice(5) : k, v]));
const invalidCt = CT.replace("</Types>", `<Override PartName="/word/x.xml" ContentType="invalid-type"/></Types>`);
const rawZip = async (entries: [string, string | Uint8Array][]) => { const z = new JSZip(); for (const [k, v] of entries) z.file(k, v); return z.generateAsync({ type: "nodebuffer", compression: "DEFLATE" }); };
const renameInFixture = async (fixture: string, from: string, to: string) => {
  const z = await JSZip.loadAsync(readFileSync(fixture));
  const out = new JSZip();
  for (const [name, f] of Object.entries(z.files)) { if (f.dir) continue; out.file(name === from ? to : name, await f.async("uint8array")); }
  return out.generateAsync({ type: "nodebuffer", compression: "DEFLATE" });
};
const DUP_THEME = "tests/fixtures/broken/package.word-duplicate-theme.docx";
const badDoc = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:document`;
const enc = (e: string) => P["word/document.xml"].replace('encoding="UTF-8"', `encoding="${e}"`);

const R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const ctSwap = (from: string, to: string) => CT.replace(from, to);
const STYLES_CT = "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml";
const MAIN_CT = "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml";
const CORE_CT = "application/vnd.openxmlformats-package.core-properties+xml";
const CORE_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/"><dc:title>t</dc:title></cp:coreProperties>`;
const APP_CT = "application/vnd.openxmlformats-officedocument.extended-properties+xml";
const APP_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"><Application>Microsoft Office Word</Application></Properties>`;
const withApp = (appXml: string) => ({
  ...P,
  "docProps/app.xml": appXml,
  "_rels/.rels": P["_rels/.rels"].replace("</Relationships>", `<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/></Relationships>`),
  "[Content_Types].xml": CT.replace("</Types>", `<Override PartName="/docProps/app.xml" ContentType="${APP_CT}"/></Types>`),
});
const withCore = (coreType: string) => ({
  ...P,
  "docProps/core.xml": CORE_XML,
  "_rels/.rels": P["_rels/.rels"].replace("</Relationships>", `<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/></Relationships>`),
  "[Content_Types].xml": CT.replace("</Types>", `<Override PartName="/docProps/core.xml" ContentType="${coreType}"/></Types>`),
});
const docH = P["word/document.xml"].replace("<w:document ", `<w:document xmlns:r="${R_NS}" `).replace("<w:r><w:t>Hello</w:t></w:r>", `<w:hyperlink r:id="rId2"><w:r><w:t>Hello</w:t></w:r></w:hyperlink>`);
const relsH = P["word/_rels/document.xml.rels"].replace("</Relationships>", `<Relationship Id="rId2" Type="${R_NS}/hyperlink" Target="https://example.com/" TargetMode="External"/></Relationships>`);
const renamed = (from: string, to: string, parts: Record<string, string | Uint8Array> = P) => Object.fromEntries(Object.entries(parts).map(([k, v]) => [k === from ? to : k, v]));

// Session 3 (2026-10-24 ~04:40Z).
// Item names JSZip cannot write (two identical names, an unresolved `..`): write a same-length placeholder, then
// patch its bytes in the local header and the central directory. Exactly two occurrences are required.
const patchNames = (buf: Buffer, pairs: [string, string][]) => {
  for (const [from, to] of pairs) {
    if (from.length !== to.length) throw new Error(`${from} -> ${to}: lengths differ`);
    const f = Buffer.from(from), t = Buffer.from(to);
    let n = 0;
    for (let i = buf.indexOf(f); i !== -1; i = buf.indexOf(f, i + f.length)) { t.copy(buf, i); n += 1; }
    if (n !== 2) throw new Error(`${from}: ${n} occurrences, expected 2`);
  }
  return buf;
};
const CUSTOM_CT = "application/vnd.openxmlformats-officedocument.custom-properties+xml";
const CUSTOM_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"><property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="2" name="P"><vt:lpwstr>v</vt:lpwstr></property></Properties>`;
const withCustom = (customXml: string) => ({
  ...P,
  "docProps/custom.xml": customXml,
  "_rels/.rels": P["_rels/.rels"].replace("</Relationships>", `<Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties" Target="docProps/custom.xml"/></Relationships>`),
  "[Content_Types].xml": CT.replace("</Types>", `<Override PartName="/docProps/custom.xml" ContentType="${CUSTOM_CT}"/></Types>`),
});
const coreWith = (coreXml: string) => ({ ...withCore(CORE_CT), "docProps/core.xml": coreXml });
// The core-properties part stored under another name, named by its package relationship and its Override.
const coreAt = (name: string, coreXml: string) => ({
  ...P,
  [name]: coreXml,
  "_rels/.rels": P["_rels/.rels"].replace("</Relationships>", `<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="${name}"/></Relationships>`),
  "[Content_Types].xml": CT.replace("</Types>", `<Override PartName="/${name}" ContentType="${CORE_CT}"/></Types>`),
});
// #410: the main part declared in a real encoding libxml2 cannot decode, with "Hello"'s "e" replaced by one
// correctly encoded non-ASCII character, so a decoder that ignores the label would misread it.
const docIn = (label: string, special: Buffer) => {
  const [a, b] = enc(label).split("Hello");
  return Buffer.concat([Buffer.from(`${a}H`), special, Buffer.from(`llo${b}`)]);
};
const STYLES_OV = `<Override PartName="/word/styles.xml" ContentType="${STYLES_CT}"/>`;
const stylesOverrideNamed = (name: string) => CT.replace(STYLES_OV, `<Override PartName="${name}" ContentType="${STYLES_CT}"/>`);
const extraInvalidOverride = (name: string) => CT.replace("</Types>", `<Override PartName="${name}" ContentType="invalid-x"/></Types>`);
const XML_DEFAULT = `<Default Extension="xml" ContentType="application/xml"/>`;
const extraItem = (name: string) => rawZip([...Object.entries(P), [name, P["word/styles.xml"]]]);

const specimens: [string, string, () => Promise<Buffer>][] = [
  ["control-minimal", "control", () => buildDocx(P)],
  // #388
  ["388-rels-zero-byte", "#388", () => buildDocx({ ...P, "_rels/.rels": "" })],
  ["388-rels-truncated", "#388", () => buildDocx({ ...P, "_rels/.rels": P["_rels/.rels"].slice(0, 120) })],
  ["388-rels-undeclared-prefix", "#388", () => buildDocx({ ...P, "_rels/.rels": P["_rels/.rels"].replace("<Relationships ", "<x:Relationships ").replace("</Relationships>", "</x:Relationships>") })],
  // #392
  ["392-WORD-folder-clean", "#392", () => buildDocx(upperWord(P))],
  ["392-WORD-folder-malformed-main", "#392", () => buildDocx(upperWord({ ...P, "word/document.xml": badDoc }))],
  ["392-word-folder-malformed-main-control", "#392", () => buildDocx({ ...P, "word/document.xml": badDoc })],
  // #396
  ["396-CT-upper-clean", "#396", () => buildDocx({ ...without(P, "[Content_Types].xml"), "[CONTENT_TYPES].XML": CT })],
  ["396-CT-upper-invalid-type", "#396", () => buildDocx({ ...without(P, "[Content_Types].xml"), "[CONTENT_TYPES].XML": invalidCt })],
  ["396-CT-exact-invalid-type-control", "#396", () => buildDocx({ ...P, "[Content_Types].xml": invalidCt })],
  ["396-CT-upper-malformed", "#396", () => buildDocx({ ...without(P, "[Content_Types].xml"), "[CONTENT_TYPES].XML": CT.slice(0, 150) })],
  // #397
  ["397-encoding-unknown", "#397", () => buildDocx({ ...P, "word/document.xml": enc("X-UNKNOWN-ENCODING") })],
  ["397-encoding-windows-1252-control", "#397", () => buildDocx({ ...P, "word/document.xml": enc("windows-1252") })],
  // #351
  ["351-dup-case-identical", "#351", () => rawZip([...Object.entries(P), ["WORD/document.xml", P["word/document.xml"]]])],
  ["351-dup-case-lower-good-upper-bad", "#351", () => rawZip([...Object.entries(P), ["WORD/document.xml", badDoc]])],
  ["351-dup-case-upper-bad-first", "#351", () => rawZip([["WORD/document.xml", badDoc], ...Object.entries(P)])],
  ["351-dup-stream-good-then-bad", "#351", () => rawZip([...Object.entries(P), ["[CONTENT_TYPES].XML", invalidCt]])],
  ["351-dup-stream-bad-then-good", "#351", () => rawZip([["[CONTENT_TYPES].XML", invalidCt], ...Object.entries(P)])],
  // #398
  ["398-RELS-sidecar-clean", "#398", () => buildDocx({ ...without(P, "word/_rels/document.xml.rels"), "word/_rels/document.xml.RELS": P["word/_rels/document.xml.rels"] })],
  ["398-dup-theme-control", "#398", async () => readFileSync(DUP_THEME)],
  ["398-dup-theme-RELS", "#398", () => renameInFixture(DUP_THEME, "word/_rels/document.xml.rels", "word/_rels/document.xml.RELS")],
  // word-content-type-invalid scope (Override for an existing part vs a missing one; unused Default)
  ["cti-override-existing-styles", "cti", () => buildDocx({ ...P, "[Content_Types].xml": ctSwap(`ContentType="${STYLES_CT}"`, `ContentType="invalid-${STYLES_CT}"`) })],
  ["cti-override-existing-main", "cti", () => buildDocx({ ...P, "[Content_Types].xml": ctSwap(`ContentType="${MAIN_CT}"`, `ContentType="invalid-${MAIN_CT}"`) })],
  ["cti-override-existing-core", "cti", () => buildDocx(withCore(`invalid-${CORE_CT}`))],
  ["cti-override-existing-core-control", "cti", () => buildDocx(withCore(CORE_CT))],
  ["cti-default-unused-extension", "cti", () => buildDocx({ ...P, "[Content_Types].xml": CT.replace("<Default Extension=\"rels\"", `<Default Extension="foo" ContentType="invalid-type"/><Default Extension="rels"`) })],
  // #351 escape / backslash duplicates, and the single-spelling controls
  ["351-dup-escape-identical", "#351", () => rawZip([...Object.entries(P), ["word/%73tyles.xml", P["word/styles.xml"]]])],
  ["351-dup-backslash-identical", "#351", () => rawZip([...Object.entries(P), ["word\\document.xml", P["word/document.xml"]]])],
  ["351-escape-only-styles", "#351", () => rawZip(Object.entries(renamed("word/styles.xml", "word/%73tyles.xml")))],
  ["351-backslash-only-main", "#351", () => rawZip(Object.entries(renamed("word/document.xml", "word\\document.xml")))],
  // #398 follow-up: hyperlink resolution through a .RELS sidecar
  ["398-hyperlink-rels-control", "#398", () => buildDocx({ ...P, "word/document.xml": docH, "word/_rels/document.xml.rels": relsH })],
  ["398-hyperlink-%72els", "#402", () => buildDocx({ ...without({ ...P, "word/document.xml": docH }, "word/_rels/document.xml.rels"), "word/_rels/document.xml.%72els": relsH })],
  ["398-hyperlink-RELS", "#398", () => buildDocx({ ...without({ ...P, "word/document.xml": docH }, "word/_rels/document.xml.rels"), "word/_rels/document.xml.RELS": relsH })],
  // docProps scope for xml-syntax / xml-encoding (#397 follow-up)
  ["docprops-core-truncated", "docprops", () => buildDocx({ ...withCore(CORE_CT), "docProps/core.xml": CORE_XML.slice(0, 140) })],
  ["docprops-app-control", "docprops", () => buildDocx(withApp(APP_XML))],
  ["docprops-app-truncated", "docprops", () => buildDocx(withApp(APP_XML.slice(0, 140)))],
  ["docprops-app-encoding-unknown", "docprops", () => buildDocx(withApp(APP_XML.replace('encoding="UTF-8"', 'encoding="X-UNKNOWN-ENCODING"')))],
  // Escaped ZIP item names: does Word decode %XX of an unreserved character, or refuse the name itself?
  ["esc-unreferenced-extra-item", "esc", () => rawZip([...Object.entries(P), ["word/%78tra.xml", P["word/styles.xml"]]])],
  ["esc-styles-missing-control", "esc", () => rawZip(Object.entries(without(P, "word/styles.xml")))],
  ["esc-only-main", "esc", () => rawZip(Object.entries(renamed("word/document.xml", "word/%64ocument.xml")))],
  // #397 scope: an unknown declared encoding outside the main part
  ["397-encoding-unknown-styles", "#397", () => buildDocx({ ...P, "word/styles.xml": P["word/styles.xml"].replace('encoding="UTF-8"', 'encoding="X-UNKNOWN-ENCODING"') })],
  ["397-encoding-unknown-content-types", "#397", () => buildDocx({ ...P, "[Content_Types].xml": CT.replace('encoding="UTF-8"', 'encoding="X-UNKNOWN-ENCODING"') })],
  ["397-encoding-unknown-package-rels", "#397", () => buildDocx({ ...P, "_rels/.rels": P["_rels/.rels"].replace('encoding="UTF-8"', 'encoding="X-UNKNOWN-ENCODING"') })],
  ["397-encoding-unknown-core", "#397", () => buildDocx(Object.fromEntries(Object.entries(withCore(CORE_CT)).map(([k, v]) => [k, k === "docProps/core.xml" ? (v as string).replace('encoding="UTF-8"', 'encoding="X-UNKNOWN-ENCODING"') : v])))],
  // #397: bytes that do not decode in the declared encoding (UTF-8 declared, a lone Latin-1 byte)
  ["397-invalid-utf8-bytes", "#397", async () => buildDocx({ ...P, "word/document.xml": Buffer.concat([Buffer.from(P["word/document.xml"].split("Hello")[0]), Buffer.from([0x48, 0xe9, 0x6c]), Buffer.from(P["word/document.xml"].split("Hello")[1])]) })],
  // ---- Session 3 ----
  ["s3-control-minimal-start", "control", () => buildDocx(P)],
  // #409: the docProps parts and defects not yet observed (the two app.xml rows repeat session 2's inconclusive ones)
  ["docprops-app-truncated", "#409", () => buildDocx(withApp(APP_XML.slice(0, 140)))],
  ["docprops-app-encoding-unknown", "#409", () => buildDocx(withApp(APP_XML.replace('encoding="UTF-8"', 'encoding="X-UNKNOWN-ENCODING"')))],
  ["docprops-app-undeclared-prefix", "#409", () => buildDocx(withApp(APP_XML.replace("<Application>Microsoft Office Word</Application>", "<ax:Application>Microsoft Office Word</ax:Application>")))],
  ["docprops-core-undeclared-prefix", "#409", () => buildDocx(coreWith(CORE_XML.replace("<dc:title>t</dc:title>", "<dcx:title>t</dcx:title>")))],
  ["docprops-custom-control", "#409", () => buildDocx(withCustom(CUSTOM_XML))],
  ["docprops-custom-truncated", "#409", () => buildDocx(withCustom(CUSTOM_XML.slice(0, 160)))],
  ["docprops-custom-undeclared-prefix", "#409", () => buildDocx(withCustom(CUSTOM_XML.replace("<vt:lpwstr>v</vt:lpwstr>", "<vx:lpwstr>v</vx:lpwstr>")))],
  ["docprops-custom-encoding-unknown", "#409", () => buildDocx(withCustom(CUSTOM_XML.replace('encoding="UTF-8"', 'encoding="X-UNKNOWN-ENCODING"')))],
  // #410: real encodings libxml2 has no decoder for
  ["410-macintosh", "#410", async () => buildDocx({ ...P, "word/document.xml": docIn("macintosh", Buffer.from([0x8e])) })],
  ["410-x-mac-cyrillic", "#410", async () => buildDocx({ ...P, "word/document.xml": docIn("x-mac-cyrillic", Buffer.from([0x80])) })],
  ["410-IBM866", "#410", async () => buildDocx({ ...P, "word/document.xml": docIn("IBM866", Buffer.from([0x80])) })],
  ["410-UTF-7", "#410", async () => buildDocx({ ...P, "word/document.xml": docIn("UTF-7", Buffer.from("+AOk-")) })],
  ["410-windows-1252-nonascii-control", "#410", async () => buildDocx({ ...P, "word/document.xml": docIn("windows-1252", Buffer.from([0xe9])) })],
  // #411: two items with the identical name, and a `..` name that JSZip resolves at load
  ["411-identical-valid-valid", "#411", async () => patchNames(await rawZip([...Object.entries(P), ["word/documenX.xml", P["word/document.xml"]]]), [["word/documenX.xml", "word/document.xml"]])],
  ["411-identical-valid-then-malformed", "#411", async () => patchNames(await rawZip([...Object.entries(P), ["word/documenX.xml", badDoc]]), [["word/documenX.xml", "word/document.xml"]])],
  ["411-identical-malformed-then-valid", "#411", async () => patchNames(await rawZip([["word/documenX.xml", badDoc], ...Object.entries(P)]), [["word/documenX.xml", "word/document.xml"]])],
  ["411-dotdot-unreferenced-item", "#411", async () => patchNames(await extraItem("word/a/zzzxtra.xml"), [["word/a/zzzxtra.xml", "word/a/../xtra.xml"]])],
  ["411-dotdot-then-plain-styles", "#411", async () => patchNames(await rawZip([["word/a/zzzstyles.xml", P["word/styles.xml"]], ...Object.entries(P)]), [["word/a/zzzstyles.xml", "word/a/../styles.xml"]])],
  // #414: an Override PartName that is not a valid part name. Valid styles type swapped onto a malformed name
  // (so styles falls back to the xml Default; the removed-Override control isolates that), and an extra
  // invalid-x Override whose malformed name would resolve to the main part.
  ["414-styles-override-removed-control", "#414", () => buildDocx({ ...P, "[Content_Types].xml": CT.replace(STYLES_OV, "") })],
  ["414-styles-override-dotdot", "#414", () => buildDocx({ ...P, "[Content_Types].xml": stylesOverrideNamed("/word/../word/styles.xml") })],
  ["414-styles-override-query", "#414", () => buildDocx({ ...P, "[Content_Types].xml": stylesOverrideNamed("/word/styles.xml?x=1") })],
  ["414-styles-override-fragment", "#414", () => buildDocx({ ...P, "[Content_Types].xml": stylesOverrideNamed("/word/styles.xml#f") })],
  ["414-styles-override-backslash", "#414", () => buildDocx({ ...P, "[Content_Types].xml": stylesOverrideNamed("/word\\styles.xml") })],
  ["414-invalid-x-dotdot-main", "#414", () => buildDocx({ ...P, "[Content_Types].xml": extraInvalidOverride("/word/../word/document.xml") })],
  ["414-invalid-x-query-main", "#414", () => buildDocx({ ...P, "[Content_Types].xml": extraInvalidOverride("/word/document.xml?x=1") })],
  ["414-invalid-x-fragment-main", "#414", () => buildDocx({ ...P, "[Content_Types].xml": extraInvalidOverride("/word/document.xml#f") })],
  ["414-invalid-x-backslash-main", "#414", () => buildDocx({ ...P, "[Content_Types].xml": extraInvalidOverride("/word\\document.xml") })],
  // Review V F2: duplicate content-type declarations (System.IO.Packaging throws on a duplicate key)
  ["dupdecl-default-identical", "dupdecl", () => buildDocx({ ...P, "[Content_Types].xml": CT.replace(XML_DEFAULT, XML_DEFAULT + XML_DEFAULT) })],
  ["dupdecl-default-case", "dupdecl", () => buildDocx({ ...P, "[Content_Types].xml": CT.replace(XML_DEFAULT, `${XML_DEFAULT}<Default Extension="XML" ContentType="application/xml"/>`) })],
  ["dupdecl-override-identical", "dupdecl", () => buildDocx({ ...P, "[Content_Types].xml": CT.replace(STYLES_OV, STYLES_OV + STYLES_OV) })],
  ["dupdecl-override-conflicting", "dupdecl", () => buildDocx({ ...P, "[Content_Types].xml": CT.replace(STYLES_OV, `${STYLES_OV}<Override PartName="/word/styles.xml" ContentType="application/xml"/>`) })],
  ["dupdecl-override-case", "dupdecl", () => buildDocx({ ...P, "[Content_Types].xml": CT.replace(STYLES_OV, `${STYLES_OV}<Override PartName="/WORD/styles.xml" ContentType="${STYLES_CT}"/>`) })],
  // #408: the unreserved escapes Word has not been shown, on an unreferenced item; a plain-name control, and %20
  // (a space is not unreserved, so [M1.8] permits its escape)
  ["esc-plain-unreferenced-control", "#408", () => extraItem("word/xtra.xml")],
  ["esc-%41-upper-letter", "#408", () => extraItem("word/x%41.xml")],
  ["esc-%30-digit", "#408", () => extraItem("word/x%30.xml")],
  ["esc-%2D-hyphen", "#408", () => extraItem("word/x%2D.xml")],
  ["esc-%2d-hyphen-lowerhex", "#408", () => extraItem("word/x%2d.xml")],
  ["esc-%2E-period", "#408", () => extraItem("word/x%2E.xml")],
  ["esc-%5F-underscore", "#408", () => extraItem("word/x%5F.xml")],
  ["esc-%7E-tilde", "#408", () => extraItem("word/x%7E.xml")],
  ["esc-%20-space-control", "#408", () => extraItem("word/x%20y.xml")],
  ["s3-control-minimal-end", "control", () => buildDocx(P)],
  // CodeRabbit on #383 (word-document-root.ts:105): the package relationship read as a descendant, not a
  // direct child of Relationships. Does Word find a main part that is not a direct child?
  ["rels-main-nested-in-unknown-element", "cr383", () => buildDocx({ ...P, "_rels/.rels": P["_rels/.rels"].replace(/(<Relationship [^>]*\/>)/, `<x:Wrap xmlns:x="urn:x">$1</x:Wrap>`) })],
  ["rels-root-not-relationships", "cr383", () => buildDocx({ ...P, "_rels/.rels": P["_rels/.rels"].replace("<Relationships ", "<Foo ").replace("</Relationships>", "</Foo>") })],
  ["rels-unknown-sibling-control", "cr383", () => buildDocx({ ...P, "_rels/.rels": P["_rels/.rels"].replace("</Relationships>", `<x:Extra xmlns:x="urn:x"/></Relationships>`) })],
  // ---- Session 3c ----
  ["s3c-control-minimal", "control", () => buildDocx(P)],
  // Review X F1: the content-types stream spelled with an escape (not a part, so outside [M1.8]'s letter)
  ["ct-stream-missing-control", "rx-f1", () => rawZip(Object.entries(without(P, "[Content_Types].xml")))],
  ["ct-stream-escaped-only", "rx-f1", () => rawZip(Object.entries(renamed("[Content_Types].xml", "[Content_Types].%78ml")))],
  ["ct-stream-escaped-extra", "rx-f1", () => rawZip([...Object.entries(P), ["[Content_Types].%78ml", CT]])],
  // [M1.7]: an escaped "/" or "\" in a stored name (reserved, so outside [M1.8]); nothing reports these today
  ["esc-pct2F-slash-unreferenced", "m17", () => extraItem("word%2Fxtra.xml")],
  ["esc-pct5C-backslash-unreferenced", "m17", () => extraItem("word%5Cxtra.xml")],
  // #409: Word reads a docProps part only through its package relationship. annotations_import_2.docx (clean-open)
  // stores an unreferenced custom.xml with an undeclared prefix; these isolate "unreferenced" in minimal packages.
  ["docprops-custom-unreferenced-undeclared-prefix", "#409", () => buildDocx({ ...withCustom(CUSTOM_XML.replace("<vt:lpwstr>v</vt:lpwstr>", "<vx:lpwstr>v</vx:lpwstr>")), "_rels/.rels": P["_rels/.rels"] })],
  ["docprops-custom-unreferenced-truncated", "#409", () => buildDocx({ ...withCustom(CUSTOM_XML.slice(0, 160)), "_rels/.rels": P["_rels/.rels"] })],
  ["docprops-app-unreferenced-truncated", "#409", () => buildDocx({ ...withApp(APP_XML.slice(0, 140)), "_rels/.rels": P["_rels/.rels"] })],
  // The existing policy blocks xml-syntax anywhere under word/. Is an unreferenced malformed word/ part read at all?
  ["word-unreferenced-malformed-part", "policy", () => buildDocx({ ...P, "word/xtra.xml": P["word/styles.xml"].slice(0, 120) })],
  ["word-unreferenced-wellformed-control", "policy", () => buildDocx({ ...P, "word/xtra.xml": P["word/styles.xml"] })],
  // #409 pass 2: does Word find the core-properties part by its package relationship or by its name?
  ["docprops-core-unreferenced-truncated", "#409", () => buildDocx({ ...withCore(CORE_CT), "docProps/core.xml": CORE_XML.slice(0, 140), "_rels/.rels": P["_rels/.rels"] })],
  ["docprops-core-renamed-truncated", "#409", () => buildDocx(coreAt("props/meta.xml", CORE_XML.slice(0, 140)))],
  ["docprops-core-renamed-control", "#409", () => buildDocx(coreAt("props/meta.xml", CORE_XML))],
  // Review W2 F7 and #380: an empty archive, and a package without its main part, pass word-valid today.
  ["empty-archive", "w2-f7", () => rawZip([])],
  ["main-part-missing", "#380", () => buildDocx(without(P, "word/document.xml"))],
  ["main-relationship-missing", "#380", () => buildDocx({ ...P, "_rels/.rels": P["_rels/.rels"].replace(/<Relationship [^>]*\/officeDocument"[^>]*\/>/, "") })],
  // #414: an invalid-x Override for a part that does not exist opens clean when its PartName is plain (#401).
  // Spelled as a non-part-name, does Word reject the declaration itself, or ignore it?
  ["414-invalid-x-plain-missing-control", "#414", () => buildDocx({ ...P, "[Content_Types].xml": extraInvalidOverride("/word/nothere.xml") })],
  ["414-invalid-x-dotdot-missing", "#414", () => buildDocx({ ...P, "[Content_Types].xml": extraInvalidOverride("/word/../word/nothere.xml") })],
  ["414-invalid-x-query-missing", "#414", () => buildDocx({ ...P, "[Content_Types].xml": extraInvalidOverride("/word/nothere.xml?x=1") })],
  ["414-invalid-x-fragment-missing", "#414", () => buildDocx({ ...P, "[Content_Types].xml": extraInvalidOverride("/word/nothere.xml#f") })],
  ["414-invalid-x-backslash-missing", "#414", () => buildDocx({ ...P, "[Content_Types].xml": extraInvalidOverride("/word\\nothere.xml") })],
  // Session 3b's batch could not attribute this control (Word opened a document under another name); observe it again.
  ["rels-unknown-sibling-control", "cr383", () => buildDocx({ ...P, "_rels/.rels": P["_rels/.rels"].replace("</Relationships>", `<x:Extra xmlns:x="urn:x"/></Relationships>`) })],
];

// Build-only mode (session 3b, 2026-10-24): write the session-3 specimens, the #383 relationship specimens and
// session-2 calibration rows into BUILD_DIR, with their hashes and validator verdicts, for
// neurotic_docx_bench/scripts/word_pdf.py to open in Word. No Word here. `%` is spelled `pct` in file names.
if (process.env.BUILD_DIR) {
  const dir = process.env.BUILD_DIR;
  const calibration = new Set(["control-minimal", "docprops-core-truncated", "397-encoding-unknown-core", "397-encoding-windows-1252-control", "cti-override-existing-core-control", "351-dup-case-identical", "esc-unreferenced-extra-item", "398-hyperlink-RELS"]);
  // BUILD_FROM picks a later block (e.g. "s3c-control-minimal") and drops the calibration rows.
  const from = process.env.BUILD_FROM;
  if (from) calibration.clear();
  const start = specimens.findIndex(([n]) => n === (from ?? "s3-control-minimal-start"));
  if (start < 0) throw new Error(`no specimen named ${from}`);
  const built = new Set<string>();
  for (const [i, [name, issue, make]] of specimens.entries()) {
    if (built.has(name) || (i < start && !calibration.has(name))) continue;
    built.add(name);
    const bytes = await make();
    const file = `${name.replace(/%/g, "pct")}.docx`;
    writeFileSync(path.join(dir, file), bytes);
    const sha256 = createHash("sha256").update(bytes).digest("hex");
    const verdict = async (profile: "word-valid" | "strict") => {
      const r = await validate(path.join(dir, file), { profile });
      return { valid: r.valid, errors: r.issues.filter((x) => x.severity === "error").map((x) => `${x.code}@${x.path ?? "-"}`) };
    };
    const row = { name, issue, file, sha256, calibration: i < start, wordValid: await verdict("word-valid"), strict: await verdict("strict") };
    appendFileSync(path.join(dir, "..", "specimens.jsonl"), JSON.stringify(row) + "\n");
    console.log(`built ${file} ${sha256.slice(0, 16)} word-valid=${row.wordValid.valid} strict=${row.strict.valid}`);
  }
  process.exit(0);
}

const only = process.argv.slice(2);
const ran = new Set<string>();
for (const [name, issue, make] of specimens) {
  if (only.length && !only.includes(name)) continue;
  if (done.has(name)) { console.log(`skip ${name}`); continue; }
  if (ran.has(name)) continue;
  ran.add(name);
  const bytes = await make();
  const file = path.join(OUT, "specimens", `${name}.docx`);
  writeFileSync(file, bytes);
  const sha = createHash("sha256").update(bytes).digest("hex");
  const verdict = async (profile: "word-valid" | "strict") => {
    const r = await validate(file, { profile });
    return { valid: r.valid, errors: r.issues.filter((i) => i.severity === "error").map((i) => `${i.code}@${i.path ?? "-"}`) };
  };
  const wv = await verdict("word-valid");
  const st = await verdict("strict");
  if (process.env.DRY) { console.log(`DRY ${name}: ${sha.slice(0, 16)} | word-valid=${wv.valid} ${wv.errors.join(",")} | strict=${st.valid} ${st.errors.join(",")}`); continue; }
  const before = { log: logSize(), drivers: foreignWordDrivers() };
  const word = await verifyInWord(new Uint8Array(bytes));
  const after = { log: logSize(), drivers: foreignWordDrivers() };
  const interference = before.log !== after.log || before.drivers.length > 0 || after.drivers.length > 0 ? { before, after } : undefined;
  const rec = { name, issue, sha256: sha, wordValid: wv, strict: st, word: { outcome: word.outcome, details: word.details, version: word.wordVersion, build: word.wordBuild, platform: word.platform, observedAt: word.observedAt, fileSha256: word.fileSha256 } };
  appendFileSync(interference ? INCONCLUSIVE : RESULTS, JSON.stringify(interference ? { ...rec, interference } : rec) + "\n");
  console.log(`${interference ? "INCONCLUSIVE " : ""}${name}: word=${word.outcome} | word-valid=${wv.valid} ${wv.errors.join(",")} | strict=${st.valid}`);
}
```

</details>

<details><summary>Session 3 outcome parser</summary>

`python3 wordpdf_outcomes.py <run-dir>` after `uv run --script scripts/word_pdf.py --src <run-dir>/src --out <run-dir>/pdf --log <run-dir>/word_pdf.log --label wfv --no-check-preset --lock-wait 600`, run from the `neurotic_docx_bench` checkout.

```python
"""Turn a neurotic_docx_bench word_pdf.py batch run into per-specimen Word outcomes.

Usage: python3 wordpdf_outcomes.py <run-dir>   (holds specimens.jsonl, src/, word_pdf.log, pdf/manifest.jsonl)

word_pdf.py runs in its default batch mode (one osascript, Word brought forward to answer
alerts) and stages item N as `NNNNN__wfv__<file>`. Attribution is by that staged name in
the dialog text, never by timing; a name Word truncated ("…__wfv__414-invalid-x-fragment-ma...")
is matched by index, else by a unique prefix. Names without the five-digit index belong to an
earlier serial run and are ignored.

  repair-warning  a repair prompt ("Word found unreadable content in <name>") names the file, or
                  word_pdf.py's manifest records "repair prompt declined" for it
  open-error      an "error trying to open the file" alert names the file, with no repair prompt
  clean-open      the manifest says ok (Word opened it and saved a PDF) and no dialog names it
  other           anything else, with the manifest's error and every dialog, for a person to read
"""

from __future__ import annotations

import json
import re
import subprocess
import sys
from datetime import datetime, timedelta, timezone
from pathlib import Path

run = Path(sys.argv[1])
ENTRY = re.compile(r"^(\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}\.\d{3}) \| (\w+)\s*\| (.*)$")
STAGED = re.compile(r"(\d{5})__wfv__([A-Za-z0-9._%+-]+)")

entries: list[tuple[str, str]] = []
for line in (run / "word_pdf.log").read_text(encoding="utf-8").splitlines():
    m = ENTRY.match(line)
    if m:
        entries.append((m.group(1), m.group(3)))
    elif entries:
        stamp, text = entries[-1]
        entries[-1] = (stamp, text + "\n" + line)

offset = timedelta(seconds=-datetime.now().astimezone().utcoffset().total_seconds())


def utc(stamp: str) -> str:
    local = datetime.strptime(stamp, "%Y-%m-%d %H:%M:%S.%f")
    return (local + offset).replace(tzinfo=timezone.utc).isoformat(timespec="milliseconds").replace("+00:00", "Z")


specs = [json.loads(l) for l in (run / "specimens.jsonl").read_text().splitlines() if l.strip()]
files = [s["file"] for s in specs]

# Index -> file, from every full (untruncated) staged name in the log.
by_index: dict[str, str] = {}
for _, text in entries:
    for idx, name in STAGED.findall(text):
        name = name.rstrip(".")
        if name in files:
            by_index[idx] = name


def named(text: str) -> set[str]:
    out: set[str] = set()
    for idx, name in STAGED.findall(text):
        full = name.rstrip(".")
        if full in files:
            out.add(full)
        elif idx in by_index:
            out.add(by_index[idx])
        else:
            hits = [f for f in files if f.startswith(full)]
            if len(hits) == 1:
                out.add(hits[0])
    return out


# Only Word's own words: the watchdog's "says:" lines and the "Word said:" text of the alerts
# word_pdf.py answered (its focus nudge and its repair-prompt decline).
said: dict[str, list[tuple[str, str]]] = {f: [] for f in files}
for stamp, text in entries:
    marker = "says:" if "says:" in text else "Word said:" if "Word said:" in text else None
    if marker is None:
        continue
    body = text.split(marker, 1)[1].strip()
    for f in named(body):
        if body not in (b for _, b in said[f]):
            said[f].append((stamp, body))

manifest: dict[str, dict] = {}
mpath = run / "pdf" / "manifest.jsonl"
if mpath.exists():
    for line in mpath.read_text().splitlines():
        if line.strip():
            rec = json.loads(line)
            manifest[rec["original_filename"]] = rec

plist = "/Applications/Microsoft Word.app/Contents/Info.plist"
version = subprocess.run(["defaults", "read", plist, "CFBundleShortVersionString"], capture_output=True, text=True).stdout.strip()
build = subprocess.run(["defaults", "read", plist, "CFBundleVersion"], capture_output=True, text=True).stdout.strip()

out = []
for spec in specs:
    f = spec["file"]
    dialogs = said[f]
    rec_error = str(manifest.get(f, {}).get("error", ""))
    # word_pdf.py's own attribution: it declined a repair prompt it tied to this item.
    repair = any("unreadable content in" in b for _, b in dialogs) or rec_error.startswith("repair prompt declined")
    open_error = any("error trying to open the file" in b for _, b in dialogs)
    rec = manifest.get(f, {})
    status = rec.get("status", "missing")
    if repair:
        outcome = "repair-warning"
    elif open_error:
        outcome = "open-error"
    elif status == "ok" and not dialogs:
        outcome = "clean-open"
    else:
        outcome = "other"
    first = dialogs[0][0] if dialogs else ""
    out.append({
        **spec,
        "word": {
            "outcome": outcome,
            "item": status,
            "error": str(rec.get("error", ""))[:240],
            "dialogs": [b[:240] for _, b in dialogs],
            "observedAt": utc(first) if first else rec.get("run_at", ""),
            "driver": "neurotic_docx_bench/scripts/word_pdf.py (default batch mode, --label wfv)",
            "version": version,
            "build": build,
        },
    })

(run / "outcomes.jsonl").write_text("".join(json.dumps(r) + "\n" for r in out))
for r in out:
    w = r["word"]
    extra = w["error"][:70] if w["outcome"] == "other" else ""
    print(f"{r['name']:50} {w['outcome']:15} word-valid={r['wordValid']['valid']!s:5} strict={r['strict']['valid']!s:5} {extra}")
```

</details>
