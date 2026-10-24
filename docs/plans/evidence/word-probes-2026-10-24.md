# Real-Word probes, 2026-10-24

Observations behind the word-valid policy for #388, #392, #396, #397, #351, #398, #402, #408 and #409, and the scope of `word-content-type-invalid` (#401).

- Word: Microsoft Word for Mac 16.115 (build 16.115.26100424), platform `darwin`.
- Transport: `verifyInWord` (`src/lib/word-verifier.ts`). It opens each specimen by path, reads Word's window and dialog state, and never saves.
- Outcomes: `clean-open` means Word opened the document with no dialog. `repair-warning` means Word showed "Word found unreadable content … Do you want to recover the contents of this document?". `open-error` means Word showed "Word experienced an error trying to open the file." Word stayed alive throughout, and `~/Library/Logs/DiagnosticReports` holds no Word report from the run.
- Specimens: built in code from `MINIMAL_PARTS` (`tests/support/minimal-docx.ts`). The `398-dup-theme-*` specimens are built from `tests/fixtures/broken/package.word-duplicate-theme.docx`, which was only read. The probe script that builds, validates and opens them is reproduced at the end of this file. No fixture or manifest was changed.
- Validator: the `word-valid` and `strict` columns record `validate()` **at probe time**, before the policy changes the probes motivated. They are not current verdicts. Session 1 (02:18Z–02:26Z) ran against `wfv/46-plan-audit` (199c2d5b); session 2 (03:10Z–03:19Z) ran against `wfv/51-directory-staging` (ad072db8).

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
- **#409: core properties.** A truncated `docProps/core.xml`, or one declaring an unknown encoding, makes Word offer recovery. Word-valid passes both, because docProps findings are downgraded. The `docProps/app.xml` control opens clean. The two `app.xml` defects are not yet observed (below).
- **#401: `word-content-type-invalid`.** An `invalid-*` type on an existing part makes Word refuse the file or offer recovery: styles, the main part and core properties. The same type on an `Override` for a missing part, or on a `Default` that covers no part, opens clean. The check over-reported; it now reports only declarations that apply to a stored part.

## Not observed

Session 1. Another session began driving Word by 02:28Z: its first log line is at 02:28:14Z, and its map file was written at 02:27:55Z. Its preflight closes every open document, so rows observed from 02:26Z on were not treated as evidence. Session 2 observed all of them again (below).

Session 2. Every record now carries a guard. Before and after each open, the probe lists the processes driving Word that it did not start itself: an `osascript` executable, or the other session's `word_pdf`/`word_redline` scripts. It also notes whether that session's log grew. A record with any such process goes to an inconclusive file instead of the table.

- The first run, at 03:08Z, matched on command-line text, not on executables: its own shell and two review agents whose prompts forbid osascript. All ten of its records were set aside, and the guard was fixed. The 03:10Z run repeated the same ten specimens with a clean guard and identical outcomes.
- At 03:19:33Z the other session started a Word batch mid-run, and the last two records were set aside.

To be observed when Word is free:

- `docprops-app-truncated` and `docprops-app-encoding-unknown` (#409), recorded during the other session's batch.
- `docProps/core.xml` with an undeclared prefix, and a truncated `docProps/custom.xml` (#409).
- Real encodings libxml2 cannot decode, such as `macintosh`, `UTF-7`, `IBM866` and `x-mac-cyrillic`, in `word/document.xml` (#410).
- Two entries with an identical name (#411).
- `Override` PartNames that are not part names: `..`, `?query`, `#fragment`, a backslash (#414).
- Duplicate `Default` or `Override` declarations in the content-types stream (review V on #406, F2).

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
