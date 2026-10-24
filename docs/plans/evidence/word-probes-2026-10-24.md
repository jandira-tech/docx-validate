# Real-Word probes, 2026-10-24

Observations behind the word-valid policy for #388, #392, #396, #397, #351 and #398, and the scope of `word-content-type-invalid` (#401).

- Word: Microsoft Word for Mac 16.115 (build 16.115.26100424), platform `darwin`.
- Transport: `verifyInWord` (`src/lib/word-verifier.ts`). It opens each specimen by path, reads Word's window and dialog state, and never saves.
- Outcomes: `clean-open` means Word opened the document with no dialog. `repair-warning` means Word showed "Word found unreadable content … Do you want to recover the contents of this document?". `open-error` means Word showed "Word experienced an error trying to open the file." Word stayed alive throughout, and `~/Library/Logs/DiagnosticReports` holds no Word report from the run.
- Specimens: built in code from `MINIMAL_PARTS` (`tests/support/minimal-docx.ts`). The `398-dup-theme-*` specimens are built from `tests/fixtures/broken/package.word-duplicate-theme.docx`, which was only read. The probe script that builds, validates and opens them is reproduced at the end of this file. No fixture or manifest was changed.
- Validator: verdicts are from `validate()` at the commit the probes ran against (`wfv/46-plan-audit`, 199c2d5b), before the policy changes they motivated.

## What the observations decide

- **#388: package relationships.** A zero-byte, truncated or undeclared-prefix `_rels/.rels` makes Word offer recovery, and the readable control opens clean. Word-valid passed all three, so `xml-syntax` and `xml-undeclared-prefix` on `_rels/.rels` are now Word-blocking.
- **#392: `WORD/` folder.** A package with every Word part under `WORD/` opens clean. With its main part malformed, Word refuses it exactly as it refuses the lower-case control. That matches the current policy, which matches `word/` by OPC part name.
- **#396: content-types stream spelling.** A stream stored as `[CONTENT_TYPES].XML` opens clean. Malformed, it makes Word offer recovery, as word-valid predicts. That matches the any-case lookup.
- **#397: unknown declared encoding.** `word/document.xml` declaring `X-UNKNOWN-ENCODING` makes Word refuse the file, and the `windows-1252` control opens clean. Word-valid passed it, so this is a false negative.
- **#351: equivalent part names.** Every duplicate probed makes Word offer recovery, even when both spellings hold identical bytes: case variants, an unreserved escape (`word/%73tyles.xml`), a backslash (`word\document.xml`), and two content-types streams. Word-valid passed the escape and backslash duplicates.
- **#398: `.RELS` sidecar.** A `.RELS`-spelled document sidecar opens clean. With two theme relationships, Word offers recovery for both the `.rels` control and the `.RELS` variant, as word-valid now predicts for both.
- **#401: `word-content-type-invalid`.** An `invalid-*` type on an existing part makes Word refuse the file or offer recovery: styles, the main part and core properties. The same type on an `Override` for a missing part, or on a `Default` that covers no part, opens clean. The check over-reports.

## Not observed

Another session began driving Word by 02:28Z: its first log line is at 02:28:14Z, and its map file was written at 02:27:55Z. Its preflight closes every open document. To be safe, rows observed from 02:26Z on are not treated as evidence and are kept out of the table below. These specimens remain to be observed when Word is free:

- `351-escape-only-styles`: the styles part stored only as `word/%73tyles.xml`. It timed out twice.
- `351-backslash-only-main`: the main part stored only as `word\document.xml`. It timed out, then Word exited with no crash report.
- `398-hyperlink-rels-control` and `398-hyperlink-RELS`: an external hyperlink resolved through a `.rels` or a `.RELS` sidecar. The control timed out twice. The variant's open-error coincided with the other session's close-all.
- `397-encoding-unknown-{styles,content-types,package-rels,core}`: the scope of #397 beyond the main part. Not yet run.

## Observations

| Specimen                                 | Issue   | SHA-256 (first 16) | Word           | word-valid                                                                                                                  | strict  |
| ---------------------------------------- | ------- | ------------------ | -------------- | --------------------------------------------------------------------------------------------------------------------------- | ------- |
| `control-minimal`                        | control | `e3fd9a9607c643b1` | clean-open     | valid                                                                                                                       | valid   |
| `388-rels-zero-byte`                     | #388    | `518a110444455811` | repair-warning | valid                                                                                                                       | invalid |
| `388-rels-truncated`                     | #388    | `a85ad66282671910` | repair-warning | valid                                                                                                                       | invalid |
| `388-rels-undeclared-prefix`             | #388    | `19383610c2a9f89a` | repair-warning | valid                                                                                                                       | invalid |
| `392-WORD-folder-clean`                  | #392    | `36d4625db191d3e9` | clean-open     | valid                                                                                                                       | valid   |
| `392-WORD-folder-malformed-main`         | #392    | `b5c3d2178d4696be` | open-error     | invalid (`xml-undeclared-prefix@WORD/document.xml`, `xml-syntax@WORD/document.xml`)                                         | invalid |
| `392-word-folder-malformed-main-control` | #392    | `95d4f0bfb82a2c4a` | open-error     | invalid (`xml-undeclared-prefix@word/document.xml`, `xml-syntax@word/document.xml`)                                         | invalid |
| `396-CT-upper-clean`                     | #396    | `744bbb14ad7f0f3e` | clean-open     | valid                                                                                                                       | valid   |
| `396-CT-upper-invalid-type`              | #396    | `ec74a991d4b8b9dc` | clean-open     | invalid (`word-content-type-invalid@[CONTENT_TYPES].XML`)                                                                   | invalid |
| `396-CT-exact-invalid-type-control`      | #396    | `9857e3aaebaa53d7` | clean-open     | invalid (`word-content-type-invalid@[Content_Types].xml`)                                                                   | invalid |
| `396-CT-upper-malformed`                 | #396    | `949e38e644ebd7db` | repair-warning | invalid (`xml-syntax@[CONTENT_TYPES].XML`)                                                                                  | invalid |
| `397-encoding-unknown`                   | #397    | `a45331da8b2129ea` | open-error     | valid                                                                                                                       | invalid |
| `397-encoding-windows-1252-control`      | #397    | `0e8165d11d1898b5` | clean-open     | valid                                                                                                                       | valid   |
| `351-dup-case-identical`                 | #351    | `48a8b62b9c83535d` | repair-warning | invalid (`package-extract-failed@-`)                                                                                        | invalid |
| `351-dup-case-lower-good-upper-bad`      | #351    | `2d58ed5a3241a634` | repair-warning | invalid (`package-extract-failed@-`)                                                                                        | invalid |
| `351-dup-case-upper-bad-first`           | #351    | `8ef60be49ff4ac55` | repair-warning | invalid (`package-extract-failed@-`)                                                                                        | invalid |
| `351-dup-stream-good-then-bad`           | #351    | `970eca474f9eba1c` | repair-warning | invalid (`package-extract-failed@-`)                                                                                        | invalid |
| `351-dup-stream-bad-then-good`           | #351    | `b503d9e06055091e` | repair-warning | invalid (`package-extract-failed@-`)                                                                                        | invalid |
| `398-RELS-sidecar-clean`                 | #398    | `ac61b47790dc028b` | clean-open     | valid                                                                                                                       | valid   |
| `398-dup-theme-control`                  | #398    | `52b18dea69dafe56` | repair-warning | invalid (`word-duplicate-theme-relationship@word/_rels/document.xml.rels`)                                                  | invalid |
| `398-dup-theme-RELS`                     | #398    | `f1a5a570c9effd78` | repair-warning | invalid (`word-duplicate-theme-relationship@word/_rels/document.xml.RELS`, `word-hyperlink-relationship@word/document.xml`) | invalid |
| `cti-override-existing-styles`           | cti     | `dde1b233d9672e0f` | repair-warning | invalid (`word-content-type-invalid@[Content_Types].xml`)                                                                   | invalid |
| `cti-override-existing-main`             | cti     | `bded55854c3edf59` | repair-warning | invalid (`word-content-type-invalid@[Content_Types].xml`)                                                                   | invalid |
| `cti-override-existing-core`             | cti     | `f4081445c08647f1` | open-error     | invalid (`word-content-type-invalid@[Content_Types].xml`)                                                                   | invalid |
| `cti-override-existing-core-control`     | cti     | `86df51ceafbcb5ba` | clean-open     | valid                                                                                                                       | valid   |
| `cti-default-unused-extension`           | cti     | `d19b131b7b46dbba` | clean-open     | invalid (`word-content-type-invalid@[Content_Types].xml`)                                                                   | invalid |
| `351-dup-escape-identical`               | #351    | `f5cd33350d5e730e` | repair-warning | valid                                                                                                                       | invalid |
| `351-dup-backslash-identical`            | #351    | `348701aaaa331f41` | repair-warning | valid                                                                                                                       | valid   |

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
```

</details>

<details><summary>Probe script</summary>

Run from the repository root as a copy: `node --import tsx ./word-probe.tmp.mts [specimen …]`.

```ts
// Real-Word probes for #388, #392, #396, #397, #351, #398. Run from the stack worktree root as a copy.
import { createHash } from "node:crypto";
import { appendFileSync, existsSync, readFileSync, writeFileSync } from "node:fs";
import path from "node:path";
import JSZip from "jszip";
import { MINIMAL_PARTS, buildDocx } from "./tests/support/minimal-docx";
import { validate } from "./src/scripts/office/validate";
import { verifyInWord } from "./src/lib/word-verifier";

const OUT = "/Users/arthrod/temp/T/wfv-scratch/word-probes";
const RESULTS = path.join(OUT, "results.jsonl");
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
        "398-hyperlink-RELS",
        "#398",
        () =>
            buildDocx({
                ...without({ ...P, "word/document.xml": docH }, "word/_rels/document.xml.rels"),
                "word/_rels/document.xml.RELS": relsH,
            }),
    ],
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
    const word = await verifyInWord(new Uint8Array(bytes));
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
    appendFileSync(RESULTS, JSON.stringify(rec) + "\n");
    console.log(`${name}: word=${word.outcome} | word-valid=${wv.valid} ${wv.errors.join(",")} | strict=${st.valid}`);
}
```

</details>
