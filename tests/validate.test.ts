/*
 * Copyright 2026 Jandira Technologies, LLC
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *     http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */

import { promises as fs } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

import JSZip from "jszip";
import { describe, expect, it } from "vitest";

import { withTempDir } from "../src/lib/run-cli";
import { runValidateFromArgv, validate } from "../src/scripts/office/validate";
import { buildDocx, MINIMAL_PARTS } from "./support/minimal-docx";
import { rels, relation } from "./support/word-rule-test-xml";
import { BaseSchemaValidator } from "../src/scripts/office/validators/base";

const HERE = path.dirname(fileURLToPath(import.meta.url));
const FIXTURES = path.join(HERE, "fixtures");
const BROKEN_DIR = path.join(HERE, "fixtures", "broken");

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';

async function writeMinimalDocxDir(root: string, body: string): Promise<void> {
    await fs.mkdir(path.join(root, "_rels"), { recursive: true });
    await fs.mkdir(path.join(root, "word"), { recursive: true });
    await fs.writeFile(
        path.join(root, "[Content_Types].xml"),
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
            '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
            '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
            '<Default Extension="xml" ContentType="application/xml"/>' +
            '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
            "</Types>",
        "utf8",
    );
    await fs.writeFile(
        path.join(root, "_rels", ".rels"),
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>' +
            "</Relationships>",
        "utf8",
    );
    await fs.writeFile(
        path.join(root, "word", "document.xml"),
        `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:document ${W}>${body}</w:document>`,
        "utf8",
    );
}

describe("validate", () => {
    it("throws when the target path does not exist", async () => {
        await withTempDir(async (dir) => {
            await expect(validate(path.join(dir, "missing.docx"))).rejects.toThrow(/does not exist/);
        });
    });

    it("throws when the original is not a real file", async () => {
        await withTempDir(async (dir) => {
            const docxPath = path.join(BROKEN_DIR, "endnotes.paraid-overflow.docx");
            await expect(validate(docxPath, { original: path.join(dir, "missing.docx") })).rejects.toThrow(/is not a file/);
        });
    });

    it("throws when extension cannot be determined and no --original is given", async () => {
        await withTempDir(async (dir) => {
            // Plain directory with no extension and no --original — Python's
            // assertion about determining file type kicks in.
            await expect(validate(dir)).rejects.toThrow(/Cannot determine file type/);
        });
    });

    it("flags malformed XML in a packed DOCX", async () => {
        const docxPath = path.join(BROKEN_DIR, "single-paragraph.malformed-xml.docx");
        const result = await validate(docxPath);
        expect(result.suffix).toBe(".docx");
        expect(result.valid).toBe(false);
        expect(result.issues.some((i) => i.severity === "error")).toBe(true);
    });

    it("does not throw on a non-ZIP / encrypted / corrupted package; returns an error verdict", async () => {
        await withTempDir(async (dir) => {
            // Encrypted OOXML is a CFB/OLE compound file (D0CF11E0 magic), not a
            // ZIP; a truncated/corrupted .docx fails the same way. JSZip throws
            // "Can't find end of central directory". validate() must surface that
            // as a structured error issue, never crash — Word rejects these files,
            // so they must read as invalid (this profile must keep it blocking).
            const bad = path.join(dir, "encrypted.docx");
            await fs.writeFile(bad, Buffer.from([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1, 0x00, 0x00, 0x00, 0x00]));
            const result = await validate(bad, { profile: "word-valid" });
            expect(result.suffix).toBe(".docx");
            expect(result.valid).toBe(false);
            expect(result.issues.some((i) => i.severity === "error" && i.code === "package-open-failed")).toBe(true);
        });
    });

    it("does not label a zip-slip extract failure as package-open-failed", async () => {
        await withTempDir(async (dir) => {
            const zip = new JSZip();
            zip.file(
                "[Content_Types].xml",
                '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>',
            );
            zip.file(
                "word/document.xml",
                '<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>',
            );
            const buf = await zip.generateAsync({ type: "nodebuffer" });
            const docxPath = path.join(dir, "slip.docx");
            await fs.writeFile(docxPath, buf);

            const originalLoadAsync = JSZip.loadAsync;
            JSZip.loadAsync = async function loadAsyncPatched(data, options) {
                const loaded = await originalLoadAsync.call(this, data, options);
                loaded.files["../../escaped.txt"] = {
                    name: "../../escaped.txt",
                    dir: false,
                    async: async () => Buffer.from("malicious"),
                } as unknown as JSZip.JSZipObject;
                return loaded;
            };
            try {
                const result = await validate(docxPath);
                expect(result.valid).toBe(false);
                expect(result.issues.some((i) => i.code === "package-open-failed")).toBe(false);
                expect(result.issues.some((i) => i.code === "package-extract-failed")).toBe(true);
            } finally {
                JSZip.loadAsync = originalLoadAsync;
            }
        });
    });

    it("flags paraId overflow in a packed DOCX (negative case)", async () => {
        const docxPath = path.join(BROKEN_DIR, "endnotes.paraid-overflow.docx");
        const result = await validate(docxPath, { profile: "lenient" });
        expect(result.valid).toBe(false);
        expect(result.issues.some((i) => i.severity === "error" && i.code === "id-paraid-overflow")).toBe(true);
    });

    it("word-valid profile downgrades Word-tolerated paraId overflow", async () => {
        const docxPath = path.join(BROKEN_DIR, "endnotes.paraid-overflow.docx");
        const result = await validate(docxPath, { profile: "word-valid" });
        expect(result.valid).toBe(true);
        expect(result.issues.some((i) => i.severity === "warning" && i.code === "id-paraid-overflow")).toBe(true);
    });

    it("word-valid profile keeps commentsIds/commentsExtensible mismatches fatal", async () => {
        const docxPath = path.join(BROKEN_DIR, "sample-document.broken-tables.docx");
        const result = await validate(docxPath, { profile: "word-valid" });
        expect(result.valid).toBe(false);
        expect(result.issues.some((i) => i.severity === "error" && i.code === "comment-thread-commentid-paraid-orphan")).toBe(true);
        expect(result.issues.some((i) => i.severity === "error" && i.code === "comment-thread-durableid-orphan")).toBe(true);
    });

    it("word-valid profile flags body-level m:oMathPara with m:sPre", async () => {
        const docxPath = path.join(FIXTURES, "external", "superdoc", "behavior", "math-spre-tests.docx");
        const lenient = await validate(docxPath, { profile: "lenient" });
        const wordValid = await validate(docxPath, { profile: "word-valid" });
        expect(lenient.valid).toBe(true);
        expect(wordValid.valid).toBe(false);
        expect(wordValid.issues.some((i) => i.severity === "error" && i.code === "word-math-spre-body")).toBe(true);
    });

    it("word-valid profile flags invalid content types that Word refuses", async () => {
        const bad = path.join(FIXTURES, "external", "open-xml-sdk", "InvalidDocPropsct.docx");
        const tolerated = path.join(FIXTURES, "external", "open-xml-sdk", "InvalidDocProps.docx");
        const badResult = await validate(bad, { profile: "word-valid" });
        const toleratedResult = await validate(tolerated, { profile: "word-valid" });
        expect(badResult.valid).toBe(false);
        expect(badResult.issues.some((i) => i.severity === "error" && i.code === "word-content-type-invalid")).toBe(true);
        expect(toleratedResult.valid).toBe(true);
    });

    it("auto-repairs missing xml:space='preserve' when --auto-repair is set", async () => {
        await withTempDir(async (tmp) => {
            const unpacked = path.join(tmp, "unpacked");
            // <w:t> with leading whitespace but no xml:space="preserve" — the
            // base validator's repairWhitespacePreservation pass should add it.
            await writeMinimalDocxDir(unpacked, "<w:body><w:p><w:r><w:t> hello</w:t></w:r></w:p></w:body>");

            // Synthesise a tiny "original" so dispatch picks the docx path even
            // though our target is a bare directory with no extension.
            const originalDocx = path.join(tmp, "orig.docx");
            const zip = new JSZip();
            zip.file(
                "[Content_Types].xml",
                '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
                    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
                    '<Default Extension="xml" ContentType="application/xml"/>' +
                    "</Types>",
            );
            zip.file(
                "word/document.xml",
                `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:document ${W}><w:body/></w:document>`,
            );
            await fs.writeFile(originalDocx, await zip.generateAsync({ type: "nodebuffer" }));

            const before = await validate(unpacked, {
                autoRepair: false,
                original: originalDocx,
                author: "Test Author",
            });
            const wsBefore = before.issues.filter((i) => i.code === "ws-missing-preserve");
            expect(wsBefore.length).toBeGreaterThan(0);

            const after = await validate(unpacked, {
                autoRepair: true,
                original: originalDocx,
                author: "Test Author",
            });
            expect(after.repairs).toBeGreaterThanOrEqual(1);
            const wsAfter = after.issues.filter((i) => i.code === "ws-missing-preserve");
            expect(wsAfter.length).toBe(0);
            expect(
                after.issues.some(
                    (i) => i.code === "repair-plan" && i.path === "word/document.xml" && i.message.includes("add xml:space='preserve'"),
                ),
            ).toBe(true);
            expect(after.issues.some((i) => i.code === "repair-content-preserved")).toBe(true);

            // Sanity-check the on-disk file actually got the attribute added.
            const docXml = await fs.readFile(path.join(unpacked, "word", "document.xml"), "utf8");
            expect(docXml).toMatch(/xml:space="preserve"/);
        });
    });

    it("dispatches by --original suffix when target is a bare directory", async () => {
        await withTempDir(async (tmp) => {
            const unpacked = path.join(tmp, "unpacked");
            await writeMinimalDocxDir(unpacked, "<w:body><w:p><w:r><w:t>hi</w:t></w:r></w:p></w:body>");

            // Synthesise a tiny DOCX to act as the "original".
            const originalDocx = path.join(tmp, "orig.docx");
            const zip = new JSZip();
            zip.file(
                "[Content_Types].xml",
                '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
                    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
                    '<Default Extension="xml" ContentType="application/xml"/>' +
                    '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
                    "</Types>",
            );
            zip.file(
                "word/document.xml",
                `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:document ${W}><w:body><w:p><w:r><w:t>hi</w:t></w:r></w:p></w:body></w:document>`,
            );
            await fs.writeFile(originalDocx, await zip.generateAsync({ type: "nodebuffer" }));

            const result = await validate(unpacked, {
                original: originalDocx,
                author: "Test Author",
            });
            expect(result.suffix).toBe(".docx");
            // Body of the test isn't asserting validity (there's no schemas dir
            // alignment to guarantee that here); the contract under test is that
            // the docx dispatch path was selected and produced a structured result.
            expect(typeof result.valid).toBe("boolean");
        });
    });
});

describe("xlsx unsupported file type", () => {
    it("returns a structured ValidationResult with unsupported-file-type rather than throwing", async () => {
        await withTempDir(async (tmp) => {
            // Create a minimal xlsx-named zip so it passes the file-exists and
            // is-a-file checks; actual xlsx content is irrelevant — the validator
            // bails out before inspecting the XML.
            const xlsxPath = path.join(tmp, "workbook.xlsx");
            const zip = new JSZip();
            zip.file("xl/workbook.xml", '<?xml version="1.0" encoding="UTF-8"?><workbook/>');
            await fs.writeFile(xlsxPath, await zip.generateAsync({ type: "nodebuffer" }));

            const result = await validate(xlsxPath);
            expect(result.valid).toBe(false);
            expect(result.suffix).toBe(".xlsx");
            expect(result.issues.some((i) => i.code === "unsupported-file-type")).toBe(true);
        });
    });

    it("runValidateFromArgv exits 1 for xlsx without throwing", async () => {
        await withTempDir(async (tmp) => {
            const xlsxPath = path.join(tmp, "workbook.xlsx");
            const zip = new JSZip();
            zip.file("xl/workbook.xml", '<?xml version="1.0" encoding="UTF-8"?><workbook/>');
            await fs.writeFile(xlsxPath, await zip.generateAsync({ type: "nodebuffer" }));

            // Capture the CLI's stderr issue-render output so the
            // "ERROR [...]: Unsupported file type: .xlsx" line — which is
            // the *expected* CLI behaviour for a structurally-failed
            // validation — does not leak into the test runner's stdout.
            const captured: string[] = [];
            const origWrite = process.stderr.write.bind(process.stderr);
            // biome-ignore lint/suspicious/noExplicitAny: vitest stub typing
            process.stderr.write = ((chunk: string | Uint8Array): boolean => {
                if (typeof chunk === "string") captured.push(chunk);
                return true;
            }) as typeof process.stderr.write;
            let exitCode: number;
            try {
                exitCode = await runValidateFromArgv([xlsxPath]);
            } finally {
                process.stderr.write = origWrite;
            }
            expect(exitCode).toBe(1);
            // Pin the CLI's expected stderr line so a regression that
            // silenced it would also fail this assertion, not just hide
            // behind quiet output.
            expect(captured.join("")).toContain("Unsupported file type: .xlsx");
        });
    });
});

describe("word-valid profile — applyWordValidProfile coverage", () => {
    it("downgrades ws-missing-preserve errors to warnings under word-valid profile (Issue H)", async () => {
        // A <w:t> with leading whitespace and no xml:space="preserve" triggers
        // ws-missing-preserve as an error in all profiles. Under word-valid the
        // error should be downgraded to a warning because it is not Word-blocking.
        await withTempDir(async (dir) => {
            const unpacked = path.join(dir, "unpacked");
            await writeMinimalDocxDir(unpacked, "<w:body><w:p><w:r><w:t> hello</w:t></w:r></w:p></w:body>");
            const result = await validate(unpacked, {
                profile: "word-valid",
                original: path.join(BROKEN_DIR, "endnotes.paraid-overflow.docx"),
                author: "Test",
            });
            expect(result.valid).toBe(true);
            const wsIssue = result.issues.find((i) => i.code === "ws-missing-preserve");
            expect(wsIssue).toBeDefined();
            expect(wsIssue?.severity).toBe("warning");
        });
    });

    it("flags Default element with invalid- ContentType under word-valid profile (Issue J)", async () => {
        // validateWordContentTypes checks both Override and Default elements.
        // The existing fixture test covers Override; this covers Default.
        await withTempDir(async (dir) => {
            const unpacked = path.join(dir, "unpacked");
            await writeMinimalDocxDir(unpacked, "<w:body><w:p/></w:body>");
            // Overwrite [Content_Types].xml to add a Default with invalid- ContentType
            await fs.writeFile(
                path.join(unpacked, "[Content_Types].xml"),
                '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
                    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
                    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
                    '<Default Extension="json" ContentType="invalid-application/json"/>' +
                    '<Default Extension="xml" ContentType="application/xml"/>' +
                    '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
                    "</Types>",
                "utf8",
            );
            const result = await validate(unpacked, {
                profile: "word-valid",
                original: path.join(BROKEN_DIR, "endnotes.paraid-overflow.docx"),
                author: "Test",
            });
            expect(result.valid).toBe(false);
            expect(result.issues.some((i) => i.severity === "error" && i.code === "word-content-type-invalid")).toBe(true);
        });
    });

    it("keeps rels-broken error for ../customXml/ targets under word-valid profile (Issue K)", async () => {
        // rels-broken referencing ../customXml/ is Word-blocking (custom XML parts
        // cause Word to refuse the file), so it must stay as an error under word-valid.
        await withTempDir(async (dir) => {
            const unpacked = path.join(dir, "unpacked");
            await writeMinimalDocxDir(unpacked, "<w:body><w:p/></w:body>");
            // Create a word/_rels/document.xml.rels that points to a missing customXml part
            await fs.mkdir(path.join(unpacked, "word", "_rels"), { recursive: true });
            await fs.writeFile(
                path.join(unpacked, "word", "_rels", "document.xml.rels"),
                '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
                    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
                    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml" Target="../customXml/item1.xml"/>' +
                    "</Relationships>",
                "utf8",
            );
            const result = await validate(unpacked, {
                profile: "word-valid",
                original: path.join(BROKEN_DIR, "endnotes.paraid-overflow.docx"),
                author: "Test",
            });
            // The rels-broken issue for ../customXml/ must remain an error, not be downgraded.
            const relsIssue = result.issues.find((i) => i.code === "rels-broken" && i.message.includes("../customXml/"));
            expect(relsIssue).toBeDefined();
            expect(relsIssue?.severity).toBe("error");
        });
    });

    it("keeps rels-empty-element error for missing required attributes under word-valid profile (Issue L)", async () => {
        // A <Relationship> missing a required attribute (Id, Type, or Target) is Word-blocking
        // and must remain an error under word-valid — only the non-self-closing variant
        // (body content) is downgraded to a warning.
        await withTempDir(async (dir) => {
            const unpacked = path.join(dir, "unpacked");
            await writeMinimalDocxDir(unpacked, "<w:body><w:p/></w:body>");
            await fs.mkdir(path.join(unpacked, "word", "_rels"), { recursive: true });
            // Relationship is missing the required Id attribute
            await fs.writeFile(
                path.join(unpacked, "word", "_rels", "document.xml.rels"),
                '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
                    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
                    '<Relationship Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>' +
                    "</Relationships>",
                "utf8",
            );
            const result = await validate(unpacked, {
                profile: "word-valid",
                original: path.join(BROKEN_DIR, "endnotes.paraid-overflow.docx"),
                author: "Test",
            });
            const emptyElemIssue = result.issues.find(
                (i) => i.code === "rels-empty-element" && i.message.includes("missing required attribute"),
            );
            expect(emptyElemIssue).toBeDefined();
            expect(emptyElemIssue?.severity).toBe("error");
        });
    });

    it("word-valid profile emits no word-math-parse when word/document.xml is absent (Issue I)", async () => {
        // validateWordOpenCompatibility() returns early when document.xml is not found,
        // so the word-math-parse code must not appear in the result.
        await withTempDir(async (dir) => {
            const unpacked = path.join(dir, "unpacked");
            // Build the minimal directory structure but deliberately omit word/document.xml.
            await fs.mkdir(path.join(unpacked, "_rels"), { recursive: true });
            await fs.mkdir(path.join(unpacked, "word"), { recursive: true });
            await fs.writeFile(
                path.join(unpacked, "[Content_Types].xml"),
                '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
                    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
                    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
                    '<Default Extension="xml" ContentType="application/xml"/>' +
                    "</Types>",
                "utf8",
            );
            await fs.writeFile(
                path.join(unpacked, "_rels", ".rels"),
                '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
                    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
                    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>' +
                    "</Relationships>",
                "utf8",
            );
            const result = await validate(unpacked, {
                profile: "word-valid",
                original: path.join(BROKEN_DIR, "endnotes.paraid-overflow.docx"),
                author: "Test",
            });
            expect(result.issues.every((i) => i.code !== "word-math-parse")).toBe(true);
        });
    });

    it("word-valid profile reports a malformed word/document.xml once, as xml-syntax (Issue M)", async () => {
        // validateWordOpenCompatibility cannot parse the part and emits
        // word-math-parse. The XML diagnostics already reject the part, so
        // that consequential parse failure is dropped and the part is
        // reported once.
        await withTempDir(async (dir) => {
            const unpacked = path.join(dir, "unpacked");
            await writeMinimalDocxDir(unpacked, "<w:body><w:p/></w:body>");
            // Replace document.xml with content that cannot be parsed as XML.
            await fs.writeFile(path.join(unpacked, "word", "document.xml"), "this is definitely <<< not valid XML >>>", "utf8");
            const result = await validate(unpacked, {
                profile: "word-valid",
                original: path.join(BROKEN_DIR, "endnotes.paraid-overflow.docx"),
                author: "Test",
            });
            expect(result.valid).toBe(false);
            const onDocument = result.issues.filter((i) => i.path === "word/document.xml");
            expect(onDocument.length).toBeGreaterThan(0);
            expect(new Set(onDocument.map((i) => i.code))).toEqual(new Set(["xml-syntax"]));
            expect(onDocument.every((i) => i.severity === "error")).toBe(true);
        });
    });

    it("xml-syntax error in word/ file stays fatal under word-valid profile (Issue N)", async () => {
        // isWordBlockingIssue returns true for xml-syntax when path starts with 'word/',
        // so the error must NOT be downgraded to a warning by applyWordValidProfile.
        await withTempDir(async (dir) => {
            const unpacked = path.join(dir, "unpacked");
            await writeMinimalDocxDir(unpacked, "<w:body><w:p/></w:body>");
            // Write an unparseable word/styles.xml so validateXml emits xml-syntax for it.
            await fs.writeFile(path.join(unpacked, "word", "styles.xml"), "this is <definitely not valid xml", "utf8");
            const result = await validate(unpacked, {
                profile: "word-valid",
                original: path.join(BROKEN_DIR, "endnotes.paraid-overflow.docx"),
                author: "Test",
            });
            const xmlSyntaxIssue = result.issues.find((i) => i.code === "xml-syntax" && (i.path?.startsWith("word/") ?? false));
            expect(xmlSyntaxIssue).toBeDefined();
            expect(xmlSyntaxIssue?.severity).toBe("error");
        });
    });

    it("xml-syntax error outside word/ is downgraded to warning under word-valid profile (Issue O)", async () => {
        // isWordBlockingIssue returns false for xml-syntax when path does NOT start with
        // 'word/', so applyWordValidProfile must downgrade it to a warning.
        await withTempDir(async (dir) => {
            const unpacked = path.join(dir, "unpacked");
            await writeMinimalDocxDir(unpacked, "<w:body><w:p/></w:body>");
            // Write unparseable XML in docProps/ (outside word/) to trigger xml-syntax there.
            await fs.mkdir(path.join(unpacked, "docProps"), { recursive: true });
            await fs.writeFile(path.join(unpacked, "docProps", "app.xml"), "this is <definitely not valid xml", "utf8");
            const result = await validate(unpacked, {
                profile: "word-valid",
                original: path.join(BROKEN_DIR, "endnotes.paraid-overflow.docx"),
                author: "Test",
            });
            const xmlSyntaxIssue = result.issues.find((i) => i.code === "xml-syntax" && (i.path?.startsWith("docProps/") ?? false));
            expect(xmlSyntaxIssue).toBeDefined();
            expect(xmlSyntaxIssue?.severity).toBe("warning");
            // word-valid remains valid because the only error was downgraded.
            expect(result.valid).toBe(true);
        });
    });

    it("xml-undeclared-prefix in word/ stays fatal under word-valid, outside word/ it is a warning", async () => {
        // An undeclared namespace prefix is the same class of defect as a
        // malformed part (Word reports unreadable content), so it follows the
        // same word/ vs non-word/ policy as xml-syntax.
        await withTempDir(async (dir) => {
            const unpacked = path.join(dir, "unpacked");
            await writeMinimalDocxDir(unpacked, "<w:body><w:p/></w:body>");
            await fs.writeFile(path.join(unpacked, "word", "comments.xml"), '<w:comments><w:comment w:id="0"/></w:comments>', "utf8");
            await fs.mkdir(path.join(unpacked, "docProps"), { recursive: true });
            await fs.writeFile(path.join(unpacked, "docProps", "custom.xml"), "<Properties><cp:Thing/></Properties>", "utf8");

            const result = await validate(unpacked, {
                profile: "word-valid",
                original: path.join(BROKEN_DIR, "endnotes.paraid-overflow.docx"),
                author: "Test",
            });

            const wordPart = result.issues.find((i) => i.code === "xml-undeclared-prefix" && (i.path?.startsWith("word/") ?? false));
            expect(wordPart).toBeDefined();
            expect(wordPart?.severity).toBe("error");
            const docPropsPart = result.issues.find(
                (i) => i.code === "xml-undeclared-prefix" && (i.path?.startsWith("docProps/") ?? false),
            );
            expect(docPropsPart).toBeDefined();
            expect(docPropsPart?.severity).toBe("warning");
        });
    });
});

describe("validate startup probe", () => {
    it("should have a working wasm-backed XSD engine", async () => {
        // The CLI calls this at the top of `runValidateFromArgv` so a broken
        // libxml2-wasm engine fails loudly instead of degrading into
        // per-file XSD errors.
        await expect(BaseSchemaValidator.assertXsdValidationAvailable()).resolves.toBeUndefined();
    });
});

// Packed integration regressions, all specimens generated without fixture edits.
describe("packed OPC-equivalent relationships parts (#398)", () => {
    it.each([
        ["theme", false],
        ["theme", true],
        ["fontTable", false],
        ["fontTable", true],
    ] as const)("rejects duplicate %s relationships with selected=%s in Word-valid", async (type, selected) => {
        await withTempDir(async (dir) => {
            const check = type === "theme" ? "word-duplicate-theme-relationship" : "word-duplicate-font-table-relationship";
            const checks = selected ? [check] : undefined;
            const run = async (extension: string) => {
                const parts = { ...MINIMAL_PARTS };
                delete parts["word/_rels/document.xml.rels"];
                const part = `word/_rels/document.xml.${extension}`;
                parts[part] = rels(
                    relation("styles", "styles", "", "styles.xml") +
                        relation("themeA", type, "", "styles.xml") +
                        relation("themeB", type, "", "styles.xml"),
                );
                const file = path.join(dir, `${extension}.docx`);
                const bytes = await buildDocx(parts);
                await fs.writeFile(file, bytes);
                const result = await validate(file, { profile: "word-valid", checks });
                expect(await fs.readFile(file)).toEqual(bytes);
                return result;
            };
            const control = await run("rels");
            expect(control.valid).toBe(false);
            expect(control.issues).toContainEqual(
                expect.objectContaining({ code: check, severity: "error", path: "word/_rels/document.xml.rels" }),
            );
            for (const extension of ["RELS", "%72els", "%52ELS"]) {
                const variant = await run(extension);
                expect(variant.valid).toBe(control.valid);
                expect(variant.issues).toEqual(
                    control.issues.map((issue) => ({
                        ...issue,
                        ...(issue.path === "word/_rels/document.xml.rels" ? { path: `word/_rels/document.xml.${extension}` } : {}),
                        message: issue.message.replaceAll("document.xml.rels", `document.xml.${extension}`),
                    })),
                );
            }
        });
    });

    it.each(["RELS", "%72els", "%52ELS"])("rejects an empty notes.%s in a full strict packed run", async (extension) => {
        await withTempDir(async (dir) => {
            const run = async (ext: string) => {
                const file = path.join(dir, `${ext}.docx`);
                const bytes = await buildDocx({ ...MINIMAL_PARTS, [`word/notes.${ext}`]: rels("") });
                await fs.writeFile(file, bytes);
                const result = await validate(file, { profile: "strict" });
                expect(await fs.readFile(file)).toEqual(bytes);
                return result;
            };
            const control = await run("rels");
            expect(control.valid).toBe(false);
            expect(control.issues).toContainEqual(
                expect.objectContaining({ code: "rels-empty-part", severity: "error", path: "word/notes.rels" }),
            );
            const variant = await run(extension);
            expect(variant.valid).toBe(control.valid);
            expect(variant.issues).toEqual(
                control.issues.map((issue) => ({
                    ...issue,
                    ...(issue.path === "word/notes.rels" ? { path: `word/notes.${extension}` } : {}),
                })),
            );
        });
    });

    it("does not introduce findings when the content-types stream is stored as [CONTENT_TYPES].XML", async () => {
        await withTempDir(async (dir) => {
            for (const profile of ["word-valid", "lenient", "strict"] as const) {
                const run = async (name: string) => {
                    const parts = { ...MINIMAL_PARTS };
                    delete parts["[Content_Types].xml"];
                    parts[name] = MINIMAL_PARTS["[Content_Types].xml"];
                    const file = path.join(dir, `${profile}-${name}.docx`);
                    await fs.writeFile(file, await buildDocx(parts));
                    return validate(file, { profile });
                };
                const control = await run("[Content_Types].xml");
                const variant = await run("[CONTENT_TYPES].XML");
                // The stream is found in any case (#396), so the expanded XML
                // list (#398) adds no finding for it.
                expect(variant.issues).toEqual(control.issues);
                expect(variant.valid).toBe(control.valid);
            }
        });
    });
});

describe("packed content-types stream names (#396)", () => {
    it.each([
        { profile: "word-valid", checks: undefined },
        { profile: "word-valid", checks: ["word-content-type-invalid"] },
        { profile: "lenient", checks: undefined },
    ] as const)("preserves the control verdict under $profile with checks=$checks", async (options) => {
        await withTempDir(async (dir) => {
            const types = MINIMAL_PARTS["[Content_Types].xml"]
                .replace('<Default Extension="xml" ContentType="application/xml"/>', "")
                .replace("</Types>", `<Override PartName="/word/x.xml" ContentType="invalid-type"/></Types>`);
            const { "[Content_Types].xml": _stream, ...otherParts } = MINIMAL_PARTS;
            const controlFile = path.join(dir, "control.docx");
            await fs.writeFile(controlFile, await buildDocx({ ...otherParts, "[Content_Types].xml": types }));
            const control = await validate(controlFile, options);
            const controlFindings = control.issues.filter((issue) => issue.code === "word-content-type-invalid");
            expect(controlFindings).toHaveLength(options.profile === "lenient" ? 0 : 1);
            for (const name of ["[CONTENT_TYPES].XML", "[content_types].xml"]) {
                const file = path.join(dir, "variant.docx");
                const bytes = await buildDocx({ ...otherParts, [name]: types });
                await fs.writeFile(file, bytes);
                const result = await validate(file, options);
                expect(result.valid).toBe(control.valid);
                expect(result.issues.filter((issue) => issue.code === "word-content-type-invalid")).toEqual(
                    controlFindings.map((issue) => ({ ...issue, path: name })),
                );
                expect(result.issues.filter((issue) => issue.code?.startsWith("ct-"))).toEqual(
                    control.issues.filter((issue) => issue.code?.startsWith("ct-")),
                );
                expect(await fs.readFile(file)).toEqual(bytes);
            }
        });
    });

    it.each(["[CONTENT_TYPES].XML", "[content_types].xml"])("dispatches stored stream %s to the OPC schema under strict", async (name) => {
        await withTempDir(async (dir) => {
            const { "[Content_Types].xml": types, ...parts } = MINIMAL_PARTS;
            // Missing PartName is an OPC XSD error; other parts remain valid.
            const invalid = types.replace("</Types>", `<Override ContentType="application/xml"/></Types>`);
            const file = path.join(dir, "variant.docx");
            const controlFile = path.join(dir, "control.docx");
            await fs.writeFile(file, await buildDocx({ ...parts, [name]: invalid }));
            await fs.writeFile(controlFile, await buildDocx({ ...parts, "[Content_Types].xml": invalid }));
            const control = await validate(controlFile, { profile: "strict" });
            const result = await validate(file, { profile: "strict" });
            expect(control.xsdWalk.parts).toContain("[Content_Types].xml");
            expect(result.xsdWalk.parts).toContain(name);
            expect(result.valid).toBe(false);
            const errors = (issues: typeof result.issues, stored: string) =>
                issues.filter((issue) => issue.path === stored && issue.code === "xsd-error");
            expect(errors(control.issues, "[Content_Types].xml")).not.toHaveLength(0);
            expect(errors(result.issues, name)).toEqual(
                errors(control.issues, "[Content_Types].xml").map((issue) => ({ ...issue, path: name })),
            );
        });
    });
});
