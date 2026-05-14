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

    it("flags paraId overflow in a packed DOCX (negative case)", async () => {
        const docxPath = path.join(BROKEN_DIR, "endnotes.paraid-overflow.docx");
        const result = await validate(docxPath);
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

describe("validate startup probe", () => {
    it("BaseSchemaValidator.assertLibxmljsAvailable() does not throw on this host", () => {
        // The CLI calls this at the top of `runValidateFromArgv` so a broken libxmljs2
        // binding fails loudly instead of degrading into per-file XSD errors.
        expect(() => BaseSchemaValidator.assertLibxmljsAvailable()).not.toThrow();
    });
});
