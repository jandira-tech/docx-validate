/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

import path from "node:path";
import { writeFileSync, mkdirSync } from "node:fs";

import { describe, expect, it, vi } from "vitest";

import { runChecks } from "../src/lib/check-scheduler";
import { withTempDir } from "../src/lib/run-cli";
import { ValidationContext } from "../src/lib/validation-context";
import { createDefaultCheckRegistry } from "../src/scripts/office/checks";
import { contextFromArchive, contextFromDirectory } from "../src/scripts/office/package-context";
import { buildDocx, MINIMAL_PARTS } from "./support/minimal-docx";

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const rowWithoutCell = MINIMAL_PARTS["word/document.xml"]!.replace(
    "<w:sectPr>",
    `<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="100"/></w:tblGrid><w:tr/></w:tbl><w:p/><w:sectPr>`,
);

describe("default check registry", () => {
    it("gives every check a unique id, a description, groups, and formats", () => {
        const registry = createDefaultCheckRegistry();
        const ids = registry.ids;
        expect(new Set(ids).size).toBe(ids.length);
        for (const def of registry.list()) {
            expect(def.description.length).toBeGreaterThan(10);
            expect(def.groups.length).toBeGreaterThan(0);
            expect(def.formats.length).toBeGreaterThan(0);
        }
        expect(ids).toEqual(
            expect.arrayContaining(["zip-bytes-after-eocd", "zip-eocd-comment-overrun", "xml-syntax", "word-row-without-cell"]),
        );
    });

    it("finds a row without a cell in an in-memory package through its named check", async () => {
        const ctx = ValidationContext.fromEntries({ ...MINIMAL_PARTS, "word/document.xml": rowWithoutCell });
        const run = await runChecks(ctx, createDefaultCheckRegistry(), { format: "docx", select: ["word-row-without-cell"] });
        expect(run.executed).toEqual(["xml-syntax", "word-row-without-cell"]);
        expect(run.issues.map((i) => [i.code, i.path])).toEqual([["word-row-without-cell", "word/document.xml"]]);
        expect(run.valid).toBe(false);
    });

    it("judges no malformed story XML instead of crashing on it", async () => {
        const ctx = ValidationContext.fromEntries({ ...MINIMAL_PARTS, "word/document.xml": `<w:document xmlns:w="${W}"><w:body>` });
        const run = await runChecks(ctx, createDefaultCheckRegistry(), { format: "docx", select: ["word-row-without-cell"] });
        // xml-syntax gates by part (#307): the table check runs, and the
        // malformed part is hidden from it.
        expect(run.outcomes.map((o) => [o.id, o.status])).toEqual([
            ["xml-syntax", "failed"],
            ["word-row-without-cell", "passed"],
        ]);
        expect(run.issues.map((i) => [i.code, i.path])).toEqual([["xml-syntax", "word/document.xml"]]);
    });

    it("reports content the DOM parser would recover from, with its line and column", async () => {
        const body = `<w:document xmlns:w="${W}">\n<w:body><w:p><w:r><w:t>a & b</w:t></w:r></w:p></w:body></w:document>`;
        const ctx = ValidationContext.fromEntries({ ...MINIMAL_PARTS, "word/document.xml": body });
        const run = await runChecks(ctx, createDefaultCheckRegistry(), { format: "docx", select: ["xml-syntax"] });
        expect(run.issues).toEqual([
            expect.objectContaining({
                severity: "error",
                code: "xml-syntax",
                path: "word/document.xml",
                line: 2,
                column: expect.any(Number),
            }),
        ]);
    });

    it("leaves non-XML parts to other checks, whatever their bytes", async () => {
        const ctx = ValidationContext.fromEntries({
            ...MINIMAL_PARTS,
            "word/media/image1.png": new Uint8Array([0x89, 0x50, 0x4e, 0x47, 0x26, 0x3c]),
        });
        const run = await runChecks(ctx, createDefaultCheckRegistry(), { format: "docx", select: ["xml-syntax"] });
        expect(run.issues).toEqual([]);
    });

    it("runs the ZIP envelope checks on archive bytes and reports one trailing byte", async () => {
        const docx = await buildDocx();
        const ctx = await ValidationContext.fromArchive(Buffer.concat([docx, Buffer.from([0x0a])]));
        const run = await runChecks(ctx, createDefaultCheckRegistry(), { format: "docx", select: ["package"] });
        expect(run.executed).toEqual(["zip-bytes-after-eocd", "zip-eocd-comment-overrun"]);
        expect(run.issues.map((i) => i.code)).toEqual(["zip-bytes-after-eocd"]);
    });

    it("reports a declared EOCD comment length past the end of the file even though the archive cannot be opened", async () => {
        const docx = Buffer.from(await buildDocx());
        docx.writeUInt16LE(8, docx.length - 2);
        const { context, openError } = await contextFromArchive(docx);
        expect(openError).toMatch(/Corrupted zip|End of data/i);
        expect(context.partNames).toEqual([]);
        const run = await runChecks(context, createDefaultCheckRegistry(), { format: "docx", select: ["package"] });
        expect(run.issues.map((i) => i.code)).toEqual(["zip-eocd-comment-overrun"]);
    });

    it("opens a readable archive without an error", async () => {
        const { context, openError } = await contextFromArchive(await buildDocx());
        expect(openError).toBeUndefined();
        expect(context.source).toBe("archive");
        expect(context.partNames).toContain("word/document.xml");
    });

    it("gates a DOM-reading rule behind the syntax check when selected alone", async () => {
        // Issue #298: the lenient DOM parser recovers from the unquoted
        // attribute value, so the shading rule itself reports nothing — only
        // the syntax prerequisite can surface the defect when the rule is
        // selected on its own.
        const ctx = ValidationContext.fromEntries({
            ...MINIMAL_PARTS,
            "word/document.xml": `<?xml version="1.0"?><w:document xmlns:w="${W}"><w:body><w:p unquoted=yes/></w:body></w:document>`,
        });
        const run = await runChecks(ctx, createDefaultCheckRegistry(), { format: "docx", select: ["word-shading-color"] });
        expect(run.executed).toContain("xml-syntax");
        expect(run.valid).toBe(false);
    });

    it("declares the xml-syntax prerequisite for every Word rule that reads a part's DOM", async () => {
        // Issue #298: a rule that reads a parsed part must be gated behind the
        // syntax check, or a part the lenient DOM parser recovers from silently
        // loses its findings. DOM reading is detected by spying on the shared
        // context, so a newly added rule cannot regress unnoticed.
        const domReaders: string[] = [];
        for (const def of createDefaultCheckRegistry().list("docx")) {
            if (def.id === "xml-syntax") continue;
            const ctx = ValidationContext.fromEntries(MINIMAL_PARTS);
            const xml = vi.spyOn(ctx, "xml");
            await def.run(ctx);
            if (xml.mock.calls.length === 0) {
                xml.mockRestore();
                continue;
            }
            domReaders.push(def.id);
            expect(def.requires ?? [], `${def.id} reads a part DOM, so it must require xml-syntax`).toContain("xml-syntax");
            xml.mockRestore();
        }
        // The probe package must actually exercise DOM reads, or this test
        // would pass vacuously.
        expect(domReaders).toEqual(expect.arrayContaining(["word-row-without-cell", "word-shading-color"]));
        expect(domReaders.length).toBeGreaterThan(2);
    });

    it("passes the clean minimal package", async () => {
        const ctx = await ValidationContext.fromArchive(await buildDocx());
        const run = await runChecks(ctx, createDefaultCheckRegistry(), { format: "docx" });
        expect(run.issues).toEqual([]);
        expect(run.valid).toBe(true);
        expect(run.omitted).toEqual([]);
    });

    it("marks the ZIP envelope checks unavailable for an unpacked directory", async () => {
        await withTempDir(async (dir) => {
            for (const [name, content] of Object.entries(MINIMAL_PARTS)) {
                mkdirSync(path.dirname(path.join(dir, name)), { recursive: true });
                writeFileSync(path.join(dir, name), content);
            }
            const ctx = contextFromDirectory(dir);
            expect(ctx.source).toBe("directory");
            expect(ctx.partNames).toEqual(Object.keys(MINIMAL_PARTS).sort());
            const run = await runChecks(ctx, createDefaultCheckRegistry(), { format: "docx" });
            expect(run.unavailable).toEqual(["zip-bytes-after-eocd", "zip-eocd-comment-overrun"]);
            expect(run.valid).toBe(true);
        });
    });

    it("reads a directory part lazily and only once", async () => {
        await withTempDir(async (dir) => {
            mkdirSync(path.join(dir, "word"), { recursive: true });
            writeFileSync(path.join(dir, "word", "document.xml"), MINIMAL_PARTS["word/document.xml"]!);
            const ctx = contextFromDirectory(dir);
            expect(ctx.readCount("word/document.xml")).toBe(0);
            ctx.xml("word/document.xml");
            ctx.text("word/document.xml");
            expect(ctx.readCount("word/document.xml")).toBe(1);
        });
    });
});
