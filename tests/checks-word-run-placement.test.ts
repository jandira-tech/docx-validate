/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

import { describe, expect, it } from "vitest";

import { ValidationContext } from "../src/lib/validation-context";
import {
    RUN_PARENTS,
    RUN_PROPERTIES_PARENTS,
    UNDERLINE_PARENTS,
    wordRunMisplacedCheck as runCheck,
    wordRunPropertiesMisplacedCheck as runPropertiesCheck,
    wordUnderlineMisplacedCheck as underlineCheck,
} from "../src/scripts/office/checks/word-run-placement";

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const M = "http://schemas.openxmlformats.org/officeDocument/2006/math";
const R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

const document = (body: string) =>
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<w:document xmlns:w="${W}" xmlns:m="${M}" xmlns:r="${R}"><w:body>${body}</w:body></w:document>`;

const styles = (body: string) => `<?xml version="1.0"?><w:styles xmlns:w="${W}">${body}</w:styles>`;

describe("word-run-misplaced", () => {
    it("reports a run directly inside a table cell, as in word-rejects-misplaced-run.docx", async () => {
        const ctx = ValidationContext.fromEntries({
            "word/document.xml": document(
                `<w:tbl><w:tblPr/><w:tblGrid><w:gridCol/><w:gridCol/></w:tblGrid>` +
                    `<w:tr><w:tc><w:r><w:t>Item</w:t></w:r></w:tc></w:tr></w:tbl>`,
            ),
        });
        expect(await runCheck.run(ctx)).toEqual([
            expect.objectContaining({
                code: runCheck.id,
                severity: "error",
                path: "word/document.xml",
                message: expect.stringMatching(/w:r sits directly inside w:tc;.*\bw:p\b/),
            }),
        ]);
    });

    it("accepts the tolerated run parents the same Word-rejected fixture also contains", async () => {
        const body =
            `<w:p><w:r><w:t>plain</w:t></w:r></w:p>` +
            `<w:p><w:ins w:id="1"><w:r><w:t>inserted</w:t></w:r></w:ins></w:p>` +
            `<w:p><w:del w:id="2"><w:r><w:delText>deleted</w:delText></w:r></w:del></w:p>` +
            `<w:p><w:hyperlink r:id="rId1"><w:r><w:t>linked</w:t></w:r></w:hyperlink></w:p>`;
        expect(await runCheck.run(ValidationContext.fromEntries({ "word/document.xml": document(body) }))).toEqual([]);
    });

    it("accepts every remaining schema-admitted run parent", async () => {
        const parents = [...RUN_PARENTS].filter((local) => !["p", "ins", "del", "hyperlink"].includes(local));
        const body = parents.map((local) => `<w:${local}><w:r><w:t>x</w:t></w:r></w:${local}>`).join("");
        expect(await runCheck.run(ValidationContext.fromEntries({ "word/document.xml": document(body) }))).toEqual([]);
    });

    it("leaves runs under foreign-namespace wrappers and part roots alone", async () => {
        const ctx = ValidationContext.fromEntries({
            "word/document.xml": document(`<w:p><x:wrap xmlns:x="urn:x"><w:r><w:t>x</w:t></w:r></x:wrap></w:p>`),
            "word/root.xml": `<w:r xmlns:w="${W}"><w:t>root</w:t></w:r>`,
            "customXml/item1.xml": `<data xmlns="urn:example"><w:r xmlns:w="${W}"/></data>`,
        });
        expect(await runCheck.run(ctx)).toEqual([]);
    });

    it("skips unparseable parts", async () => {
        const ctx = ValidationContext.fromEntries({ "word/document.xml": document("<w:tc>") });
        expect(ctx.xml("word/document.xml")?.ok).toBe(false);
        expect(await runCheck.run(ctx)).toEqual([]);
    });
});

describe("word-run-properties-misplaced", () => {
    it("reports the duplicated run properties inside one run, as in 5Errors.docx", async () => {
        const ctx = ValidationContext.fromEntries({
            "word/document.xml": document(`<w:p><w:r><w:rPr/><w:rPr><w:b/></w:rPr><w:t>A test file</w:t></w:r></w:p>`),
        });
        expect(await runPropertiesCheck.run(ctx)).toEqual([
            expect.objectContaining({
                code: runPropertiesCheck.id,
                severity: "error",
                path: "word/document.xml",
                message: expect.stringMatching(/w:rPr follows another element inside w:r;.*\bfirst\b/),
            }),
        ]);
    });

    it("reports run properties after run content and under a parent that cannot hold them", async () => {
        const ctx = ValidationContext.fromEntries({
            "word/document.xml": document(`<w:p><w:r><w:t>x</w:t><w:rPr/></w:r></w:p>`),
        });
        expect(await runPropertiesCheck.run(ctx)).toEqual([
            expect.objectContaining({ code: runPropertiesCheck.id, message: expect.stringMatching(/first, before any run content/) }),
        ]);
        const misplaced = ValidationContext.fromEntries({
            "word/document.xml": document(`<w:tbl><w:rPr/></w:tbl>`),
        });
        expect(await runPropertiesCheck.run(misplaced)).toEqual([
            expect.objectContaining({
                code: runPropertiesCheck.id,
                message: expect.stringMatching(/w:rPr sits directly inside w:tbl;.*w:pPr.*w:r\b/),
            }),
        ]);
    });

    it("accepts run properties in every schema-admitted parent", async () => {
        const ctx = ValidationContext.fromEntries({
            "word/document.xml": document(
                `<w:p><w:pPr><w:rPr/></w:pPr><w:r><w:rPr><w:b/></w:rPr><w:t>x</w:t></w:r>` +
                    `<m:oMath><m:r><m:t>1</m:t></m:r></m:oMath></w:p>` +
                    `<w:sdt><w:sdtPr><w:rPr/></w:sdtPr><w:sdtContent><w:p/></w:sdtContent></w:sdt>`,
            ),
            "word/styles.xml": styles(`<w:style w:type="character"><w:rPr/></w:style><w:tblStylePr><w:rPr/></w:tblStylePr>`),
            "word/numbering.xml": `<?xml version="1.0"?><w:numbering xmlns:w="${W}"><w:lvl><w:rPr/></w:lvl></w:numbering>`,
            "word/settings.xml": `<?xml version="1.0"?><w:settings xmlns:w="${W}"><w:shapeDefaults/></w:settings>`,
        });
        // docDefaults/rPrDefault, sdtEndPr, rPrChange, and math ctrlPr go in their own
        // parts/positions below; this part covers the common containers.
        expect(await runPropertiesCheck.run(ctx)).toEqual([]);
        const mathCtrl = ValidationContext.fromEntries({
            "word/document.xml": document(`<w:p><m:oMath><m:sPre><m:ctrlPr><w:rPr/></m:ctrlPr></m:sPre></m:oMath></w:p>`),
        });
        expect(await runPropertiesCheck.run(mathCtrl)).toEqual([]);
        const mathRun = ValidationContext.fromEntries({
            "word/document.xml": document(`<w:p><m:oMath><m:r><m:rPr/><w:rPr/><m:t>1</m:t></m:r></m:oMath></w:p>`),
        });
        expect(await runPropertiesCheck.run(mathRun)).toEqual([]);
        const defaults = ValidationContext.fromEntries({
            "word/styles.xml": styles(`<w:docDefaults><w:rPrDefault><w:rPr/></w:rPrDefault></w:docDefaults>`),
        });
        expect(await runPropertiesCheck.run(defaults)).toEqual([]);
        const endPr = ValidationContext.fromEntries({
            "word/document.xml": document(`<w:p><w:sdt><w:sdtEndPr><w:rPr/></w:sdtEndPr></w:sdt></w:p>`),
        });
        expect(await runPropertiesCheck.run(endPr)).toEqual([]);
        const change = ValidationContext.fromEntries({
            "word/document.xml": document(`<w:p><w:r><w:rPr><w:rPrChange w:id="1"><w:rPr/></w:rPrChange></w:rPr></w:r></w:p>`),
        });
        expect(await runPropertiesCheck.run(change)).toEqual([]);
    });

    it("treats a parent whose namespace is named like an Object.prototype member as foreign", async () => {
        for (const namespace of ["constructor", "__proto__", "toString"]) {
            const ctx = ValidationContext.fromEntries({
                "word/document.xml": document(`<w:p><x:wrap xmlns:x="${namespace}"><w:rPr/></x:wrap></w:p>`),
            });
            expect(ctx.xml("word/document.xml")?.ok).toBe(true);
            await expect((async () => runPropertiesCheck.run(ctx))(), namespace).resolves.toEqual([
                expect.objectContaining({
                    code: runPropertiesCheck.id,
                    message: expect.stringMatching(/w:rPr sits directly inside w:wrap;/),
                }),
            ]);
        }
    });

    it("leaves a part rooted at run properties and unparseable parts alone", async () => {
        const ctx = ValidationContext.fromEntries({
            "word/root.xml": `<w:rPr xmlns:w="${W}"/>`,
            "word/document.xml": document("<w:p"),
        });
        expect(ctx.xml("word/document.xml")?.ok).toBe(false);
        expect(await runPropertiesCheck.run(ctx)).toEqual([]);
    });
});

describe("word-underline-misplaced", () => {
    it("reports underline at paragraph-properties level, as in invalid-underline.docx", async () => {
        const ctx = ValidationContext.fromEntries({
            "word/document.xml": document(`<w:p><w:pPr><w:pStyle w:val="BodyAA"/><w:spacing/><w:u w:val="single"/></w:pPr></w:p>`),
        });
        expect(await underlineCheck.run(ctx)).toEqual([
            expect.objectContaining({
                code: underlineCheck.id,
                severity: "error",
                path: "word/document.xml",
                message: expect.stringMatching(/w:u sits directly inside w:pPr;.*w:rPr/),
            }),
        ]);
    });

    it("accepts underline inside run properties and skips foreign wrappers, roots, and broken parts", async () => {
        const ctx = ValidationContext.fromEntries({
            "word/document.xml": document(`<w:p><w:r><w:rPr><w:u w:val="single"/></w:rPr><w:t>x</w:t></w:r></w:p>`),
            "word/other.xml": `<x:u xmlns:x="urn:x"><w:u xmlns:w="${W}"/></x:u>`,
            "word/root.xml": `<w:u xmlns:w="${W}"/>`,
            "word/broken.xml": "<w:u",
        });
        expect(await underlineCheck.run(ctx)).toEqual([]);
    });
});

describe("run placement tables", () => {
    it("expose the parents the checks tolerate", () => {
        expect([...UNDERLINE_PARENTS]).toEqual(["rPr"]);
        expect(RUN_PARENTS.has("tc")).toBe(false);
        expect(RUN_PARENTS.has("tbl")).toBe(false);
        expect(RUN_PARENTS.has("body")).toBe(false);
        expect([...RUN_PROPERTIES_PARENTS[M]!]).toEqual(["ctrlPr", "r"]);
        expect(RUN_PROPERTIES_PARENTS[W]?.has("r")).toBe(true);
        expect(RUN_PROPERTIES_PARENTS[W]?.has("pPr")).toBe(true);
        expect(RUN_PROPERTIES_PARENTS[W]?.has("style")).toBe(true);
    });
});
