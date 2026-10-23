/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

import { createHash } from "node:crypto";
import { readFile, writeFile } from "node:fs/promises";
import path from "node:path";

import { describe, expect, it } from "vitest";

import { ValidationContext } from "../src/lib/validation-context";
import { withTempDir } from "../src/lib/run-cli";
import { getElementsByTagNameNSAll, parseXml, serializeXml } from "../src/lib/xml-helpers";
import { validate } from "../src/scripts/office/validate";
import { buildDocx, MINIMAL_PARTS } from "./support/minimal-docx";

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const lane = path.join(import.meta.dirname, "lane/jubarte-0.13.0");

describe("table border Word evidence boundary", { timeout: 180_000 }, () => {
    it.each([
        "k07_1.redline.docx",
        "k07_2.redline.docx",
        "k08_1.redline.docx",
        "k08_2.redline.docx",
        "k09_1.redline.docx",
        "k09_2.redline.docx",
    ])("keeps %s historically Word-opened style override advisory", async (file) => {
        const bytes = await readFile(path.join(lane, file));
        const hash = createHash("sha256").update(bytes).digest("hex");
        expect(await readFile(path.join(lane, "manifest.csv"), "utf8")).toContain(`fixtures/${file},redline,True,${hash},`);
        const ctx = await ValidationContext.fromArchive(bytes);
        const styles = ctx.xml("word/styles.xml");
        expect(styles?.ok).toBe(true);
        if (!styles?.ok) {
            throw new Error("Expected readable lane styles");
        }
        const misplaced = getElementsByTagNameNSAll(styles.document, W, "tblBorders").filter(
            (el) => (el.parentNode as Element).localName === "tcPr",
        );
        expect(misplaced).toHaveLength(1);
        const parent = misplaced[0]!.parentNode as Element;
        expect(parent.namespaceURI).toBe(W);
        expect(
            Array.from(parent.childNodes)
                .filter((node) => node.nodeType === 1)
                .map((node) => (node as Element).localName),
        ).toEqual(["tblBorders", "tcPr"]);
        const override = parent.parentNode as Element;
        expect(override.localName).toBe("tblStylePr");
        expect(override.getAttributeNS(W, "type")).toBe("firstRow");
        expect((override.parentNode as Element).localName).toBe("style");
        const result = await validate(path.join(lane, file), { profile: "word-valid" });
        expect(result.issues).toContainEqual(
            expect.objectContaining({
                code: "xsd-error",
                path: "word/styles.xml",
                severity: "warning",
                message: expect.stringContaining("}tblBorders': This element is not expected"),
            }),
        );
        expect(result.valid).toBe(true);
        expect(await readFile(path.join(lane, file))).toEqual(bytes);
    });

    it("does not hide a different tblBorders ordering error beside the tolerated style override", async () => {
        const style =
            '<w:style w:type="table" w:styleId="Mixed"><w:name w:val="Mixed"/>' +
            '<w:tblPr><w:tblLayout w:type="fixed"/><w:tblBorders/></w:tblPr>' +
            '<w:tblStylePr w:type="firstRow"><w:tcPr><w:tblBorders/></w:tcPr></w:tblStylePr></w:style>';
        await withTempDir(async (dir) => {
            const file = path.join(dir, "mixed-style-borders.docx");
            await writeFile(
                file,
                await buildDocx({
                    ...MINIMAL_PARTS,
                    "word/styles.xml": MINIMAL_PARTS["word/styles.xml"]!.replace("</w:styles>", `${style}</w:styles>`),
                }),
            );
            const result = await validate(file, { profile: "word-valid" });
            expect(result.valid).toBe(false);
            expect(result.issues).toContainEqual(
                expect.objectContaining({
                    code: "xsd-error",
                    severity: "error",
                    path: "word/styles.xml",
                    message: expect.stringContaining("}tblBorders'"),
                }),
            );
        });
    });

    it("pins the rejected body-table ordering and the historical UNREADABLE_CONTENT specimen", async () => {
        const file = "tests__fixtures__external__docx-parser-converter__tagged_tests__table_tests_v2.docx";
        const corpus = path.join(import.meta.dirname, "../word-invalid-fixtures");
        const probe = (await readFile(path.join(corpus, "probe-reconfirm-results.jsonl"), "utf8"))
            .trim()
            .split("\n")
            .map((line) => JSON.parse(line));
        expect(probe).toContainEqual(expect.objectContaining({ path: expect.stringContaining(file), verdict: "UNREADABLE_CONTENT" }));
        const ctx = await ValidationContext.fromArchive(await readFile(path.join(corpus, "UNREADABLE_CONTENT", file)));
        const parsed = ctx.xml("word/document.xml");
        if (!parsed?.ok) {
            throw new Error("Expected readable table specimen");
        }
        const border = getElementsByTagNameNSAll(parsed.document, W, "tblBorders")[0]!;
        const tblPr = border.parentNode as Element;
        expect(tblPr.localName).toBe("tblPr");
        const names = Array.from(tblPr.childNodes)
            .filter((node) => node.nodeType === 1)
            .map((node) => (node as Element).localName);
        expect(names.slice(0, 3)).toEqual(["tblW", "tblLayout", "tblBorders"]);
        // Isolate that ordering from the source's unrelated blocking defects.
        const table =
            `<w:tbl>${serializeXml(tblPr)}<w:tblGrid><w:gridCol w:w="100"/></w:tblGrid>` +
            '<w:tr w14:paraId="00000003" w14:textId="00000003"><w:tc><w:p w14:paraId="00000004" w14:textId="00000004"/></w:tc></w:tr></w:tbl>';
        await withTempDir(async (dir) => {
            const target = path.join(dir, "body-border-order.docx");
            await writeFile(
                target,
                await buildDocx({
                    ...MINIMAL_PARTS,
                    "word/document.xml": MINIMAL_PARTS["word/document.xml"]!.replace("<w:sectPr>", `${table}<w:sectPr>`),
                }),
            );
            const result = await validate(target, { profile: "word-valid" });
            expect(result.valid).toBe(false);
            expect(result.issues).toContainEqual(
                expect.objectContaining({
                    code: "xsd-error",
                    severity: "error",
                    path: "word/document.xml",
                    message: expect.stringContaining("}tblBorders': This element is not expected"),
                }),
            );
        });
        const schema = parseXml(
            await readFile(path.join(import.meta.dirname, "../src/scripts/office/schemas/ISO-IEC29500-4_2016/wml.xsd"), "utf8"),
        );
        const type = getElementsByTagNameNSAll(schema, "http://www.w3.org/2001/XMLSchema", "complexType").find(
            (el) => el.getAttribute("name") === "CT_TblPrBase",
        )!;
        const order = getElementsByTagNameNSAll(type, "http://www.w3.org/2001/XMLSchema", "element").map((el) => el.getAttribute("name"));
        expect(order.indexOf("tblBorders")).toBeGreaterThanOrEqual(0);
        expect(order.indexOf("tblBorders")).toBeLessThan(order.indexOf("tblLayout"));
    });
});
