/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

import { readFile, writeFile } from "node:fs/promises";
import path from "node:path";

import JSZip from "jszip";
import { describe, expect, it, vi } from "vitest";

import { withTempDir } from "../src/lib/run-cli";
import { getElementsByTagNameNSAll, parseXml, serializeXml } from "../src/lib/xml-helpers";
import * as xsdMod from "../src/lib/xsd-validator";
import {
    DRAWING_EXTENT_PREDECESSORS,
    wordDrawingExtentMisplacedCheck as check,
} from "../src/scripts/office/checks/word-drawing-extent-position";
import { validate } from "../src/scripts/office/validate";

const WP = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
const PIC = "http://schemas.openxmlformats.org/drawingml/2006/picture";
const XSD = "http://www.w3.org/2001/XMLSchema";
const rejected = path.join(import.meta.dirname, "fixtures/broken/header-with-rid-no-sidecar.docx");

// The source header has no extent. Complete its scaffold, then move just the
// extent. Rejection provenance is the historical extent heuristic comment,
// not a recorded Word probe of this synthetic package. Never launch Word.
const headerWithExtent = async (misplaced: boolean): Promise<Uint8Array> => {
    const zip = await JSZip.loadAsync(await readFile(rejected));
    const dom = parseXml(await zip.file("word/header1.xml")!.async("string"));
    const inline = getElementsByTagNameNSAll(dom, WP, "inline")[0]!;
    const extent = dom.createElementNS(WP, "wp:extent");
    extent.setAttribute("cx", "95250");
    extent.setAttribute("cy", "95250");
    inline.insertBefore(extent, inline.firstChild);
    const docPr = dom.createElementNS(WP, "wp:docPr");
    docPr.setAttribute("id", "1");
    docPr.setAttribute("name", "Picture 1");
    inline.insertBefore(docPr, extent.nextSibling);
    const pic = getElementsByTagNameNSAll(dom, PIC, "pic")[0]!;
    const nvPicPr = dom.createElementNS(PIC, "pic:nvPicPr");
    const cNvPr = dom.createElementNS(PIC, "pic:cNvPr");
    cNvPr.setAttribute("id", "0");
    cNvPr.setAttribute("name", "image1.png");
    nvPicPr.appendChild(cNvPr);
    nvPicPr.appendChild(dom.createElementNS(PIC, "pic:cNvPicPr"));
    pic.insertBefore(nvPicPr, pic.firstChild);
    pic.appendChild(dom.createElementNS(PIC, "pic:spPr"));
    if (misplaced) {
        inline.appendChild(extent);
    }
    zip.file("word/header1.xml", serializeXml(dom));
    return zip.generateAsync({ type: "uint8array" });
};

describe("drawing extent position", { timeout: 120_000 }, () => {
    it("rejects the synthetic header with extent after graphic and accepts the ordered control", async () => {
        await withTempDir(async (dir) => {
            const file = path.join(dir, "extent-after-graphic.docx");
            const bytes = await headerWithExtent(true);
            await writeFile(file, bytes);
            const result = await validate(file, { profile: "word-valid" });
            expect(result.valid).toBe(false);
            expect(result.issues).toContainEqual(expect.objectContaining({ code: check.id, severity: "error", path: "word/header1.xml" }));
            expect(await readFile(file)).toEqual(Buffer.from(bytes));
            const control = path.join(dir, "ordered-header.docx");
            await writeFile(control, await headerWithExtent(false));
            expect((await validate(control, { profile: "word-valid" })).valid).toBe(true);
        });
    });

    it("rejects for the structural code alone when XSD returns no issues", async () => {
        const factory = vi.spyOn(xsdMod, "createXsdValidator").mockResolvedValue({ validate: async () => [] });
        try {
            await withTempDir(async (dir) => {
                const file = path.join(dir, "extent-after-graphic.docx");
                await writeFile(file, await headerWithExtent(true));
                const result = await validate(file, { profile: "word-valid" });
                expect(result.valid).toBe(false);
                expect(result.issues.filter((issue) => issue.severity === "error").map((issue) => issue.code)).toEqual([check.id]);
                expect(result.xsdWalk.started).toBe(false);
            });
        } finally {
            factory.mockRestore();
        }
    });

    it("pins the legal parent set and extent sequence prefixes to the bundled drawing XSD", async () => {
        const schema = parseXml(
            await readFile(
                path.join(import.meta.dirname, "../src/scripts/office/schemas/ISO-IEC29500-4_2016/dml-wordprocessingDrawing.xsd"),
                "utf8",
            ),
        );
        const types = getElementsByTagNameNSAll(schema, XSD, "complexType");
        const globals = getElementsByTagNameNSAll(schema, XSD, "element").filter((el) => el.parentNode === schema.documentElement);
        const prefixes = types.flatMap((type) => {
            const children = getElementsByTagNameNSAll(type, XSD, "element");
            const extentIndex = children.findIndex((el) => el.getAttribute("name") === "extent");
            if (extentIndex < 0) {
                return [];
            }
            expect(children[extentIndex]!.getAttribute("minOccurs") || "1").toBe("1");
            expect(children[extentIndex]!.getAttribute("maxOccurs") || "1").toBe("1");
            for (const suffix of ["effectExtent", "docPr", "cNvGraphicFramePr", "a:graphic"]) {
                expect(children.findIndex((el) => (el.getAttribute("name") || el.getAttribute("ref")) === suffix)).toBeGreaterThan(
                    extentIndex,
                );
            }
            const parent = globals.find((el) => el.getAttribute("type") === type.getAttribute("name"))!.getAttribute("name");
            return [[parent, children.slice(0, extentIndex).map((el) => el.getAttribute("name"))]];
        });
        expect(prefixes).toEqual([...DRAWING_EXTENT_PREDECESSORS]);
    });
});
