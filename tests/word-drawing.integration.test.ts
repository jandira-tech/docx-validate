/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

import { readFile } from "node:fs/promises";
import path from "node:path";

import { describe, expect, it, vi } from "vitest";

import { ValidationContext } from "../src/lib/validation-context";
import { getElementsByTagNameNSAll, parseXml } from "../src/lib/xml-helpers";
import * as xsdMod from "../src/lib/xsd-validator";
import { wordDrawingExtentMissingCheck as check } from "../src/scripts/office/checks/word-drawing";
import { validate } from "../src/scripts/office/validate";

const fixtures = path.join(import.meta.dirname, "fixtures");
const rejected = path.join(fixtures, "broken/header-with-rid-no-sidecar.docx");
const tolerated = path.join(fixtures, "working/package.word-drawing-extent-positive.docx");
const XSD = "http://www.w3.org/2001/XMLSchema";

describe("drawing extent Word evidence", { timeout: 120_000 }, () => {
    it("rejects the Word open-error header fixture and accepts the positive-extent drawing control", async () => {
        const manifest = JSON.parse(await readFile(path.join(fixtures, "../fixtures-all.manifest.json"), "utf8"));
        expect(
            manifest.entries.find((entry: { relativePath: string }) => entry.relativePath === "broken/header-with-rid-no-sidecar.docx")
                ?.word,
        ).toBe("open-error");
        const result = await validate(rejected, { profile: "word-valid" });
        expect(result.valid).toBe(false);
        expect(result.issues).toContainEqual(expect.objectContaining({ code: check.id, severity: "error", path: "word/header1.xml" }));
        const bytes = await readFile(tolerated);
        expect(await check.run(await ValidationContext.fromArchive(bytes))).toEqual([]);
        expect((await validate(tolerated, { profile: "word-valid" })).valid).toBe(true);
        expect(await readFile(tolerated)).toEqual(bytes);
    });

    it("rejects without any XSD diagnostic", async () => {
        const factory = vi.spyOn(xsdMod, "createXsdValidator").mockResolvedValue({ validate: async () => [] });
        try {
            const result = await validate(rejected, { profile: "word-valid" });
            expect(result.valid).toBe(false);
            expect(result.issues).toContainEqual(expect.objectContaining({ code: check.id, severity: "error" }));
            expect(result.xsdWalk?.started).toBe(false);
        } finally {
            factory.mockRestore();
        }
    });

    it("pins the required extent parents and its position before graphic to the bundled XSD", async () => {
        const schema = parseXml(
            await readFile(
                path.join(import.meta.dirname, "../src/scripts/office/schemas/ISO-IEC29500-4_2016/dml-wordprocessingDrawing.xsd"),
                "utf8",
            ),
        );
        const types = getElementsByTagNameNSAll(schema, XSD, "complexType");
        const extentParents = types.filter((type) =>
            getElementsByTagNameNSAll(type, XSD, "element").some((el) => el.getAttribute("name") === "extent"),
        );
        const globals = getElementsByTagNameNSAll(schema, XSD, "element").filter((el) => el.parentNode === schema.documentElement);
        expect(
            extentParents.map((type) => globals.find((el) => el.getAttribute("type") === type.getAttribute("name"))?.getAttribute("name")),
        ).toEqual(["inline", "anchor"]);
        for (const parent of extentParents) {
            const children = getElementsByTagNameNSAll(parent, XSD, "element");
            const extentIndex = children.findIndex((el) => el.getAttribute("name") === "extent");
            const graphicIndex = children.findIndex((el) => el.getAttribute("ref") === "a:graphic");
            expect(extentIndex).toBeGreaterThanOrEqual(0);
            expect(extentIndex).toBeLessThan(graphicIndex);
            expect(children[extentIndex]?.getAttribute("minOccurs") || "1").toBe("1");
        }
    });
});
