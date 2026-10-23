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
import { wordPictureNvpicprMissingCheck, wordPictureSpprMissingCheck } from "../src/scripts/office/checks/word-picture";
import { validate } from "../src/scripts/office/validate";

const fixtures = path.join(import.meta.dirname, "fixtures");
const rejected = path.join(fixtures, "broken/header-with-rid-no-sidecar.docx");
const tolerated = path.join(fixtures, "working/package.word-drawing-extent-positive.docx");
const XSD = "http://www.w3.org/2001/XMLSchema";

describe.each([wordPictureNvpicprMissingCheck, wordPictureSpprMissingCheck])("$id Word evidence", { timeout: 120_000 }, (check) => {
    it("rejects the recorded Word open-error header and accepts the complete picture control", async () => {
        const manifest = JSON.parse(await readFile(path.join(fixtures, "../fixtures-all.manifest.json"), "utf8"));
        expect(
            manifest.entries.find((entry: { relativePath: string }) => entry.relativePath === "broken/header-with-rid-no-sidecar.docx")
                ?.word,
        ).toBe("open-error");
        const bytes = await readFile(rejected);
        const result = await validate(rejected, { profile: "word-valid" });
        expect(result.valid).toBe(false);
        expect(result.issues).toContainEqual(expect.objectContaining({ code: check.id, severity: "error", path: "word/header1.xml" }));
        expect(await readFile(rejected)).toEqual(bytes);
        expect(await check.run(await ValidationContext.fromArchive(await readFile(tolerated)))).toEqual([]);
        expect((await validate(tolerated, { profile: "word-valid" })).valid).toBe(true);
    });

    it("rejects independently of XSD diagnostics", async () => {
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

    it("pins the required picture children and their order to CT_Picture without compiling XSD", async () => {
        const schema = parseXml(
            await readFile(path.join(import.meta.dirname, "../src/scripts/office/schemas/ISO-IEC29500-4_2016/dml-picture.xsd"), "utf8"),
        );
        const type = getElementsByTagNameNSAll(schema, XSD, "complexType").find((el) => el.getAttribute("name") === "CT_Picture")!;
        const children = getElementsByTagNameNSAll(type, XSD, "element");
        expect(children.map((el) => el.getAttribute("name"))).toEqual(["nvPicPr", "blipFill", "spPr"]);
        expect(children.map((el) => el.getAttribute("minOccurs"))).toEqual(["1", "1", "1"]);
        expect(children.map((el) => el.getAttribute("maxOccurs"))).toEqual(["1", "1", "1"]);
        const parents = getElementsByTagNameNSAll(schema, XSD, "element").filter((el) => el.getAttribute("type") === "CT_Picture");
        expect(parents.map((el) => el.getAttribute("name"))).toEqual(["pic"]);
    });
});
