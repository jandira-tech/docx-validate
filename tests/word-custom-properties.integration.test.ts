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
import { wordCustomPropertyFiletimeEmptyCheck as check } from "../src/scripts/office/checks/word-custom-properties";
import { validate } from "../src/scripts/office/validate";
import { observedWordOutcome } from "./support/word-observation";

const fixtures = path.join(import.meta.dirname, "fixtures");
const rejected = path.join(fixtures, "external/superdoc/super-editor/paragraph-sectpr-breaks.docx");
const tolerated = path.join(fixtures, "external/superdoc/super-editor/complex2.docx");
const schemas = path.join(import.meta.dirname, "../src/scripts/office/schemas/ISO-IEC29500-4_2016");
const XSD = "http://www.w3.org/2001/XMLSchema";

describe("custom-property Word evidence", { timeout: 120_000 }, () => {
    it("pairs the manifest's recovery fixture with its clean-open empty-string control", async () => {
        for (const [file, outcome, count] of [
            [rejected, "unreadable-content-warning", 2],
            [tolerated, "clean-open", 0],
        ] as const) {
            const relativePath = path.relative(fixtures, file);
            expect(await observedWordOutcome(relativePath)).toBe(outcome);
            const bytes = await readFile(file);
            const ctx = await ValidationContext.fromArchive(bytes);
            expect(await check.run(ctx)).toHaveLength(count);
            expect(await readFile(file)).toEqual(bytes);
        }
        const result = await validate(rejected, { profile: "word-valid" });
        expect(result.valid).toBe(false);
        expect(result.issues).toContainEqual(expect.objectContaining({ code: check.id, severity: "error", path: "docProps/custom.xml" }));
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

    it("derives the dateTime scalar and direct custom-property parent from the bundled schemas without compiling them", async () => {
        const variants = parseXml(await readFile(path.join(schemas, "shared-documentPropertiesVariantTypes.xsd"), "utf8"));
        const filetime = getElementsByTagNameNSAll(variants, XSD, "element").find((el) => el.getAttribute("name") === "filetime");
        expect(filetime?.getAttribute("type")).toBe("xsd:dateTime");
        const custom = parseXml(await readFile(path.join(schemas, "shared-documentPropertiesCustom.xsd"), "utf8"));
        const property = getElementsByTagNameNSAll(custom, XSD, "complexType").find((el) => el.getAttribute("name") === "CT_Property");
        expect(property).toBeDefined();
        expect(getElementsByTagNameNSAll(property!, XSD, "element").some((el) => el.getAttribute("ref") === "vt:filetime")).toBe(true);
        expect(
            getElementsByTagNameNSAll(custom, XSD, "element")
                .find((el) => el.getAttribute("name") === "property")
                ?.getAttribute("type"),
        ).toBe("CT_Property");
    });
});
