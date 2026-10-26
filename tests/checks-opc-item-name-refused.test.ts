/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */
import { describe, expect, it } from "vitest";
import { entryNameRefusal, ValidationContext } from "../src/lib/validation-context";
import { createDefaultCheckRegistry } from "../src/scripts/office/checks";
import { isWordBlockingIssue } from "../src/scripts/office/validators/word-blocking";

const ID = "opc-item-name-refused";
const REFUSED = [
    "word/a/../b.xml",
    "../evil.xml",
    "/word/abs.xml",
    "\\word\\abs.xml",
    "c:doc/x.xml",
    "C:/x.xml",
    "word//x.xml",
    "word/./x.xml",
    "word/.../x.xml",
];
const CLEAN = ["word/document.xml", "[Content_Types].xml", "word/..x.xml", "word./x.xml", "word/c:x.xml", "word/a.b.xml"];

describe("refused item names: deterministic check (#629)", () => {
    it("registers a stored-names-only package check for all formats", () => {
        expect(createDefaultCheckRegistry().get(ID)).toMatchObject({
            groups: ["package"],
            formats: ["docx", "pptx", "xlsx"],
            storedNamesOnly: true,
        });
    });

    it.each(REFUSED)("reports %s once, as the refusal the package view gives it, without reading parts", async (name) => {
        const check = createDefaultCheckRegistry().get(ID);
        expect(check).toBeDefined();
        const ctx = ValidationContext.fromEntries([
            [name, "unused"],
            [name, "duplicate"],
        ]);
        const issues = await check!.run(ctx);
        expect(issues).toEqual([{ code: ID, severity: "error", path: name, message: entryNameRefusal(name) }]);
        expect(isWordBlockingIssue(issues[0]!, new Set())).toBe(true);
        expect(ctx.readCount(name)).toBe(0);
    });

    it("reports every refused name, not only the first", async () => {
        const check = createDefaultCheckRegistry().get(ID)!;
        const issues = await check.run(ValidationContext.fromEntries([...REFUSED, ...CLEAN].map((name) => [name, "unused"] as const)));
        expect(issues.map((issue) => issue.path)).toEqual([...REFUSED].sort());
    });

    it("does not flag names that place a part", async () => {
        const check = createDefaultCheckRegistry().get(ID)!;
        expect(await check.run(ValidationContext.fromEntries(CLEAN.map((name) => [name, "unused"] as const)))).toEqual([]);
    });
});
