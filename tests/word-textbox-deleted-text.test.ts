/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

import path from "node:path";
import { fileURLToPath } from "node:url";

import { describe, expect, it } from "vitest";

import { validate } from "../src/scripts/office/validate";

const HERE = path.dirname(fileURLToPath(import.meta.url));
const DEL_TEXT = path.join(HERE, "fixtures", "broken", "package.word-textbox-deleted-text.docx");
const DEL_INSTR = path.join(HERE, "fixtures", "broken", "package.word-textbox-deleted-instruction.docx");
const DEL_SHAPE = path.join(HERE, "fixtures", "broken", "package.word-textbox-deleted-text-shape.docx");
const CLEAN = path.join(HERE, "fixtures", "working", "package.word-textbox-deleted-text-removed.docx");
const SHAPE_CLEAN = path.join(HERE, "fixtures", "working", "package.word-textbox-deleted-text-shape-repaired.docx");
const CHOICE_DEL = path.join(HERE, "fixtures", "broken", "package.word-textbox-choice-deleted-text.docx");
const FALLBACK_ONLY = path.join(HERE, "fixtures", "working", "package.word-textbox-fallback-deleted-text.docx");
const INNER_DEL = path.join(HERE, "fixtures", "working", "package.word-textbox-inner-deletion.docx");

describe("word-valid deleted text in a text box", { timeout: 30000 }, () => {
    it("rejects w:delText inside a VML text box", async () => {
        const result = await validate(DEL_TEXT, { profile: "word-valid" });
        expect(result.valid).toBe(false);
        expect(result.issues.some((issue) => issue.severity === "error" && issue.code === "word-textbox-deleted-text")).toBe(true);
    });

    it("rejects w:delText inside a drawingML text box", async () => {
        const result = await validate(DEL_SHAPE, { profile: "word-valid" });
        expect(result.valid).toBe(false);
        expect(result.issues.some((issue) => issue.severity === "error" && issue.code === "word-textbox-deleted-text")).toBe(true);
    });

    it("does not reject that drawing once the deleted text is ordinary text", async () => {
        const result = await validate(SHAPE_CLEAN, { profile: "word-valid" });
        expect(result.issues.some((issue) => issue.code === "word-textbox-deleted-text")).toBe(false);
        expect(result.valid).toBe(true);
    });

    it("rejects w:delInstrText inside a VML text box", async () => {
        const result = await validate(DEL_INSTR, { profile: "word-valid" });
        expect(result.valid).toBe(false);
        expect(result.issues.some((issue) => issue.severity === "error" && issue.code === "word-textbox-deleted-text")).toBe(true);
    });

    it("does not reject the text box once that deleted text is removed", async () => {
        const result = await validate(CLEAN, { profile: "word-valid" });
        expect(result.issues.some((issue) => issue.code === "word-textbox-deleted-text")).toBe(false);
        expect(result.valid).toBe(true);
    });

    it.each(["lenient", "strict"] as const)("does not report the cause on %s", async (profile) => {
        const result = await validate(DEL_TEXT, { profile });
        expect(result.issues.some((issue) => issue.code === "word-textbox-deleted-text")).toBe(false);
    });

    it("rejects w:delText inside an mc:Choice text box", async () => {
        const result = await validate(CHOICE_DEL, { profile: "word-valid" });
        expect(result.valid).toBe(false);
        expect(result.issues.some((issue) => issue.severity === "error" && issue.code === "word-textbox-deleted-text")).toBe(true);
    });

    it("does not reject w:delText wrapped in a w:del inside the text box", async () => {
        const result = await validate(INNER_DEL, { profile: "word-valid" });
        expect(result.issues.some((issue) => issue.code === "word-textbox-deleted-text")).toBe(false);
        expect(result.valid).toBe(true);
    });

    it("does not reject w:delText that remains only inside mc:Fallback", async () => {
        const result = await validate(FALLBACK_ONLY, { profile: "word-valid" });
        expect(result.issues.some((issue) => issue.code === "word-textbox-deleted-text")).toBe(false);
        expect(result.valid).toBe(true);
    });

    it.each(["lenient", "strict"] as const)("does not report the choice-box cause on %s", async (profile) => {
        const result = await validate(CHOICE_DEL, { profile });
        expect(result.issues.some((issue) => issue.code === "word-textbox-deleted-text")).toBe(false);
    });
});
