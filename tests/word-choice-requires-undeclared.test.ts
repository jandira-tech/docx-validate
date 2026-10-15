/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

import path from "node:path";
import { promises as fs } from "node:fs";
import { fileURLToPath } from "node:url";

import { describe, expect, it } from "vitest";

import { withTempDir } from "../src/lib/run-cli";
import { DOCXSchemaValidator } from "../src/scripts/office/validators/docx";
import { validate } from "../src/scripts/office/validate";

const HERE = path.dirname(fileURLToPath(import.meta.url));
const BROKEN = path.join(HERE, "fixtures", "broken", "package.word-choice-requires-undeclared.docx");
const DESCENDANT = path.join(HERE, "fixtures", "broken", "package.word-choice-requires-on-descendant.docx");
const DECLARED = path.join(HERE, "fixtures", "working", "package.word-choice-requires-declared.docx");
const ANCESTOR = path.join(HERE, "fixtures", "working", "package.word-choice-requires-on-ancestor.docx");

describe("word-valid choice requires prefix", { timeout: 30000 }, () => {
    it("rejects mc:Choice when Requires names a prefix that is not in scope", async () => {
        const result = await validate(BROKEN, { profile: "word-valid" });
        expect(result.valid).toBe(false);
        expect(result.issues.some((issue) => issue.severity === "error" && issue.code === "word-choice-requires-undeclared")).toBe(true);
    });

    it("rejects the same choice when the prefix is declared only on a descendant", async () => {
        const result = await validate(DESCENDANT, { profile: "word-valid" });
        expect(result.valid).toBe(false);
        expect(result.issues.some((issue) => issue.severity === "error" && issue.code === "word-choice-requires-undeclared")).toBe(true);
    });

    it("does not reject the choice once the prefix is declared on it", async () => {
        const result = await validate(DECLARED, { profile: "word-valid" });
        expect(result.issues.some((issue) => issue.code === "word-choice-requires-undeclared")).toBe(false);
        expect(result.valid).toBe(true);
    });

    it("does not reject the choice when an ancestor declares the prefix", async () => {
        const result = await validate(ANCESTOR, { profile: "word-valid" });
        expect(result.issues.some((issue) => issue.code === "word-choice-requires-undeclared")).toBe(false);
        expect(result.valid).toBe(true);
    });

    it("accepts the predefined xml prefix without an explicit declaration", async () => {
        await withTempDir(async (dir) => {
            await fs.mkdir(path.join(dir, "word"));
            await fs.writeFile(
                path.join(dir, "word/document.xml"),
                '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ' +
                    'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">' +
                    '<w:body><w:p><mc:AlternateContent><mc:Choice Requires="xml"><w:r><w:t>hello</w:t></w:r>' +
                    "</mc:Choice><mc:Fallback/></mc:AlternateContent></w:p></w:body></w:document>",
            );
            const result = await new DOCXSchemaValidator({ unpackedDir: dir, profile: "word-valid" }).validate();
            expect(result.issues.some((issue) => issue.code === "word-choice-requires-undeclared")).toBe(false);
        });
    });

    it.each(["lenient", "strict"] as const)("does not report the cause on %s", async (profile) => {
        const result = await validate(BROKEN, { profile });
        expect(result.issues.some((issue) => issue.code === "word-choice-requires-undeclared")).toBe(false);
    });
});
