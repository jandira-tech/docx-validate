/*
 * Copyright 2026 Jandira Technologies, LLC
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 */

import path from "node:path";
import { describe, expect, it } from "vitest";
import { WORD_ERROR_EXPLANATIONS, explainWordError } from "../src/scripts/office/validators/word-error-explanations";

describe("word error explanations", () => {
    it("explains relationships markup compatibility as a strict defect Word accepts (#632)", () => {
        const explanation = explainWordError("opc-relationships-markup-compatibility");
        expect(explanation).toContain("Word");
        expect(explanation).toContain("markup compatibility");
        expect(explanation).toContain("strict");
        expect(explanation).toContain("Relationship");
    });

    it("explains the package/story-root duplicate-Id scope and repair (#615, #617)", () => {
        const explanation = explainWordError("rels-id-duplicate");
        expect(explanation).toEqual(expect.stringContaining("_rels/.rels"));
        expect(explanation).toContain("main document");
        expect(explanation).toContain("glossary");
        expect(explanation).toContain("unreferenced");
        expect(explanation).toContain("unique");
    });

    it("explains a story-reference repair that needs no existing id (#623)", () => {
        const explanation = explainWordError("word-header-footer-relationship");
        expect(explanation).toContain(
            "Give the reference a non-empty r:id naming a Relationship of the matching Type: add that Relationship when none exists, or retype the one it names.",
        );
        expect(explanation).not.toContain("add one with that Id and Type");
    });

    it("explains both absent and multiple main relationships (#515)", () => {
        const explanation = explainWordError("word-main-document-part");
        expect(explanation).toContain("exactly one internal officeDocument relationship");
        expect(explanation).toContain("Remove extra");
    });

    it("explains the unconditional Word-blocking codes", () => {
        // The codes isWordBlockingIssue always treats as Word-blocking must each
        // carry a user-facing explanation so the CLI can tell the user why Word fails.
        const alwaysBlocking = [
            "ignorable-undeclared",
            "word-math-spre-body",
            "word-math-parse",
            "word-content-type-invalid",
            "word-drawing-scalar-whitespace",
            "id-durable-overflow",
            "comment-thread-commentid-paraid-orphan",
            "comment-thread-durableid-orphan",
        ];
        for (const code of alwaysBlocking) {
            expect(WORD_ERROR_EXPLANATIONS[code], `missing explanation for ${code}`).toBeTruthy();
        }
    });

    it("explainWordError returns the mapped string or undefined", () => {
        expect(explainWordError("ignorable-undeclared")).toContain("mc:Ignorable");
        expect(explainWordError("not-a-real-code")).toBeUndefined();
        expect(explainWordError(undefined)).toBeUndefined();
    });
});

describe("word error explanations — CLI wiring", () => {
    it("prints the explanation line under --profile word-valid for a mapped code", async () => {
        const { execFile } = await import("node:child_process");
        const { promisify } = await import("node:util");
        const execFileAsync = promisify(execFile);
        const fixture = path.join(
            import.meta.dirname,
            "fixtures",
            "word-regenerate-invalid",
            "original",
            "external",
            "superdoc",
            "super-editor",
            "annotations_import.docx",
        );
        const proc = await execFileAsync("node", [
            "--import",
            "tsx",
            path.join(import.meta.dirname, "..", "src", "scripts", "office", "validate.ts"),
            fixture,
            "--profile",
            "word-valid",
        ]).catch((e: NodeJS.ErrnoException & { stderr?: string }) => e);
        const stderr = typeof proc === "string" ? "" : ((proc as { stderr?: string }).stderr ?? "");
        expect(stderr).toMatch(/↳ A <Relationship> entry is missing a required attribute/);
    });
});
