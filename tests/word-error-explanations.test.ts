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
    it.each(["mce-malformed", "mce-must-understand"])("explains %s by reached story roots (#644)", (code) => {
        const explanation = explainWordError(code);
        expect(explanation).toContain("_rels/.rels");
        expect(explanation).toContain("main document");
        expect(explanation).toContain("glossary");
        expect(explanation).toContain("stray");
        expect(explanation).toContain("unobserved");
    });
    it.each(["mce-malformed", "mce-must-understand"])("explains %s as a warning in a word/ part no relationship names (#1116)", (code) => {
        expect(explainWordError(code)).toContain(
            "In an XML part under word/ that no relationship names, Word opened an undeclared Choice Requires (session 83) and a root MustUnderstand (session 84) clean,",
        );
    });
    it.each(["mce-malformed", "mce-must-understand"])(
        "explains %s as a warning in a story part a root relates but does not use (#1125)",
        (code) => {
            expect(explainWordError(code)).toContain(
                "and both in a header the main document relates but does not use (session 84), so word-valid warns there (#1116, #1125)",
            );
        },
    );
    it("explains a package root that unwraps to the Relationships element as accepted (#646)", () => {
        const explanation = explainWordError("mce-malformed");
        expect(explanation).not.toContain("still rejects");
        expect(explanation).toContain("opened clean roots that unwrap to the Relationships element, which word-valid accepts (#646)");
        expect(explanation).toContain("unwraps to none");
    });
    it("explains character data directly inside AlternateContent (#851, sessions 65 and 67)", () => {
        const explanation = explainWordError("mce-malformed");
        expect(explanation).toContain("must not hold character data directly inside it");
        expect(explanation).toContain(
            "before the Choice, between the branches, after the Fallback, after a Choice with no Fallback and between two Choices",
        );
        expect(explanation).toContain(
            "opened the same roots clean with the text inside a Choice or the Fallback or with whitespace only between the branches",
        );
        expect(explanation).toContain("(sessions 65 and 67)");
        expect(explanation).toContain("other parts and relationships parts are unprobed");
    });
    it("explains relationships markup compatibility with the observed Word behaviour (#632)", () => {
        const explanation = explainWordError("opc-relationships-markup-compatibility");
        expect(explanation).toContain("Word");
        expect(explanation).toContain("markup compatibility");
        expect(explanation).toContain("opened each tested shape of this markup clean");
        expect(explanation).toContain("other shapes are untested");
        expect(explanation).not.toMatch(/session|s33/);
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

    it("explains duplicates revealed by markup compatibility (#635)", () => {
        const explanation = explainWordError("rels-id-duplicate");
        expect(explanation).toContain("markup compatibility");
        expect(explanation).toContain("Fallback");
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
