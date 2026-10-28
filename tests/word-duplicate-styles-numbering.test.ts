/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

import path from "node:path";
import { fileURLToPath } from "node:url";

import { describe, expect, it } from "vitest";

import type { CheckDefinition } from "../src/lib/check-registry";
import { ValidationContext } from "../src/lib/validation-context";
import { WORD_RULE_CHECKS } from "../src/scripts/office/checks/index";
import { validate } from "../src/scripts/office/validate";
import { rels, relation, R, RS } from "./support/word-rule-test-xml";

const HERE = path.dirname(fileURLToPath(import.meta.url));
const BROKEN_STYLES = path.join(HERE, "fixtures", "broken", "package.word-duplicate-styles.docx");
const CLEAN_STYLES = path.join(HERE, "fixtures", "working", "package.word-duplicate-styles-removed.docx");
const BROKEN_NUMBERING = path.join(HERE, "fixtures", "broken", "package.word-duplicate-numbering.docx");
const CLEAN_NUMBERING = path.join(HERE, "fixtures", "working", "package.word-duplicate-numbering-removed.docx");

const STYLES_WITH_EFFECTS = "http://schemas.microsoft.com/office/2007/relationships/stylesWithEffects";

const relationOfType = (id: string, type: string, target = "part.xml"): string =>
    `<Relationship Id="${id}" Type="${type}" Target="${target}"/>`;

const registeredCheck = (id: string): CheckDefinition => {
    const check = WORD_RULE_CHECKS.find((rule) => rule.id === id);
    expect(check?.id, `the ${id} check is registered in WORD_RULE_CHECKS`).toBe(id);
    return check!;
};

describe("word-valid duplicate styles relationship", { timeout: 30000 }, () => {
    it("rejects a document relationships part with two styles relationships", async () => {
        const result = await validate(BROKEN_STYLES, { profile: "word-valid" });
        expect(result.valid).toBe(false);
        expect(result.issues.some((issue) => issue.severity === "error" && issue.code === "word-duplicate-styles-relationship")).toBe(true);
    });

    it("does not reject the document once only one styles relationship remains", async () => {
        const result = await validate(CLEAN_STYLES, { profile: "word-valid" });
        expect(result.issues.some((issue) => issue.code === "word-duplicate-styles-relationship")).toBe(false);
        expect(result.valid).toBe(true);
    });

    it("does not report the cause on lenient", async () => {
        const result = await validate(BROKEN_STYLES, { profile: "lenient" });
        expect(result.issues.some((issue) => issue.code === "word-duplicate-styles-relationship")).toBe(false);
    });

    // strict includes every Word-valid check (validation contract, Task 1).
    it("reports the cause on strict", async () => {
        const result = await validate(BROKEN_STYLES, { profile: "strict" });
        expect(result.issues.some((issue) => issue.severity === "error" && issue.code === "word-duplicate-styles-relationship")).toBe(true);
    });
});

describe("word-valid duplicate numbering relationship", { timeout: 30000 }, () => {
    it("rejects a document relationships part with two numbering relationships", async () => {
        const result = await validate(BROKEN_NUMBERING, { profile: "word-valid" });
        expect(result.valid).toBe(false);
        expect(result.issues.some((issue) => issue.severity === "error" && issue.code === "word-duplicate-numbering-relationship")).toBe(
            true,
        );
    });

    it("does not reject the document once only one numbering relationship remains", async () => {
        const result = await validate(CLEAN_NUMBERING, { profile: "word-valid" });
        expect(result.issues.some((issue) => issue.code === "word-duplicate-numbering-relationship")).toBe(false);
        expect(result.valid).toBe(true);
    });

    it("does not report the cause on lenient", async () => {
        const result = await validate(BROKEN_NUMBERING, { profile: "lenient" });
        expect(result.issues.some((issue) => issue.code === "word-duplicate-numbering-relationship")).toBe(false);
    });

    // strict includes every Word-valid check (validation contract, Task 1).
    it("reports the cause on strict", async () => {
        const result = await validate(BROKEN_NUMBERING, { profile: "strict" });
        expect(result.issues.some((issue) => issue.severity === "error" && issue.code === "word-duplicate-numbering-relationship")).toBe(
            true,
        );
    });
});

describe("word-duplicate-styles-relationship", () => {
    it("reports the Word-rejected shape with unchanged issue details", async () => {
        const ctx = ValidationContext.fromEntries({
            "word/_rels/document.xml.rels": rels(relation("rId1", "styles") + relation("rId2", "styles")),
        });
        expect(await registeredCheck("word-duplicate-styles-relationship").run(ctx)).toEqual([
            {
                severity: "error",
                code: "word-duplicate-styles-relationship",
                path: "word/_rels/document.xml.rels",
                message:
                    "word/_rels/document.xml.rels contains 2 styles relationships (rId1, rId2). Microsoft Word offers to recover the file. Keep one styles relationship.",
            },
        ]);
    });
    it("still reports two styles relationships to the same target", async () => {
        const ctx = ValidationContext.fromEntries({
            "word/_rels/document.xml.rels": rels(
                relation("rId1", "styles", "", "styles.xml") + relation("rId2", "styles", "", "styles.xml"),
            ),
        });
        expect(await registeredCheck("word-duplicate-styles-relationship").run(ctx)).toEqual([
            {
                severity: "error",
                code: "word-duplicate-styles-relationship",
                path: "word/_rels/document.xml.rels",
                message:
                    "word/_rels/document.xml.rels contains 2 styles relationships (rId1, rId2). Microsoft Word offers to recover the file. Keep one styles relationship.",
            },
        ]);
    });
    it("draws nothing from one styles relationship beside a Microsoft stylesWithEffects relationship", async () => {
        const ctx = ValidationContext.fromEntries({
            "word/_rels/document.xml.rels": rels(
                relation("rId1", "styles") + relationOfType("rId2", STYLES_WITH_EFFECTS, "stylesWithEffects.xml"),
            ),
        });
        expect(await registeredCheck("word-duplicate-styles-relationship").run(ctx)).toEqual([]);
    });
    // A type counts only as the namespace plus the one segment — Word reads
    // Types exactly (session 27) — so `custom/styles` is no styles type
    // (#519), and a foreign styles-suffixed type is none either.
    it("draws nothing from one styles relationship beside a `custom/styles` type", async () => {
        const ctx = ValidationContext.fromEntries({
            "word/_rels/document.xml.rels": rels(relation("rId1", "styles") + relationOfType("rId2", `${R}/custom/styles`, "custom.xml")),
        });
        expect(await registeredCheck("word-duplicate-styles-relationship").run(ctx)).toEqual([]);
    });
    it("draws nothing from one styles relationship beside a foreign styles-suffixed type", async () => {
        const ctx = ValidationContext.fromEntries({
            "word/_rels/document.xml.rels": rels(relation("rId1", "styles") + relationOfType("rId2", "http://example.com/styles")),
        });
        expect(await registeredCheck("word-duplicate-styles-relationship").run(ctx)).toEqual([]);
    });
    it("draws nothing from one styles relationship beside a footnotes relationship", async () => {
        const ctx = ValidationContext.fromEntries({
            "word/_rels/document.xml.rels": rels(relation("rId1", "styles") + relation("rId2", "footnotes")),
        });
        expect(await registeredCheck("word-duplicate-styles-relationship").run(ctx)).toEqual([]);
    });
    it("counts the ISO Strict styles type URI", async () => {
        const strict = `${RS}/styles`;
        const ctx = ValidationContext.fromEntries({
            "word/_rels/document.xml.rels": rels(
                relationOfType("rId1", strict, "styles.xml") + relationOfType("rId2", strict, "styles.xml"),
            ),
        });
        expect(await registeredCheck("word-duplicate-styles-relationship").run(ctx)).toEqual([
            {
                severity: "error",
                code: "word-duplicate-styles-relationship",
                path: "word/_rels/document.xml.rels",
                message:
                    "word/_rels/document.xml.rels contains 2 styles relationships (rId1, rId2). Microsoft Word offers to recover the file. Keep one styles relationship.",
            },
        ]);
    });
});

describe("word-duplicate-numbering-relationship", () => {
    it("reports the Word-rejected shape with unchanged issue details", async () => {
        const ctx = ValidationContext.fromEntries({
            "word/_rels/document.xml.rels": rels(relation("rId1", "numbering") + relation("rId2", "numbering")),
        });
        expect(await registeredCheck("word-duplicate-numbering-relationship").run(ctx)).toEqual([
            {
                severity: "error",
                code: "word-duplicate-numbering-relationship",
                path: "word/_rels/document.xml.rels",
                message:
                    "word/_rels/document.xml.rels contains 2 numbering relationships (rId1, rId2). Microsoft Word offers to recover the file. Keep one numbering relationship.",
            },
        ]);
    });
    it("draws nothing from one numbering relationship beside a `custom/numbering` type", async () => {
        const ctx = ValidationContext.fromEntries({
            "word/_rels/document.xml.rels": rels(
                relation("rId1", "numbering") + relationOfType("rId2", `${R}/custom/numbering`, "custom.xml"),
            ),
        });
        expect(await registeredCheck("word-duplicate-numbering-relationship").run(ctx)).toEqual([]);
    });
    it("counts the ISO Strict numbering type URI", async () => {
        const strict = `${RS}/numbering`;
        const ctx = ValidationContext.fromEntries({
            "word/_rels/document.xml.rels": rels(
                relationOfType("rId1", strict, "numbering.xml") + relationOfType("rId2", strict, "numbering.xml"),
            ),
        });
        expect(await registeredCheck("word-duplicate-numbering-relationship").run(ctx)).toEqual([
            {
                severity: "error",
                code: "word-duplicate-numbering-relationship",
                path: "word/_rels/document.xml.rels",
                message:
                    "word/_rels/document.xml.rels contains 2 numbering relationships (rId1, rId2). Microsoft Word offers to recover the file. Keep one numbering relationship.",
            },
        ]);
    });
});

describe("duplicate footnotes relationships stay unchecked", () => {
    // Session 45 (s45g): two footnotes relationships opened clean, so
    // footnotes keeps no duplicate check (#694).
    it("draws nothing from either check for two footnotes relationships", async () => {
        const ctx = ValidationContext.fromEntries({
            "word/_rels/document.xml.rels": rels(relation("rId1", "footnotes") + relation("rId2", "footnotes")),
        });
        expect(await registeredCheck("word-duplicate-styles-relationship").run(ctx)).toEqual([]);
        expect(await registeredCheck("word-duplicate-numbering-relationship").run(ctx)).toEqual([]);
    });
});
