/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

import { describe, expect, it } from "vitest";
import { runChecks } from "../src/lib/check-scheduler";
import { NS } from "../src/lib/types";
import { ValidationContext } from "../src/lib/validation-context";
import { createDefaultCheckRegistry, WORD_RULE_CHECKS } from "../src/scripts/office/checks";
import { MINIMAL_PARTS } from "./support/minimal-docx";
import {
    fallback,
    MCE_DECLARATIONS,
    RELATIONSHIPS_MCE_CASES,
    RELATIONSHIPS_MCE_CODE,
} from "./support/relationships-markup-compatibility-specimens";

const CODE = "opc-relationships-structure";
const diagnose = async (parts: Readonly<Record<string, string>>) => {
    const check = createDefaultCheckRegistry().get(CODE);
    expect(check, "relationships structure is a registered Word rule").toBeDefined();
    return check!.run(ValidationContext.fromEntries(parts));
};

describe(CODE, () => {
    it.each(RELATIONSHIPS_MCE_CASES)("judges the processed children of $label (#632)", async ({ parts, part, clean }) => {
        expect(await diagnose(parts)).toEqual([
            expect.objectContaining({ code: clean ? RELATIONSHIPS_MCE_CODE : CODE, path: part, severity: "error" }),
        ]);
    });

    it.each([
        { label: "foreign child in the selected Fallback", content: fallback("<x:Wrap/>"), code: CODE },
        {
            label: "foreign child in the understood Choice",
            content: '<mc:AlternateContent><mc:Choice Requires="w14"><x:Wrap/></mc:Choice><mc:Fallback/></mc:AlternateContent>',
            code: CODE,
        },
        {
            label: "package-namespace non-Relationship in the Fallback",
            content: fallback("<Other/>"),
            code: CODE,
        },
        {
            label: "Ignorable package-namespace non-Relationship",
            content: `<p:Other xmlns:p="${NS.PR}" mc:Ignorable="p"/>`,
            code: CODE,
        },
        {
            label: "Ignorable child nested inside a Fallback",
            content: fallback('<x:Wrap mc:Ignorable="x"/>'),
            code: RELATIONSHIPS_MCE_CODE,
        },
        {
            label: "AlternateContent selecting nothing",
            content: '<mc:AlternateContent><mc:Choice Requires="foo"><x:Wrap/></mc:Choice></mc:AlternateContent>',
            code: RELATIONSHIPS_MCE_CODE,
        },
        {
            label: "foreign child in an unselected branch",
            content: '<mc:AlternateContent><mc:Choice Requires="foo"><x:Wrap/></mc:Choice><mc:Fallback/></mc:AlternateContent>',
            code: RELATIONSHIPS_MCE_CODE,
        },
    ])("classifies $label without mutating the stored view", async ({ content, code }) => {
        const xml = `<Relationships xmlns="${NS.PR}" ${MCE_DECLARATIONS}>${content}</Relationships>`;
        const ctx = ValidationContext.fromEntries({ "_rels/.rels": xml });
        const before = ctx.xml("_rels/.rels");
        const serialized = before?.ok ? before.document.toString() : undefined;
        expect(await createDefaultCheckRegistry().get(CODE)!.run(ctx)).toEqual([
            expect.objectContaining({ code, path: "_rels/.rels", severity: "error" }),
        ]);
        expect(before?.ok && before.document.toString()).toBe(serialized);
    });

    it("distinguishes a removed sibling from a surviving foreign child", async () => {
        const xml = `<Relationships xmlns="${NS.PR}" ${MCE_DECLARATIONS}><x:Ignored mc:Ignorable="x"/><x:Wrap/></Relationships>`;
        const issues = await diagnose({ "_rels/.rels": xml });
        expect(issues).toHaveLength(2);
        expect(issues).toEqual(
            expect.arrayContaining([
                expect.objectContaining({ code: RELATIONSHIPS_MCE_CODE, message: expect.stringContaining("Ignored") }),
                expect.objectContaining({ code: CODE, message: expect.stringContaining("Wrap") }),
            ]),
        );
    });

    it("keeps a wrong raw root blocking even when Ignorable removes it", async () => {
        expect(await diagnose({ "_rels/.rels": `<x:Relationships ${MCE_DECLARATIONS} mc:Ignorable="x"/>` })).toEqual([
            expect.objectContaining({ code: CODE, message: expect.stringContaining("root") }),
        ]);
    });

    it("registers in relationships with XML syntax as a prerequisite", () => {
        expect(WORD_RULE_CHECKS.find((check) => check.id === CODE)).toMatchObject({
            groups: ["relationships"],
            formats: ["docx"],
            requires: ["xml-syntax"],
        });
    });

    it.each(["_rels/.rels", "_RELS/.RELS", "_rels/.%72els", "word/_rels/document.xml.rels", "WORD/_RELS/document.xml.RELS"])(
        "reports each foreign child at the stored part name %s",
        async (part) => {
            const xml = `<Relationships xmlns="${NS.PR}"><x:Extra xmlns:x="urn:x"/><Relationship/><Relationship xmlns="urn:x"/></Relationships>`;
            const issues = await diagnose({ [part]: xml });
            expect(issues).toEqual([
                expect.objectContaining({ code: CODE, path: part, severity: "error", message: expect.stringContaining("Extra") }),
                expect.objectContaining({ code: CODE, path: part, severity: "error", message: expect.stringContaining("urn:x") }),
            ]);
        },
    );

    it.each([`<Foo xmlns="${NS.PR}"/>`, '<Relationships xmlns="urn:x"/>', "<Relationships/>", `<Relationship xmlns="${NS.PR}"/>`])(
        "rejects the wrong root: %s",
        async (xml) => {
            expect(await diagnose({ "_rels/.rels": xml })).toEqual([
                expect.objectContaining({ code: CODE, path: "_rels/.rels", severity: "error", message: expect.stringContaining("root") }),
            ]);
        },
    );

    it("accepts direct relationships, empty roots, comments and text without mutating XML", async () => {
        const parts = {
            ...MINIMAL_PARTS,
            "empty.rels": `<p:Relationships xmlns:p="${NS.PR}"> \n<!-- comment --><?ok yes?></p:Relationships>`,
        };
        const before = { ...parts };
        expect(await diagnose(parts)).toEqual([]);
        expect(parts).toEqual(before);
    });

    it("ignores foreign elements in non-relationships XML", async () => {
        expect(await diagnose({ "word/extra.xml": "<Foo/>" })).toEqual([]);
    });

    it("leaves malformed parts to xml-syntax while judging well-formed parts", async () => {
        const registry = createDefaultCheckRegistry();
        expect(registry.get(CODE)).toBeDefined();
        const ctx = ValidationContext.fromEntries({
            "_rels/.rels": "<Relationships",
            "word/_rels/document.xml.rels": "<Foo/>",
        });
        const run = await runChecks(ctx, registry, { format: "docx", select: [CODE] });
        expect(run.issues.filter((issue) => issue.path === "_rels/.rels").map((issue) => issue.code)).toEqual(["xml-syntax"]);
        expect(run.issues.filter((issue) => issue.code === CODE)).toEqual([
            expect.objectContaining({ path: "word/_rels/document.xml.rels", severity: "error" }),
        ]);
        expect(await diagnose({ "_rels/.rels": "<Relationships" })).toEqual([]);
    });
});
