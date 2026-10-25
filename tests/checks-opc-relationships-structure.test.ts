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
    PACKAGE_RELS,
    RELATIONSHIPS_MCE_CASES,
    RELATIONSHIPS_MCE_CODE,
    ROOT_MCE_SHAPES,
    rootMceRelationshipsParts,
    rootMceWrappedRelationshipsParts,
    rootMceXml,
    rootUnselectedChoiceRelationshipsParts,
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

    it("states the observed Word behaviour without generalizing it (#632 review finding 5)", async () => {
        const issues = await diagnose(RELATIONSHIPS_MCE_CASES[0]!.parts);
        expect(issues).toHaveLength(1);
        const message = issues[0]!.message;
        expect(message).toContain("removed or replaced by markup compatibility, leaving only Relationship children");
        // Word was shown the four shapes of s33d/s33e that this code reports;
        // each opened clean. Nothing else was tested, so say that, without
        // internal session ids in user-facing text.
        expect(message).toContain("Microsoft Word opened each tested shape of this markup clean; other shapes are untested.");
        expect(message).not.toMatch(/session|s33/);
        expect(message).not.toContain("Microsoft Word accepts this markup");
        expect(message).toContain("Strict OPC requires direct Relationship children in the stored XML");
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

    it.each(ROOT_MCE_SHAPES)(
        "accepts a package relationships root that a %s wrapper resolves to one Relationships element (#646)",
        async (shape) => {
            const parts = rootMceWrappedRelationshipsParts(PACKAGE_RELS, shape);
            const ctx = ValidationContext.fromEntries(parts);
            const before = ctx.xml(PACKAGE_RELS);
            const serialized = before?.ok ? before.document.toString() : undefined;
            expect(await diagnose(parts)).toEqual([]);
            expect(before?.ok && before.document.toString()).toBe(serialized);
        },
    );

    it.each(["_RELS/.RELS", "_rels/.%72els"])(
        "accepts the resolved root under the equivalent package-relationships spelling %s (#646)",
        async (part) => {
            const wrapped = rootMceWrappedRelationshipsParts(PACKAGE_RELS)[PACKAGE_RELS]!;
            const parts = Object.fromEntries(Object.entries(MINIMAL_PARTS).filter(([name]) => name !== PACKAGE_RELS));
            expect(await diagnose({ ...parts, [part]: wrapped })).toEqual([]);
        },
    );

    it("still reports a surviving foreign child through the resolved root (#646)", async () => {
        const rels = `<Relationships xmlns="${NS.PR}"><Relationship Id="rId1" Type="${NS.R}/officeDocument" Target="word/document.xml"/><x:Wrap/></Relationships>`;
        expect(await diagnose({ ...MINIMAL_PARTS, "_rels/.rels": rootMceXml("fallback", rels) })).toEqual([
            expect.objectContaining({ code: CODE, path: "_rels/.rels", severity: "error", message: expect.stringContaining("Wrap") }),
        ]);
    });

    it.each([
        { label: "two Relationship elements", parts: () => rootMceRelationshipsParts(PACKAGE_RELS, "fallback") },
        { label: "a single foreign element", parts: () => ({ ...MINIMAL_PARTS, "_rels/.rels": rootMceXml("fallback", "<x:Foo/>") }) },
        { label: "no element at all", parts: () => ({ ...MINIMAL_PARTS, "_rels/.rels": rootMceXml("fallback", "") }) },
        {
            label: "no element, its Relationships in an unselected Choice",
            parts: () => rootUnselectedChoiceRelationshipsParts(PACKAGE_RELS),
        },
    ])("keeps the wrong-root finding when the package root resolves to $label (#640, #646)", async ({ parts }) => {
        expect(await diagnose(parts())).toEqual([
            expect.objectContaining({ code: CODE, path: PACKAGE_RELS, severity: "error", message: expect.stringContaining("root") }),
        ]);
    });

    it("keeps the stored-root verdict for a root-wrapped sidecar, which Word was not shown (#646)", async () => {
        expect(await diagnose(rootMceWrappedRelationshipsParts("word/_rels/document.xml.rels"))).toEqual([
            expect.objectContaining({
                code: CODE,
                path: "word/_rels/document.xml.rels",
                severity: "error",
                message: expect.stringContaining("root"),
            }),
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
