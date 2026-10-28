/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

import { describe, expect, it } from "vitest";
import { ValidationContext } from "../src/lib/validation-context";
import { createDefaultCheckRegistry, WORD_RULE_CHECKS } from "../src/scripts/office/checks";
import { MINIMAL_PARTS } from "./support/minimal-docx";

const CODE = "opc-relationships-target-mode";
const PACKAGE = "_rels/.rels";
const SIDECAR = "word/_rels/document.xml.rels";
const PR = "http://schemas.openxmlformats.org/package/2006/relationships";
const R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const relationship = (mode: string) =>
    `<Relationship Id="rId99" Type="${R}/hyperlink" Target="https://example.com/" TargetMode="${mode}"/>`;
const append = (xml: string, child: string) => xml.replace("</Relationships>", `${child}</Relationships>`);
const diagnose = (parts: Readonly<Record<string, string>>) =>
    createDefaultCheckRegistry().get(CODE)!.run(ValidationContext.fromEntries(parts));

describe(CODE, () => {
    it("is a registered Word rule requiring XML syntax", () => {
        const check = createDefaultCheckRegistry().get(CODE)!;
        expect(WORD_RULE_CHECKS).toContain(check);
        expect(check.requires).toEqual(["xml-syntax"]);
        expect(check.groups).toEqual(["relationships"]);
        expect(check.formats).toEqual(["docx"]);
    });

    it.each(["external", "External ", ""])("reports the exact spelling %j, including an empty present attribute", async (mode) => {
        expect(await diagnose({ ...MINIMAL_PARTS, [SIDECAR]: append(MINIMAL_PARTS[SIDECAR]!, relationship(mode)) })).toEqual([
            expect.objectContaining({
                code: CODE,
                path: SIDECAR,
                severity: "error",
                message: expect.stringContaining(JSON.stringify(mode)),
            }),
        ]);
    });

    it("reports each offending Relationship and leaves the stored view untouched", async () => {
        const ctx = ValidationContext.fromEntries({
            ...MINIMAL_PARTS,
            [SIDECAR]: append(MINIMAL_PARTS[SIDECAR]!, relationship("external") + relationship("")),
        });
        const before = ctx.xml(SIDECAR);
        const serialized = before?.ok && before.document.toString();
        const issues = await createDefaultCheckRegistry().get(CODE)!.run(ctx);
        expect(issues).toHaveLength(2);
        expect(issues[0]!.message).toContain("rId99");
        expect(issues[0]!.message).toContain("Internal");
        expect(issues[0]!.message).toContain("External");
        expect(before?.ok && before.document.toString()).toBe(serialized);
    });

    it("finds the main document by relationship and compares its sidecar as OPC does", async () => {
        const parts = { ...MINIMAL_PARTS };
        delete parts["word/document.xml"];
        delete parts[SIDECAR];
        parts["stories/main.xml"] = MINIMAL_PARTS["word/document.xml"]!;
        parts["stories/_RELS/MAIN.%78ml.RELS"] = `<Relationships xmlns="${PR}">${relationship("external")}</Relationships>`;
        parts[PACKAGE] = MINIMAL_PARTS[PACKAGE]!.replace("word/document.xml", "stories/main.xml");
        expect(await diagnose(parts)).toEqual([
            expect.objectContaining({ code: CODE, path: "stories/_RELS/MAIN.%78ml.RELS", severity: "error" }),
        ]);
    });

    it("compares the package relationships name as OPC does, even without a main document", async () => {
        const part = "_RELS/.%72els";
        expect(await diagnose({ [part]: `<Relationships xmlns="${PR}">${relationship("external")}</Relationships>` })).toEqual([
            expect.objectContaining({ code: CODE, path: part, severity: "error" }),
        ]);
    });

    it("does not treat a reached glossary as a main document", async () => {
        const part = "word/glossary/_rels/document.xml.rels";
        expect(
            await diagnose({
                ...MINIMAL_PARTS,
                [SIDECAR]: append(
                    MINIMAL_PARTS[SIDECAR]!,
                    `<Relationship Id="rIdG" Type="${R}/glossaryDocument" Target="glossary/document.xml"/>`,
                ),
                "word/glossary/document.xml": `<w:glossaryDocument xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:docParts/></w:glossaryDocument>`,
                [part]: `<Relationships xmlns="${PR}">${relationship("external")}</Relationships>`,
            }),
        ).toEqual([]);
    });

    it.each(["Internal", "External"])("accepts %s and absent TargetMode", async (mode) => {
        expect(await diagnose({ ...MINIMAL_PARTS, [SIDECAR]: append(MINIMAL_PARTS[SIDECAR]!, relationship(mode)) })).toEqual([]);
    });

    it.each([
        "word/_rels/header1.xml.rels",
        "word/glossary/_rels/document.xml.rels",
        "word/_rels/footnotes.xml.rels",
        "word/_rels/comments.xml.rels",
    ])("keeps the existing verdict in %s", async (part) => {
        expect(
            await diagnose({ ...MINIMAL_PARTS, [part]: `<Relationships xmlns="${PR}">${relationship("external")}</Relationships>` }),
        ).toEqual([]);
    });

    it.each([
        "<Relationships",
        `<Other xmlns="${PR}">${relationship("external")}</Other>`,
        `<Relationships xmlns="urn:other">${relationship("external")}</Relationships>`,
    ])("leaves malformed XML and wrong roots to their own checks: %s", async (xml) => {
        expect(await diagnose({ ...MINIMAL_PARTS, [SIDECAR]: xml })).toEqual([]);
    });

    it("does not require a main-document sidecar", async () => {
        const parts = { ...MINIMAL_PARTS };
        delete parts[SIDECAR];
        expect(await diagnose(parts)).toEqual([]);
    });

    it("reads only direct Relationship children in the package namespace", async () => {
        expect(
            await diagnose({
                ...MINIMAL_PARTS,
                [SIDECAR]: append(
                    MINIMAL_PARTS[SIDECAR]!,
                    `<Wrap>${relationship("external")}</Wrap><x:Relationship xmlns:x="urn:other" TargetMode="external"/>`,
                ),
            }),
        ).toEqual([]);
    });
});
