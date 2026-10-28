/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

import { writeFileSync } from "node:fs";
import path from "node:path";
import { withTempDir } from "../src/lib/run-cli";
import { validate } from "../src/scripts/office/validate";
import { describe, expect, it } from "vitest";
import { ValidationContext } from "../src/lib/validation-context";
import { createDefaultCheckRegistry } from "../src/scripts/office/checks";
import { isWordBlockingIssue } from "../src/scripts/office/validators/word-blocking";
import { wordObservedContentTypeParts } from "../src/scripts/office/checks/word-part-content-type-unrecognized";
import { buildDocx, MINIMAL_PARTS } from "./support/minimal-docx";

const ID = "word-part-content-type-unrecognized";
const OCTET_STREAM = "application/octet-stream";
const STYLES_CT = "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml";
const SETTINGS_CT = "application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml";
const CT = "[Content_Types].xml";
const SIDECAR = "word/_rels/document.xml.rels";
const STYLES_OVERRIDE = `<Override PartName="/word/styles.xml" ContentType="${STYLES_CT}"/>`;
const R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

const replaceOnce = (text: string, from: string, to: string): string => {
    if (text.split(from).length !== 2) throw new Error(`expected exactly one ${from}`);
    return text.replace(from, to);
};

/** Session 46's styles part (s46e, #707): the minimal styles plus a visible Heading1. */
const STYLES_XML = replaceOnce(
    MINIMAL_PARTS["word/styles.xml"]!,
    "</w:styles>",
    `<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/></w:style></w:styles>`,
);

/** Session 46's `stylesBin` (s46e): the styles relationship names word/styles.bin; `asStyles` types it. */
const stylesBin = (bytes: string | Uint8Array, asStyles: boolean): Record<string, string | Uint8Array> => {
    const base = Object.fromEntries(Object.entries(MINIMAL_PARTS).filter(([name]) => name !== "word/styles.xml"));
    const types = replaceOnce(
        MINIMAL_PARTS[CT]!,
        STYLES_OVERRIDE,
        asStyles ? `<Override PartName="/word/styles.bin" ContentType="${STYLES_CT}"/>` : "",
    );
    return {
        ...base,
        [SIDECAR]: replaceOnce(MINIMAL_PARTS[SIDECAR]!, 'Target="styles.xml"', 'Target="styles.bin"'),
        "word/styles.bin": bytes,
        [CT]: asStyles
            ? types
            : replaceOnce(
                  types,
                  '<Default Extension="rels"',
                  `<Default Extension="bin" ContentType="${OCTET_STREAM}"/><Default Extension="rels"`,
              ),
    };
};

/** The `stylesBin` shape with the part's Override carrying `declared` instead of the styles type. */
const stylesTyped = (declared: string): Record<string, string | Uint8Array> => {
    const parts = stylesBin(STYLES_XML, true);
    return { ...parts, [CT]: (parts[CT] as string).replace(`ContentType="${STYLES_CT}"`, `ContentType="${declared}"`) };
};

/** The `stylesBin` shape with the `bin` Default carrying `declared` instead of octet-stream. */
const stylesDefaultTyped = (declared: string): Record<string, string | Uint8Array> => {
    const parts = stylesBin(STYLES_XML, false);
    return { ...parts, [CT]: (parts[CT] as string).replace(`ContentType="${OCTET_STREAM}"`, `ContentType="${declared}"`) };
};

/** A related settings part stored as word/settings.bin and typed application/octet-stream by Default (unobserved, #707). */
const settingsBin = (): Record<string, string | Uint8Array> => ({
    ...MINIMAL_PARTS,
    [SIDECAR]: replaceOnce(
        MINIMAL_PARTS[SIDECAR]!,
        "</Relationships>",
        `<Relationship Id="rId2" Type="${R}/settings" Target="settings.bin"/></Relationships>`,
    ),
    "word/settings.bin": `<w:settings xmlns:w="${W}"/>`,
    [CT]: replaceOnce(
        MINIMAL_PARTS[CT]!,
        '<Default Extension="rels"',
        `<Default Extension="bin" ContentType="${OCTET_STREAM}"/><Default Extension="rels"`,
    ),
});

const diagnose = async (parts: Record<string, string | Uint8Array>) => {
    const check = createDefaultCheckRegistry().get(ID);
    expect(check?.id).toBe(ID);
    return (await check?.run(ValidationContext.fromEntries(parts))) ?? [];
};

/** The observed reach a word-valid caller resolves for `isWordBlockingIssue` (#718). */
const observedParts = (parts: Record<string, string | Uint8Array>) => wordObservedContentTypeParts(ValidationContext.fromEntries(parts));

// Session 46 (s46e): both refused shapes, their clean control, and the disjointness boundary.
describe("unrecognized part content type (#707, session 46)", () => {
    it("registers a DOCX content-types check behind xml-syntax", () => {
        expect(createDefaultCheckRegistry().get(ID)).toMatchObject({
            id: ID,
            groups: ["content-types"],
            formats: ["docx"],
            requires: ["xml-syntax"],
        });
    });

    it.each([
        ["non-XML bytes", "this is not XML"],
        ["styles XML", STYLES_XML],
    ] as const)("rejects word/styles.bin holding %s typed application/octet-stream", async (_label, bytes) => {
        const parts = stylesBin(bytes, false);
        const found = await diagnose(parts);
        expect(found).toHaveLength(1);
        expect(found[0]).toMatchObject({ code: ID, severity: "error", path: "word/styles.bin" });
        expect(found[0]?.message).toContain("the main document's styles relationship");
        expect(found[0]?.message).toContain(OCTET_STREAM);
        expect(found[0]?.message).toContain(STYLES_CT);
        expect(found[0]?.message).toContain("Word 16.115 offers to recover the package (session 46)");
        expect(isWordBlockingIssue(found[0]!, new Set(), new Set(), new Set(), await observedParts(parts))).toBe(true);
        // The styles-typed control opens clean.
        expect(await diagnose(stylesBin(STYLES_XML, true))).toEqual([]);
    });

    it("leaves the minimal package and the other checks' classes to them", async () => {
        expect(await diagnose({ ...MINIMAL_PARTS })).toEqual([]);
        // A generic XML type is word-part-content-type-generic's finding.
        expect(await diagnose(stylesDefaultTyped("application/xml"))).toEqual([]);
        // Another mapped Word type is word-part-content-type-mismatch's class.
        expect(await diagnose(stylesTyped(SETTINGS_CT))).toEqual([]);
        // An inexact spelling of the part's own type is word-part-content-type-inexact's finding.
        expect(await diagnose(stylesTyped(`${STYLES_CT}; charset=utf-8`))).toEqual([]);
        // An empty declared type is word-content-type-invalid's class.
        expect(await diagnose(stylesTyped(""))).toEqual([]);
    });

    it("keeps the four part content-type checks disjoint", async () => {
        for (const [parts, expected] of [
            [stylesDefaultTyped("application/xml"), "word-part-content-type-generic"],
            [stylesTyped(`${STYLES_CT}; charset=utf-8`), "word-part-content-type-inexact"],
            [stylesBin(STYLES_XML, false), ID],
        ] as const) {
            const ctx = ValidationContext.fromEntries(parts);
            const issues = (
                await Promise.all(
                    ["word-part-content-type-generic", "word-part-content-type-inexact", "word-part-content-type-mismatch", ID].map(
                        async (id) => (await createDefaultCheckRegistry().get(id)?.run(ctx)) ?? [],
                    ),
                )
            ).flat();
            expect(issues.map((issue) => issue.code)).toEqual([expected]);
        }
    });

    it("blocks only the observed reach: a related settings part is an error without the recovery claim", async () => {
        const parts = settingsBin();
        const found = await diagnose(parts);
        expect(found).toHaveLength(1);
        expect(found[0]).toMatchObject({ code: ID, severity: "error", path: "word/settings.bin" });
        expect(found[0]?.message).toContain("the main document's settings relationship");
        expect(found[0]?.message).toContain("clause 11.3");
        expect(found[0]?.message).not.toContain("Word 16.115");
        // The observed set holds only the styles reach, so the finding stays a
        // word-valid warning even though its message could carry any text.
        expect(isWordBlockingIssue(found[0]!, new Set(), new Set(), new Set(), await observedParts(parts))).toBe(false);
    });
});

// File-backed profile checks are integration tests; the rows above use in-memory package entries.
describe("unrecognized profile scope (#707)", () => {
    it.each([
        ["non-XML bytes", "this is not XML"],
        ["styles XML", STYLES_XML],
    ] as const)(
        "judges word/styles.bin holding %s in full profiles",
        async (_label, bytes) => {
            await withTempDir(async (dir) => {
                const file = path.join(dir, "styles-bin.docx");
                writeFileSync(file, await buildDocx(stylesBin(bytes, false)));
                for (const profile of ["word-valid", "strict", "lenient"] as const) {
                    const result = await validate(file, { profile });
                    const found = result.issues.filter((issue) => issue.code === ID);
                    expect(found, profile).toHaveLength(1);
                    expect(found[0]?.severity).toBe(profile === "lenient" ? "warning" : "error");
                    if (profile !== "lenient") expect(result.valid).toBe(false);
                    if (profile === "word-valid") expect(found[0]?.message).toContain("session 46");
                }
                writeFileSync(file, await buildDocx(stylesBin(STYLES_XML, true)));
                const control = await validate(file);
                expect(control.issues.filter((issue) => issue.code === ID)).toEqual([]);
                expect(control.valid).toBe(true);
            });
        },
        60_000,
    );

    it("warns under word-valid and lenient, errors under strict, for a related settings part (unobserved)", async () => {
        await withTempDir(async (dir) => {
            const file = path.join(dir, "settings-bin.docx");
            writeFileSync(file, await buildDocx(settingsBin()));
            const wordValid = await validate(file);
            const found = wordValid.issues.filter((issue) => issue.code === ID);
            expect(found).toHaveLength(1);
            expect(found[0]?.severity).toBe("warning");
            const strict = await validate(file, { profile: "strict" });
            expect(strict.issues.filter((issue) => issue.code === ID)).toEqual([expect.objectContaining({ severity: "error" })]);
            const lenient = await validate(file, { profile: "lenient" });
            expect(lenient.issues.filter((issue) => issue.code === ID)).toEqual([expect.objectContaining({ severity: "warning" })]);
        });
    }, 60_000);

    it("full, quick and named runs agree on the refused shape (#381)", async () => {
        await withTempDir(async (dir) => {
            const file = path.join(dir, "styles-bin.docx");
            writeFileSync(file, await buildDocx(stylesBin("this is not XML", false)));
            const full = await validate(file);
            const found = full.issues.filter(({ code }) => code === ID);
            expect(full.valid).toBe(false);
            expect(found).toHaveLength(1);
            for (const selection of [{ quick: true }, { checks: [ID] }, { checks: ["content-types"] }]) {
                const partial = await validate(file, selection);
                expect(partial.valid).toBe(full.valid);
                expect(partial.issues.filter(({ code }) => code === ID)).toEqual(found);
            }
        });
    }, 60_000);
});

// #718: the message embeds the declared content type and the stored part name
// verbatim, so a package can carry the recovery sentence itself; the verdict
// must read the reach structurally, never that text.
describe("unrecognized blocking verdict is structural (#718)", () => {
    const SENTENCE = "Word 16.115 offers to recover the package";

    /** The minimal package plus a related webSettings part — an unobserved reach — stored as `part`, its Override carrying `type`. */
    const webSettingsTyped = (part: string, type: string): Record<string, string> => ({
        ...MINIMAL_PARTS,
        [part]: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:webSettings xmlns:w="${W}"/>`,
        [SIDECAR]: replaceOnce(
            MINIMAL_PARTS[SIDECAR]!,
            "</Relationships>",
            `<Relationship Id="rId2" Type="${R}/webSettings" Target="${part.slice("word/".length)}"/></Relationships>`,
        ),
        [CT]: replaceOnce(MINIMAL_PARTS[CT]!, "</Types>", `<Override PartName="/${part}" ContentType="${type}"/></Types>`),
    });

    it.each([
        ["the declared content type", `application/x-${SENTENCE}`, "word/webSettings.xml"],
        ["the stored part name", OCTET_STREAM, `word/${SENTENCE}.bin`],
    ] as const)(
        "keeps an unobserved webSettings reach a word-valid warning when the sentence rides %s",
        async (_label, declared, part) => {
            await withTempDir(async (dir) => {
                const file = path.join(dir, "crafted.docx");
                writeFileSync(file, await buildDocx(webSettingsTyped(part, declared)));
                const full = await validate(file);
                const found = full.issues.filter((issue) => issue.code === ID);
                expect(found).toHaveLength(1);
                expect(found[0]?.severity).toBe("warning");
                expect(found[0]?.message).toContain(SENTENCE);
                // A named run reaches the same verdict (#381).
                const partial = await validate(file, { checks: [ID] });
                expect(partial.issues.filter((issue) => issue.code === ID)).toEqual(found);
            });
        },
        60_000,
    );
});
