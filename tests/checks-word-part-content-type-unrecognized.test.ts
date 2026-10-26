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
import {
    wordObservedContentTypeParts,
    RECOVERY_SESSIONS,
    UNRECOGNIZED_OBSERVED_SEGMENTS,
} from "../src/scripts/office/checks/word-part-content-type-unrecognized";
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
        expect(found[0]?.message).toContain("Word 16.115 offers to recover the package (sessions 46 and 49)");
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

    it("blocks the related settings reach too: an error with the recovery claim (#721, sessions 50 and 53)", async () => {
        const parts = settingsBin();
        const found = await diagnose(parts);
        expect(found).toHaveLength(1);
        expect(found[0]).toMatchObject({ code: ID, severity: "error", path: "word/settings.bin" });
        expect(found[0]?.message).toContain("the main document's settings relationship");
        expect(found[0]?.message).toContain("Word 16.115 offers to recover the package (sessions 50 and 53)");
        expect(isWordBlockingIssue(found[0]!, new Set(), new Set(), new Set(), await observedParts(parts))).toBe(true);
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
                    if (profile === "word-valid") expect(found[0]?.message).toContain("sessions 46 and 49");
                }
                writeFileSync(file, await buildDocx(stylesBin(STYLES_XML, true)));
                const control = await validate(file);
                expect(control.issues.filter((issue) => issue.code === ID)).toEqual([]);
                expect(control.valid).toBe(true);
            });
        },
        60_000,
    );

    it("warns under word-valid and lenient, errors under strict, for a related webSettings part (unobserved)", async () => {
        await withTempDir(async (dir) => {
            const file = path.join(dir, "websettings-bin.docx");
            writeFileSync(
                file,
                await buildDocx({
                    ...MINIMAL_PARTS,
                    [SIDECAR]: replaceOnce(
                        MINIMAL_PARTS[SIDECAR]!,
                        "</Relationships>",
                        `<Relationship Id="rId2" Type="${R}/webSettings" Target="webSettings.bin"/></Relationships>`,
                    ),
                    "word/webSettings.bin": `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:webSettings xmlns:w="${W}"/>`,
                    [CT]: replaceOnce(
                        MINIMAL_PARTS[CT]!,
                        '<Default Extension="rels"',
                        `<Default Extension="bin" ContentType="${OCTET_STREAM}"/><Default Extension="rels"`,
                    ),
                }),
            );
            const wordValid = await validate(file);
            const found = wordValid.issues.filter((issue) => issue.code === ID);
            expect(found).toHaveLength(1);
            expect(found[0]?.severity).toBe("warning");
            expect(found[0]?.message).toContain("clause 11.3");
            expect(found[0]?.message).not.toContain("Word 16.115");
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

// #719: the invalid-* class belongs to word-content-type-invalid alone, so one
// defect draws one finding.
describe("unrecognized leaves the invalid-* class to word-content-type-invalid (#719)", () => {
    it("reports a styles Override typed invalid-thing once, through word-content-type-invalid", async () => {
        await withTempDir(async (dir) => {
            const file = path.join(dir, "styles-invalid.docx");
            writeFileSync(file, await buildDocx(stylesTyped("invalid-thing")));
            for (const profile of ["word-valid", "strict"] as const) {
                const result = await validate(file, { profile });
                // Strict also judges the stream itself against the OPC schema,
                // as it did before #719; among the part content-type checks the
                // invalid-* class has exactly one owner.
                const contentTypeErrors = result.issues.filter(
                    (issue) =>
                        issue.severity === "error" &&
                        [
                            "word-content-type-invalid",
                            "word-part-content-type-generic",
                            "word-part-content-type-inexact",
                            "word-part-content-type-mismatch",
                            ID,
                        ].includes(issue.code ?? ""),
                );
                expect(contentTypeErrors, profile).toHaveLength(1);
                expect(contentTypeErrors[0], profile).toMatchObject({ code: "word-content-type-invalid" });
                expect(result.valid, profile).toBe(false);
                expect(
                    result.issues.some((issue) => issue.code === ID),
                    profile,
                ).toBe(false);
            }
            // Word-valid runs no XSD, so the invalid-* Override is its only error.
            const wordValid = await validate(file);
            expect(wordValid.issues.filter((issue) => issue.severity === "error")).toHaveLength(1);
        });
    }, 60_000);
});

// #721: sessions 46 (s46e), 49 (s49a/s49b) and 50 (s50c/s50d) observed Word
// 16.115 refuse an unrecognised type on the main document and the parts its
// own relationships relate by styles, settings, fontTable and numbering.
describe("unrecognized observed reach (#721, sessions 46, 49 and 50)", () => {
    const WML = "application/vnd.openxmlformats-officedocument.wordprocessingml";
    const MAIN_TYPE = `${WML}.document.main+xml`;
    const XML_DECL = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n`;

    /** The minimal package with the Override of `part` carrying `type` instead of `own` (s49a/s49b's retypedOverride). */
    const retyped = (part: string, own: string, type: string): Record<string, string | Uint8Array> => ({
        ...MINIMAL_PARTS,
        [CT]: replaceOnce(
            MINIMAL_PARTS[CT]!,
            `<Override PartName="/${part}" ContentType="${own}"/>`,
            `<Override PartName="/${part}" ContentType="${type}"/>`,
        ),
    });
    const mainTyped = (type: string) => retyped("word/document.xml", MAIN_TYPE, type);
    const stylesXmlTyped = (type: string) => retyped("word/styles.xml", STYLES_CT, type);

    /** The s50d shape: the minimal package plus `word/<kind>.xml` related from the main document's sidecar, its Override carrying `type`. */
    const relatedTyped = (kind: "settings" | "fontTable" | "numbering", type: string): Record<string, string | Uint8Array> => {
        const root = { settings: "settings", fontTable: "fonts", numbering: "numbering" }[kind];
        const part = `word/${kind}.xml`;
        return {
            ...MINIMAL_PARTS,
            [part]: `${XML_DECL}<w:${root} xmlns:w="${W}"/>`,
            [SIDECAR]: replaceOnce(
                MINIMAL_PARTS[SIDECAR]!,
                "</Relationships>",
                `<Relationship Id="rId2" Type="${R}/${kind}" Target="${kind}.xml"/></Relationships>`,
            ),
            [CT]: replaceOnce(MINIMAL_PARTS[CT]!, "</Types>", `<Override PartName="/${part}" ContentType="${type}"/></Types>`),
        };
    };

    it.each([
        [
            "main document Override typed application/octet-stream (s49b)",
            mainTyped(OCTET_STREAM),
            "word/document.xml",
            "sessions 49 and 50",
        ],
        [
            "main document Override typed APPLICATION/OCTET-STREAM (s50c)",
            mainTyped("APPLICATION/OCTET-STREAM"),
            "word/document.xml",
            "sessions 49 and 50",
        ],
        [
            "main document Override typed application/octet-stream; charset=utf-8 (s50c)",
            mainTyped("application/octet-stream; charset=utf-8"),
            "word/document.xml",
            "sessions 49 and 50",
        ],
        ["main document Override typed image/png (s50c)", mainTyped("image/png"), "word/document.xml", "sessions 49 and 50"],
        [
            "main document Override typed application/foo+xml (s50c)",
            mainTyped("application/foo+xml"),
            "word/document.xml",
            "sessions 49 and 50",
        ],
        [
            "styles.xml Override typed application/octet-stream (s49a)",
            stylesXmlTyped(OCTET_STREAM),
            "word/styles.xml",
            "sessions 46 and 49",
        ],
        [
            "styles.xml Override typed APPLICATION/OCTET-STREAM (s49a)",
            stylesXmlTyped("APPLICATION/OCTET-STREAM"),
            "word/styles.xml",
            "sessions 46 and 49",
        ],
        [
            "styles.xml Override typed application/octet-stream; charset=utf-8 (s49a)",
            stylesXmlTyped("application/octet-stream; charset=utf-8"),
            "word/styles.xml",
            "sessions 46 and 49",
        ],
        ["styles.xml Override typed image/png (s49a)", stylesXmlTyped("image/png"), "word/styles.xml", "sessions 46 and 49"],
        [
            "styles.xml Override typed application/foo+xml (s49a)",
            stylesXmlTyped("application/foo+xml"),
            "word/styles.xml",
            "sessions 46 and 49",
        ],
        [
            "styles.bin typed application/octet-stream by a Default (s46e)",
            stylesBin(STYLES_XML, false),
            "word/styles.bin",
            "sessions 46 and 49",
        ],
        [
            "a related settings part typed application/octet-stream (s50d)",
            relatedTyped("settings", OCTET_STREAM),
            "word/settings.xml",
            "sessions 50 and 53",
        ],
        [
            "a related fontTable part typed application/octet-stream (s50d)",
            relatedTyped("fontTable", OCTET_STREAM),
            "word/fontTable.xml",
            "sessions 50 and 53",
        ],
        [
            "a related numbering part typed application/octet-stream (s50d)",
            relatedTyped("numbering", OCTET_STREAM),
            "word/numbering.xml",
            "sessions 50 and 53",
        ],
    ] as const)(
        "rejects %s under word-valid",
        async (_label, parts, part, citation) => {
            await withTempDir(async (dir) => {
                const file = path.join(dir, "specimen.docx");
                writeFileSync(file, await buildDocx(parts));
                const result = await validate(file);
                expect(result.valid).toBe(false);
                const found = result.issues.filter((issue) => issue.code === ID);
                expect(found).toHaveLength(1);
                expect(found[0]).toMatchObject({ severity: "error", path: part });
                expect(found[0]?.message).toContain(`Word 16.115 offers to recover the package (${citation})`);
            });
        },
        60_000,
    );

    it("keeps a named run as blocking as the full run on the settings row (#381)", async () => {
        await withTempDir(async (dir) => {
            const file = path.join(dir, "settings-octet.docx");
            writeFileSync(file, await buildDocx(relatedTyped("settings", OCTET_STREAM)));
            const full = await validate(file);
            const found = full.issues.filter((issue) => issue.code === ID);
            expect(found[0]?.severity).toBe("error");
            const partial = await validate(file, { checks: [ID] });
            expect(partial.issues.filter((issue) => issue.code === ID)).toEqual(found);
            expect(partial.valid).toBe(false);
        });
    }, 60_000);

    // Unobserved reaches keep their word-valid warning, citing the spec.
    it.each([
        [
            "a header the body references",
            (): Record<string, string | Uint8Array> => ({
                ...MINIMAL_PARTS,
                "word/document.xml": replaceOnce(
                    replaceOnce(MINIMAL_PARTS["word/document.xml"]!, `xmlns:w="${W}"`, `xmlns:w="${W}" xmlns:r="${R}"`),
                    "<w:sectPr>",
                    `<w:sectPr><w:headerReference w:type="default" r:id="rIdH"/>`,
                ),
                "word/header9.xml": `${XML_DECL}<w:hdr xmlns:w="${W}"><w:p><w:r><w:t>Header</w:t></w:r></w:p></w:hdr>`,
                [SIDECAR]: replaceOnce(
                    MINIMAL_PARTS[SIDECAR]!,
                    "</Relationships>",
                    `<Relationship Id="rIdH" Type="${R}/header" Target="header9.xml"/></Relationships>`,
                ),
                [CT]: replaceOnce(
                    MINIMAL_PARTS[CT]!,
                    "</Types>",
                    `<Override PartName="/word/header9.xml" ContentType="${OCTET_STREAM}"/></Types>`,
                ),
            }),
            "word/header9.xml",
        ],
        [
            "a reached glossary's own styles part",
            (): Record<string, string | Uint8Array> => ({
                ...MINIMAL_PARTS,
                "word/glossary/document.xml":
                    `${XML_DECL}<w:glossaryDocument xmlns:w="${W}"><w:docParts><w:docPart><w:docPartPr><w:name w:val="Part"/>` +
                    `<w:category><w:name w:val="General"/><w:gallery w:val="placeholder"/></w:category>` +
                    `<w:guid w:val="{6A9E2C40-3F1B-4C55-9D2E-7B1A0C9D4E11}"/></w:docPartPr>` +
                    `<w:docPartBody><w:p><w:r><w:t>Part</w:t></w:r></w:p></w:docPartBody></w:docPart></w:docParts></w:glossaryDocument>`,
                "word/glossary/styles.xml": MINIMAL_PARTS["word/styles.xml"]!,
                "word/glossary/_rels/document.xml.rels":
                    `${XML_DECL}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
                    `<Relationship Id="rId1" Type="${R}/styles" Target="styles.xml"/></Relationships>`,
                [SIDECAR]: replaceOnce(
                    MINIMAL_PARTS[SIDECAR]!,
                    "</Relationships>",
                    `<Relationship Id="rIdG" Type="${R}/glossaryDocument" Target="glossary/document.xml"/></Relationships>`,
                ),
                [CT]: replaceOnce(
                    MINIMAL_PARTS[CT]!,
                    "</Types>",
                    `<Override PartName="/word/glossary/document.xml" ContentType="${WML}.document.glossary+xml"/>` +
                        `<Override PartName="/word/glossary/styles.xml" ContentType="${OCTET_STREAM}"/></Types>`,
                ),
            }),
            "word/glossary/styles.xml",
        ],
    ] as const)(
        "keeps %s typed application/octet-stream a word-valid warning",
        async (_label, make, part) => {
            await withTempDir(async (dir) => {
                const file = path.join(dir, "unobserved.docx");
                writeFileSync(file, await buildDocx(make()));
                const result = await validate(file);
                const found = result.issues.filter((issue) => issue.code === ID);
                expect(found).toHaveLength(1);
                expect(found[0]).toMatchObject({ severity: "warning", path: part });
                expect(found[0]?.message).not.toContain("Word 16.115");
            });
        },
        60_000,
    );

    it("keeps the clean controls valid under word-valid", async () => {
        await withTempDir(async (dir) => {
            const minimal = path.join(dir, "minimal.docx");
            writeFileSync(minimal, await buildDocx({ ...MINIMAL_PARTS }));
            expect((await validate(minimal)).valid).toBe(true);
            const related = path.join(dir, "related.docx");
            writeFileSync(related, await buildDocx(relatedTyped("settings", SETTINGS_CT)));
            expect((await validate(related)).valid).toBe(true);
            const typedMain = path.join(dir, "typed-main.docx");
            writeFileSync(typedMain, await buildDocx(mainTyped(MAIN_TYPE)));
            expect((await validate(typedMain)).valid).toBe(true);
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

// #741: `UNRECOGNIZED_OBSERVED_SEGMENTS` and `RECOVERY_SESSIONS` were two
// independent literals, so a segment added to the first without an entry in
// the second would print "(undefined)" in the recovery sentence. The set is
// derived from the map's keys now; this guard fails if the two ever drift.
describe("every observed segment cites a session (#741)", () => {
    it("gives each observed reach its recovery-session citation", () => {
        expect(UNRECOGNIZED_OBSERVED_SEGMENTS.size).toBeGreaterThan(0);
        for (const segment of UNRECOGNIZED_OBSERVED_SEGMENTS) {
            expect(RECOVERY_SESSIONS.get(segment), segment).toBeDefined();
            expect(RECOVERY_SESSIONS.get(segment), segment).toMatch(/^sessions? \d+/);
        }
    });
});
