/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

/**
 * The `word-part-content-type-generic` check (#432): a part Word reads, typed
 * by a generic XML content type — `application/xml` or `text/xml`, from a
 * Default or an Override (#517).
 *
 * Word 16.115 offers to recover the package for every part probed that
 * resolves only through the generic Default: the main document and styles
 * with their Override removed, and settings, fontTable and webSettings added
 * with a relationship but no Override. Every control with the part's proper
 * Override opens clean (session 8a in
 * docs/plans/evidence/word-probes-2026-10-24.md).
 *
 * The check reports one finding per part, at the part's path, naming the
 * relationship that reaches it and the content type the part needs. It is an
 * error — and Word-blocking — under word-valid for DOCX/DOCM, an error under
 * strict (ECMA-376 Part 1 gives each of these parts its own content type),
 * and a warning under lenient. When strict rejects the same main path
 * through `ct-undeclared-part`, it suppresses the duplicate generic finding.
 */

import { mkdirSync, readFileSync, writeFileSync } from "node:fs";
import path from "node:path";

import { describe, expect, it } from "vitest";

import { withTempDir } from "../src/lib/run-cli";
import type { ValidationIssue } from "../src/lib/types";
import { ValidationContext } from "../src/lib/validation-context";
import { createDefaultCheckRegistry } from "../src/scripts/office/checks";
import { WORD_PART_CONTENT_TYPE_GENERIC } from "../src/scripts/office/checks/word-part-content-type-generic";
import { validate } from "../src/scripts/office/validate";
import { isWordBlockingIssue } from "../src/scripts/office/validators/word-blocking";
import { buildDocx, MINIMAL_PARTS } from "./support/minimal-docx";

import { GENERIC_RENAMED_MAIN_PARTS } from "./support/word-profile-regression-specimens";

const ID = WORD_PART_CONTENT_TYPE_GENERIC;

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const R_STRICT = "http://purl.oclc.org/ooxml/officeDocument/relationships";

type Parts = Record<string, string | Uint8Array>;

const CT = {
    document: "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml",
    styles: "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml",
    settings: "application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml",
    fontTable: "application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml",
    webSettings: "application/vnd.openxmlformats-officedocument.wordprocessingml.webSettings+xml",
    header: "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml",
    numbering: "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml",
} as const;

/** The five parts session 8a observed Word reject when typed by the generic Default. */
const SPECIMENS = [
    { segment: "officeDocument", part: "word/document.xml" },
    { segment: "styles", part: "word/styles.xml" },
    { segment: "settings", part: "word/settings.xml" },
    { segment: "fontTable", part: "word/fontTable.xml" },
    { segment: "webSettings", part: "word/webSettings.xml" },
] as const;

/** The well-formed extra parts the new-part specimens add, with their relationship. */
const EXTRA_PARTS: Readonly<Record<string, { readonly xml: string; readonly type: string }>> = {
    "word/settings.xml": { xml: `<w:settings xmlns:w="${W}"/>`, type: `${R}/settings` },
    "word/fontTable.xml": { xml: `<w:fonts xmlns:w="${W}"/>`, type: `${R}/fontTable` },
    "word/webSettings.xml": { xml: `<w:webSettings xmlns:w="${W}"/>`, type: `${R}/webSettings` },
};

/** Remove the Override that types `partName`. */
const withoutOverride = (parts: Parts, partName: string): Parts => ({
    ...parts,
    "[Content_Types].xml": (parts["[Content_Types].xml"] as string).replace(
        new RegExp(`<Override PartName="/${partName.replaceAll(".", "\\.")}" ContentType="[^"]+"/>`),
        "",
    ),
});

/** Retype `partName` with `contentType`, replacing its Override. */
const withOverride = (parts: Parts, partName: string, contentType: string): Parts => ({
    ...parts,
    "[Content_Types].xml": (withoutOverride(parts, partName)["[Content_Types].xml"] as string).replace(
        "</Types>",
        `<Override PartName="/${partName}" ContentType="${contentType}"/></Types>`,
    ),
});

/** Add a main-document relationship for an extra part. */
const withDocRel = (parts: Parts, type: string, target: string): Parts => ({
    ...parts,
    "word/_rels/document.xml.rels": (parts["word/_rels/document.xml.rels"] as string).replace(
        "</Relationships>",
        `<Relationship Id="rId9" Type="${type}" Target="${target}"/></Relationships>`,
    ),
});

/** The specimen whose `part` resolves only through the generic Default. */
const genericDefault = (part: string): Parts => {
    const parts = withoutOverride({ ...MINIMAL_PARTS }, part);
    const extra = EXTRA_PARTS[part];
    return extra === undefined ? parts : withDocRel({ ...parts, [part]: extra.xml }, extra.type, path.posix.basename(part));
};

/** The same package with the `xml` Default carrying `contentType` instead of application/xml. */
const typedByDefault = (part: string, contentType: string): Parts => ({
    ...genericDefault(part),
    "[Content_Types].xml": (genericDefault(part)["[Content_Types].xml"] as string).replace(
        '<Default Extension="xml" ContentType="application/xml"/>',
        `<Default Extension="xml" ContentType="${contentType}"/>`,
    ),
});

/** The control: the same package with the part's proper Override. */
const properOverride = (part: string): Parts => {
    const contentType =
        part === "word/document.xml" ? CT.document : CT[path.posix.basename(part).replace(/\.[^.]+$/, "") as keyof typeof CT];
    const parts = genericDefault(part);
    return {
        ...parts,
        "[Content_Types].xml": (parts["[Content_Types].xml"] as string).replace(
            "</Types>",
            `<Override PartName="/${part}" ContentType="${contentType}"/></Types>`,
        ),
    };
};

/** The Word main-document content types the officeDocument finding names (#505). */
const MAIN_DOCUMENT_TYPES = [
    "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml",
    "application/vnd.ms-word.document.macroEnabled.main+xml",
    "application/vnd.openxmlformats-officedocument.wordprocessingml.template.main+xml",
    "application/vnd.ms-word.template.macroEnabledTemplate.main+xml",
] as const;

const mainDocumentNeeds = (source: "default" | "override"): string => {
    const list = MAIN_DOCUMENT_TYPES.map((type) => `"${type}"`);
    const types = `${list.slice(0, -1).join(", ")} or ${list[list.length - 1]}`;
    return source === "default"
        ? `it needs its own Override carrying one of the Word main-document content types — ${types}`
        : `its Override should carry one of the Word main-document content types — ${types}`;
};

const expectedFinding = (
    part: string,
    segment: string,
    severity: "error" | "warning",
    options: { readonly source?: "default" | "override"; readonly contentType?: string } = {},
) => {
    const source = options.source ?? "default";
    const contentType = options.contentType ?? "application/xml";
    const by = segment === "officeDocument" ? "the package officeDocument relationship" : `the main document's ${segment} relationship`;
    const extension = part.slice(part.lastIndexOf(".") + 1);
    const resolved =
        source === "default"
            ? `resolves only through the Default for extension "${extension}" to "${contentType}"`
            : `is typed "${contentType}" by an Override`;
    const needed =
        segment === "officeDocument"
            ? mainDocumentNeeds(source)
            : source === "default"
              ? `it needs its own Override "${CT[segment as keyof typeof CT]}"`
              : `its Override should carry "${CT[segment as keyof typeof CT]}"`;
    return {
        code: ID,
        severity,
        path: part,
        message: `${part}, named by ${by}, ${resolved}; ${needed}. Word 16.115 offers to recover the package.`,
    };
};

const findings = (issues: readonly { code?: string }[]) => issues.filter((issue) => issue.code === ID);

const diagnose = async (ctx: ValidationContext): Promise<readonly ValidationIssue[]> => {
    const check = createDefaultCheckRegistry().get(ID);
    expect(check, "the check is registered").toBeDefined();
    return check ? await check.run(ctx) : [];
};

describe("word-part-content-type-generic (#432)", () => {
    it("registers in the content-types group for DOCX/DOCM only", () => {
        const check = createDefaultCheckRegistry().get(ID);
        expect(check).toMatchObject({ id: ID, groups: ["content-types"], formats: ["docx"] });
        // It reads the content-types stream and the relationships parts, so
        // the XML syntax gate must precede it.
        expect(check?.requires ?? []).toContain("xml-syntax");
    });

    it.each(SPECIMENS)("reports the $segment part typed by the generic Default", { timeout: 120_000 }, async ({ segment, part }) => {
        const issues = await diagnose(ValidationContext.fromEntries(genericDefault(part)));
        expect(issues).toEqual([expectedFinding(part, segment, "error")]);
    });

    it.each(SPECIMENS)("accepts the $segment part with its proper Override", async ({ part }) => {
        expect(await diagnose(ValidationContext.fromEntries(properOverride(part)))).toEqual([]);
    });

    it("stays silent for a specific but wrong Override, which is out of scope", async () => {
        // A styles Override carrying the settings content type is a different
        // defect; this check judges only the generic application/xml Default.
        const parts = withOverride({ ...MINIMAL_PARTS }, "word/styles.xml", CT.settings);
        expect(await diagnose(ValidationContext.fromEntries(parts))).toEqual([]);
    });

    it.each([
        ["Transitional", R],
        ["Strict", R_STRICT],
    ] as const)("stays silent for the %s type `…/relationships/custom/styles`, an extra path segment (#519)", async (_label, namespace) => {
        // A type has an Office segment only when it is exactly the namespace
        // plus one segment: `custom/styles` is an unobserved type, not styles,
        // so its target is judged by no relationship segment.
        const parts = withDocRel(
            { ...MINIMAL_PARTS, "word/extra.xml": `<w:styles xmlns:w="${W}"/>` },
            `${namespace}/custom/styles`,
            "extra.xml",
        );
        expect(await diagnose(ValidationContext.fromEntries(parts))).toEqual([]);
    });

    it("stays silent for the unobserved types, until a probe observes them", async () => {
        // Session 8a observed only the main document, styles, settings,
        // fontTable and webSettings; numbering and customXml are not reported.
        const numbering = withDocRel(
            withoutOverride({ ...MINIMAL_PARTS, "word/numbering.xml": `<w:numbering xmlns:w="${W}"/>` }, "word/numbering.xml"),
            `${R}/numbering`,
            "numbering.xml",
        );
        expect(await diagnose(ValidationContext.fromEntries(numbering))).toEqual([]);
    });

    it("reports a part reached through its Strict-conformance relationship type", async () => {
        const parts = withDocRel(
            { ...withoutOverride({ ...MINIMAL_PARTS }, "word/settings.xml"), "word/settings.xml": `<w:settings xmlns:w="${W}"/>` },
            "http://purl.oclc.org/ooxml/officeDocument/relationships/settings",
            "settings.xml",
        );
        expect(await diagnose(ValidationContext.fromEntries(parts))).toEqual([expectedFinding("word/settings.xml", "settings", "error")]);
    });

    it("reports a part whose Default is spelt in another case", async () => {
        // `application/XML` is the same content type, compared ASCII case-insensitively.
        const parts = {
            ...genericDefault("word/styles.xml"),
            "[Content_Types].xml": (genericDefault("word/styles.xml")["[Content_Types].xml"] as string).replace(
                'ContentType="application/xml"',
                'ContentType="application/XML"',
            ),
        };
        expect(await diagnose(ValidationContext.fromEntries(parts))).toEqual([expectedFinding("word/styles.xml", "styles", "error")]);
    });

    it("reports each part once, however many relationships reach it", async () => {
        const parts = withDocRel(
            withDocRel({ ...genericDefault("word/styles.xml") }, `${R}/styles`, "styles.xml"),
            `${R}/settings`,
            "styles.xml",
        );
        expect(await diagnose(ValidationContext.fromEntries(parts))).toEqual([expectedFinding("word/styles.xml", "styles", "error")]);
    });

    it("reports nothing without a content-types stream, or from one that does not parse", async () => {
        const { "[Content_Types].xml": _stream, ...noStream } = genericDefault("word/styles.xml");
        expect(await diagnose(ValidationContext.fromEntries(noStream))).toEqual([]);
        expect(
            await diagnose(ValidationContext.fromEntries({ ...genericDefault("word/styles.xml"), "[Content_Types].xml": "<Types" })),
        ).toEqual([]);
    });

    it("is Word-blocking for DOCX and DOCM", () => {
        expect(isWordBlockingIssue({ code: ID, severity: "error", message: "", path: "word/styles.xml" }, new Set(), new Set())).toBe(true);
    });

    it("resolves through the stream after Markup Compatibility processing, as every content-type judgment does (#459, #489)", async () => {
        // The styles Override sits in one branch of an AlternateContent whose
        // Choice requires an unknown namespace, so a processor takes the
        // Fallback (OPC [M1.19]); ct-markup-compatibility reports the markup.
        const alternate = (choice: string, fallback: string): Parts => {
            const parts = withoutOverride(MINIMAL_PARTS, "word/styles.xml");
            return {
                ...parts,
                "[Content_Types].xml": (parts["[Content_Types].xml"] as string)
                    .replace(
                        /<Types [^>]*>/,
                        (open) =>
                            open.slice(0, -1) +
                            ' xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x="urn:x:unknown">',
                    )
                    .replace(
                        "</Types>",
                        `<mc:AlternateContent><mc:Choice Requires="x">${choice}</mc:Choice><mc:Fallback>${fallback}</mc:Fallback></mc:AlternateContent></Types>`,
                    ),
            };
        };
        const styles = `<Override PartName="/word/styles.xml" ContentType="${CT.styles}"/>`;
        expect(await diagnose(ValidationContext.fromEntries(alternate(styles, "")))).toEqual([
            expectedFinding("word/styles.xml", "styles", "error"),
        ]);
        expect(await diagnose(ValidationContext.fromEntries(alternate("", styles)))).toEqual([]);
    });

    it("names the Default the part's extension actually used, whatever the extension (#504)", async () => {
        // word/styles.sty is covered by the `sty` Default, not the `xml` one;
        // Word receives the resolved content type either way, so the check
        // reports it and names the Default that resolved the part.
        const parts = {
            ...MINIMAL_PARTS,
            "word/styles.sty": `<w:styles xmlns:w="${W}"/>`,
            "word/_rels/document.xml.rels": MINIMAL_PARTS["word/_rels/document.xml.rels"]!.replace(
                'Target="styles.xml"',
                'Target="styles.sty"',
            ),
            "[Content_Types].xml": MINIMAL_PARTS["[Content_Types].xml"]!.replace(
                "</Types>",
                '<Default Extension="sty" ContentType="application/xml"/></Types>',
            ),
        };
        const [issue] = await diagnose(ValidationContext.fromEntries(parts));
        expect(issue).toMatchObject({ code: ID, severity: "error", path: "word/styles.sty" });
        expect(issue?.message).toContain('the Default for extension "sty" to "application/xml"');
    });

    it.each([
        { label: "an Override to application/xml", source: "override", contentType: "application/xml" },
        { label: "an Override to text/xml", source: "override", contentType: "text/xml" },
        { label: "the Default to text/xml", source: "default", contentType: "text/xml" },
    ] as const)("reports styles typed by $label (#517)", async ({ source, contentType }) => {
        const parts =
            source === "override"
                ? withOverride(genericDefault("word/styles.xml"), "word/styles.xml", contentType)
                : typedByDefault("word/styles.xml", contentType);
        expect(await diagnose(ValidationContext.fromEntries(parts))).toEqual([
            expectedFinding("word/styles.xml", "styles", "error", { source, contentType }),
        ]);
    });
});

describe.each(["docx", "docm"] as const)("word-part-content-type-generic in a %s package", { timeout: 120_000 }, (extension) => {
    const specimen = async (dir: string, label: string, parts: Parts): Promise<string> => {
        const file = path.join(dir, `${label}.${extension}`);
        writeFileSync(file, await buildDocx(parts));
        return file;
    };

    it.each(SPECIMENS)("reports the $segment part as an error under word-valid, and warns under lenient", async ({ segment, part }) => {
        await withTempDir(async (dir) => {
            const file = await specimen(dir, "generic-default", genericDefault(part));
            for (const profile of ["word-valid", "lenient"] as const) {
                const result = await validate(file, { profile });
                expect(findings(result.issues), profile).toEqual([
                    expectedFinding(part, segment, profile === "lenient" ? "warning" : "error"),
                ]);
                // Word-valid fails on the error; lenient only warns — except the
                // main document, whose missing Override `ct-undeclared-part`
                // already rejects there.
                expect(result.valid, profile).toBe(profile === "lenient" && part !== "word/document.xml");
                expect(result.metadata?.executedChecks, profile).toContain(ID);
            }
        });
    });

    it.each(SPECIMENS)("fails strict for the $segment part, once", async ({ segment, part }) => {
        await withTempDir(async (dir) => {
            const file = await specimen(dir, "strict", genericDefault(part));
            const result = await validate(file, { profile: "strict" });
            expect(result.valid).toBe(false);
            if (segment === "officeDocument") {
                // Strict already rejects the main document through
                // ct-undeclared-part; the same defect is not reported twice.
                expect(findings(result.issues)).toEqual([]);
                expect(result.issues.some((issue) => issue.code === "ct-undeclared-part" && issue.path === part)).toBe(true);
            } else {
                expect(findings(result.issues)).toEqual([expectedFinding(part, segment, "error")]);
            }
        });
    });

    it.each(SPECIMENS)("passes every profile with the proper Override", async ({ part }) => {
        await withTempDir(async (dir) => {
            const file = await specimen(dir, "control", properOverride(part));
            for (const profile of ["word-valid", "lenient", "strict"] as const) {
                const result = await validate(file, { profile });
                expect(findings(result.issues), profile).toEqual([]);
                expect(result.valid, profile).toBe(true);
            }
        });
    });

    it("does not report a part named only by `…/relationships/custom/styles` through validate() (#519)", async () => {
        await withTempDir(async (dir) => {
            const parts = withDocRel(
                { ...MINIMAL_PARTS, "word/extra.xml": `<w:styles xmlns:w="${W}"/>` },
                `${R}/custom/styles`,
                "extra.xml",
            );
            const file = await specimen(dir, "custom-segment", parts);
            for (const profile of ["word-valid", "lenient"] as const) {
                const result = await validate(file, { profile });
                expect(findings(result.issues), profile).toEqual([]);
                expect(result.valid, profile).toBe(true);
            }
        });
    });

    it("agrees between the full, quick and named runs, packed and directory", { timeout: 240_000 }, async () => {
        await withTempDir(async (dir) => {
            const parts = genericDefault("word/styles.xml");
            const packed = await specimen(dir, "packed", parts);
            const directory = path.join(dir, "unpacked.docx");
            for (const [name, content] of Object.entries(parts)) {
                mkdirSync(path.dirname(path.join(directory, name)), { recursive: true });
                writeFileSync(path.join(directory, name), content as string);
            }
            const errorPairs = (result: Awaited<ReturnType<typeof validate>>) =>
                result.issues.filter((issue) => issue.severity === "error").map((issue) => `${issue.code}@${issue.path ?? ""}`);

            for (const file of [packed, directory]) {
                for (const profile of ["word-valid", "lenient"] as const) {
                    const full = await validate(file, { profile });
                    const severity = profile === "lenient" ? "warning" : "error";
                    for (const selection of [{ quick: true }, { checks: ["content-types"] }, { checks: [ID] }]) {
                        const partial = await validate(file, { profile, ...selection });
                        const tag = `${path.basename(file)} ${profile} ${JSON.stringify(selection)}`;
                        expect(findings(partial.issues), tag).toEqual([expectedFinding("word/styles.xml", "styles", severity)]);
                        // #381: the partial run fails exactly as the full run
                        // does for the errors it shares with it.
                        if (full.valid) expect(partial.valid, tag).toBe(true);
                        expect(errorPairs(full), tag).toEqual(expect.arrayContaining(errorPairs(partial)));
                    }
                }
            }
        });
    });

    it("keeps the packed input bytes unchanged after validation", { timeout: 120_000 }, async () => {
        await withTempDir(async (dir) => {
            const file = await specimen(dir, "unchanged", genericDefault("word/styles.xml"));
            const bytes = readFileSync(file);
            await validate(file, { profile: "word-valid" });
            await validate(file, { profile: "lenient", checks: [ID] });
            expect(readFileSync(file)).toEqual(bytes);
        });
    });

    it("rejects a partial strict run through ModeConflictError", { timeout: 120_000 }, async () => {
        await withTempDir(async (dir) => {
            const file = await specimen(dir, "strict-partial", genericDefault("word/styles.xml"));
            for (const selection of [{ quick: true }, { checks: [ID] }]) {
                await expect(validate(file, { profile: "strict", ...selection }), JSON.stringify(selection)).rejects.toThrow(/strict/i);
            }
        });
    });
});

/**
 * Session 15 (2026-10-24, `s15-evidence/table.md`): Word 16.115 offers to
 * recover the five generic-XML shapes — `text/xml` by Default or Override
 * and `application/xml` by Override — on the parts probed, so word-valid and
 * strict both reject them (#517).
 */
describe("word-part-content-type-generic on the generic XML spellings (#517)", { timeout: 120_000 }, () => {
    const specimen = async (dir: string, label: string, parts: Parts): Promise<string> => {
        const file = path.join(dir, `${label}.docx`);
        writeFileSync(file, await buildDocx(parts));
        return file;
    };

    it.each([
        {
            label: "main part typed by the Default to text/xml",
            part: "word/document.xml",
            segment: "officeDocument",
            source: "default",
            contentType: "text/xml",
        },
        {
            label: "main part Override to application/xml",
            part: "word/document.xml",
            segment: "officeDocument",
            source: "override",
            contentType: "application/xml",
        },
        {
            label: "main part Override to text/xml",
            part: "word/document.xml",
            segment: "officeDocument",
            source: "override",
            contentType: "text/xml",
        },
        {
            label: "styles Override to application/xml",
            part: "word/styles.xml",
            segment: "styles",
            source: "override",
            contentType: "application/xml",
        },
        {
            label: "styles typed by the Default to text/xml",
            part: "word/styles.xml",
            segment: "styles",
            source: "default",
            contentType: "text/xml",
        },
        {
            label: "settings Override to application/xml (inferred, not probed)",
            part: "word/settings.xml",
            segment: "settings",
            source: "override",
            contentType: "application/xml",
        },
    ] as const)("word-valid and strict both fail the $label", { timeout: 120_000 }, async ({ part, segment, source, contentType }) => {
        await withTempDir(async (dir) => {
            const parts = source === "override" ? withOverride(genericDefault(part), part, contentType) : typedByDefault(part, contentType);
            const file = await specimen(dir, "generic-spelling", parts);
            const wanted = expectedFinding(part, segment, "error", { source, contentType });
            const wordValid = await validate(file, { profile: "word-valid" });
            expect(wordValid.issues).toContainEqual(wanted);
            expect(wordValid.valid).toBe(false);
            const strict = await validate(file, { profile: "strict" });
            expect(strict.valid).toBe(false);
            if (segment === "officeDocument" && source === "default") {
                // The declarable `document` root without an Override is
                // ct-undeclared-part's rejection; it is not reported twice.
                expect(findings(strict.issues)).toEqual([]);
                expect(strict.issues.some((issue) => issue.code === "ct-undeclared-part" && issue.path === part)).toBe(true);
            } else {
                expect(strict.issues).toContainEqual(wanted);
            }
        });
    });
});

/**
 * ECMA-376 Part 1 (4th edition), clause 11.3 gives every WordprocessingML
 * part its own content type — not only the five Word was probed on — so
 * strict, which justifies this check by the spec, reports a Part 1 part
 * typed only by the generic Default, while word-valid keeps judging only
 * the probed parts until a probe observes more (#506).
 */
describe("word-part-content-type-generic under strict covers Part 1 parts (#506)", { timeout: 120_000 }, () => {
    const WML = "application/vnd.openxmlformats-officedocument.wordprocessingml";
    const SPEC_PARTS = [
        { segment: "numbering", part: "word/numbering.xml", root: "numbering", contentType: `${WML}.numbering+xml` },
        { segment: "header", part: "word/header1.xml", root: "hdr", contentType: `${WML}.header+xml` },
        { segment: "footer", part: "word/footer1.xml", root: "ftr", contentType: `${WML}.footer+xml` },
    ] as const;

    /** A Part 1 part with its relationship and no Override: it resolves through the xml Default. */
    const specTyped = (part: string, segment: string, root: string): Parts =>
        withDocRel({ ...MINIMAL_PARTS, [part]: `<w:${root} xmlns:w="${W}"/>` }, `${R}/${segment}`, path.posix.basename(part));

    it.each(SPEC_PARTS)(
        "strict reports the $segment part, citing ECMA-376 Part 1",
        { timeout: 120_000 },
        async ({ segment, part, root, contentType }) => {
            await withTempDir(async (dir) => {
                const file = path.join(dir, `generic-${segment}.docx`);
                writeFileSync(file, await buildDocx(specTyped(part, segment, root)));
                const result = await validate(file, { profile: "strict" });
                const [finding] = result.issues.filter((issue) => issue.code === ID);
                expect(finding).toMatchObject({ code: ID, severity: "error", path: part });
                expect(finding?.message).toContain(`the main document's ${segment} relationship`);
                expect(finding?.message).toContain(`"${contentType}"`);
                expect(finding?.message).toMatch(/ECMA-376 Part 1.*11\.3/);
                expect(result.valid).toBe(false);
            });
        },
    );

    it.each(SPEC_PARTS)(
        "word-valid stays silent for the $segment part until a probe observes it",
        { timeout: 120_000 },
        async ({ segment, part, root }) => {
            await withTempDir(async (dir) => {
                const file = path.join(dir, `generic-${segment}.docx`);
                writeFileSync(file, await buildDocx(specTyped(part, segment, root)));
                const result = await validate(file, { profile: "word-valid" });
                expect(result.issues.filter((issue) => issue.code === ID)).toEqual([]);
                expect(result.valid).toBe(true);
            });
        },
    );
});

it("names the Word main-document content types, including the macro-enabled one, for a .docm (#505)", async () => {
    // The check judges the package, not the file name, so the officeDocument
    // finding cannot assert the macro-free type alone: advising it would turn
    // a macro-enabled package into a macro-free document.
    await withTempDir(async (dir) => {
        const file = path.join(dir, "main-generic.docm");
        writeFileSync(file, await buildDocx(genericDefault("word/document.xml")));
        const result = await validate(file, { profile: "word-valid" });
        const [finding] = result.issues.filter((issue) => issue.code === ID);
        expect(finding?.message).toContain("application/vnd.ms-word.document.macroEnabled.main+xml");
        expect(finding?.message).toContain(CT.document);
    });
});

describe("generic renamed main parts (#535)", () => {
    it.each(GENERIC_RENAMED_MAIN_PARTS)(
        "retains strict's error for $part typed by $contentType",
        async ({ part, parts }) => {
            await withTempDir(async (dir) => {
                const file = path.join(dir, "specimen.docx");
                writeFileSync(file, await buildDocx(parts));
                const word = await validate(file, { profile: "word-valid" });
                const strict = await validate(file, { profile: "strict" });
                expect(word.valid).toBe(false);
                expect(strict.valid && !word.valid, "strict success must imply word-valid success").toBe(false);
                expect(strict.issues).toContainEqual(expect.objectContaining({ code: ID, path: part, severity: "error" }));
            });
        },
        60_000,
    );
});

/**
 * Session 16 (2026-10-24, the table living with issue #540): Word 16.115
 * offers to recover the package for a Word-read part whose content type
 * carries a MIME parameter. The generic check judges the media type —
 * everything before the first `;`, ASCII-lowercased — so
 * `application/xml; charset=utf-8` counts as generic (#540).
 */
describe("word-part-content-type-generic judges the media type, parameters stripped (#540)", { timeout: 120_000 }, () => {
    it("reports styles typed `application/xml; charset=utf-8` by an Override, naming the declared value", async () => {
        const parts = withOverride(genericDefault("word/styles.xml"), "word/styles.xml", "application/xml; charset=utf-8");
        expect(await diagnose(ValidationContext.fromEntries(parts))).toEqual([
            expectedFinding("word/styles.xml", "styles", "error", { source: "override", contentType: "application/xml; charset=utf-8" }),
        ]);
    });

    it("reports styles typed `text/xml; charset=utf-8` through the Default", async () => {
        const parts = typedByDefault("word/styles.xml", "text/xml; charset=utf-8");
        expect(await diagnose(ValidationContext.fromEntries(parts))).toEqual([
            expectedFinding("word/styles.xml", "styles", "error", { contentType: "text/xml; charset=utf-8" }),
        ]);
    });

    it("reports the main document typed `application/xml; charset=utf-8` by an Override", async () => {
        const parts = withOverride(genericDefault("word/document.xml"), "word/document.xml", "application/xml; charset=utf-8");
        expect(await diagnose(ValidationContext.fromEntries(parts))).toEqual([
            expectedFinding("word/document.xml", "officeDocument", "error", {
                source: "override",
                contentType: "application/xml; charset=utf-8",
            }),
        ]);
    });

    it("infers the same verdict for settings, an unprobed spelling of the same resolution", async () => {
        // Session 16 probed the parameterized generic type only on styles; the
        // settings spelling follows from the same resolution, as #517's did.
        const parts = withOverride(
            withDocRel(
                { ...withoutOverride({ ...MINIMAL_PARTS, "word/settings.xml": `<w:settings xmlns:w="${W}"/>` }, "word/settings.xml") },
                `${R}/settings`,
                "settings.xml",
            ),
            "word/settings.xml",
            "application/xml; charset=utf-8",
        );
        expect(await diagnose(ValidationContext.fromEntries(parts))).toEqual([
            expectedFinding("word/settings.xml", "settings", "error", {
                source: "override",
                contentType: "application/xml; charset=utf-8",
            }),
        ]);
    });

    it("fails word-valid and strict for the observed specimen: styles Override `application/xml; charset=utf-8`", async () => {
        await withTempDir(async (dir) => {
            const file = path.join(dir, "styles-appxml-charset.docx");
            writeFileSync(
                file,
                await buildDocx(withOverride(genericDefault("word/styles.xml"), "word/styles.xml", "application/xml; charset=utf-8")),
            );
            const wanted = expectedFinding("word/styles.xml", "styles", "error", {
                source: "override",
                contentType: "application/xml; charset=utf-8",
            });
            for (const profile of ["word-valid", "strict"] as const) {
                const result = await validate(file, { profile });
                expect(result.issues, profile).toContainEqual(wanted);
                expect(result.valid, profile).toBe(false);
            }
        });
    }, 60_000);
});

/**
 * The `word-part-content-type-inexact` check (#550, widening #540): a part
 * Word reads, typed by an inexact spelling of its own Word content type —
 * any MIME parameter (session 16 on the main document and styles) or
 * padding whitespace, with the media type still the part's own type
 * (session 17, 2026-10-24; the table lives with issue #550). An upper-case
 * spelling opens clean, so Word folds ASCII case and nothing else. The
 * finding is an error — and Word-blocking — under word-valid, an error
 * under strict, and a warning under lenient; strict drops a Default-typed
 * main only when `ct-undeclared-part` errors at the same path (#535).
 */
describe("word-part-content-type-inexact (#550)", { timeout: 120_000 }, () => {
    const PID = "word-part-content-type-inexact";
    const inexactFindings = (issues: readonly { code?: string }[]) => issues.filter((issue) => issue.code === PID);

    const mainTypesList = MAIN_DOCUMENT_TYPES.map((type) => `"${type}"`)
        .join(", ")
        .replace(/, ([^,]*)$/, " or $1");

    /** The segments Word was observed refusing an inexact type on (session 17); strict-only segments cite Part 1 instead. */
    const RECOVERY_SEGMENTS: ReadonlySet<string> = new Set(["officeDocument", "styles", "settings", "fontTable", "webSettings", "header"]);

    const expectedInexactFinding = (
        part: string,
        segment: string,
        severity: "error" | "warning",
        contentType: string,
        options: { readonly source?: "default" | "override" } = {},
    ) => {
        const source = options.source ?? "override";
        const by = segment === "officeDocument" ? "the package officeDocument relationship" : `the main document's ${segment} relationship`;
        const extension = part.slice(part.lastIndexOf(".") + 1);
        const resolved =
            source === "default"
                ? `resolves only through the Default for extension "${extension}" to "${contentType}"`
                : `is typed "${contentType}" by an Override`;
        const type = CT[segment as keyof typeof CT];
        const clause = `, the content type ECMA-376 Part 1 (4th edition), clause 11.3 assigns the ${segment} part`;
        const needed =
            segment === "officeDocument"
                ? source === "default"
                    ? `it needs its own Override carrying exactly one of the Word main-document content types — ${mainTypesList}`
                    : `its content type must be exactly one of the Word main-document content types — ${mainTypesList}`
                : !RECOVERY_SEGMENTS.has(segment)
                  ? source === "default"
                      ? `it needs its own Override "${type}"${clause}`
                      : `its Override should carry "${type}"${clause}`
                  : source === "default"
                    ? `it needs its own Override "${type}" exactly`
                    : `its Override should carry "${type}" exactly`;
        const verdict = RECOVERY_SEGMENTS.has(segment) ? " Word 16.115 offers to recover the package." : "";
        return {
            code: PID,
            severity,
            path: part,
            message: `${part}, named by ${by}, ${resolved}; ${needed}.${verdict}`,
        };
    };

    const diagnoseInexact = async (ctx: ValidationContext): Promise<readonly ValidationIssue[]> => {
        const check = createDefaultCheckRegistry().get(PID);
        expect(check, "the check is registered").toBeDefined();
        return check ? await check.run(ctx) : [];
    };

    /** A settings, fontTable or webSettings part, related from the main document, its Override carrying `contentType`. */
    const docPartSpecimen = (part: string, contentType: string): Parts => {
        const extra = EXTRA_PARTS[part]!;
        return withOverride(withDocRel({ ...MINIMAL_PARTS, [part]: extra.xml }, extra.type, path.posix.basename(part)), part, contentType);
    };

    /**
     * A header part typed `contentType`, related from the main document and —
     * when `referenced` — used through a `w:headerReference` in `w:sectPr`
     * (session 17's `headerInMain` shape).
     */
    const headerSpecimen = (contentType: string, referenced: boolean): Parts => ({
        ...MINIMAL_PARTS,
        "word/document.xml": referenced
            ? MINIMAL_PARTS["word/document.xml"]!.replace(`xmlns:w="${W}"`, `xmlns:w="${W}" xmlns:r="${R}"`).replace(
                  "<w:sectPr>",
                  `<w:sectPr><w:headerReference w:type="default" r:id="rIdH"/>`,
              )
            : MINIMAL_PARTS["word/document.xml"]!,
        "word/header9.xml": `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:hdr xmlns:w="${W}"><w:p><w:r><w:t>Header</w:t></w:r></w:p></w:hdr>`,
        "word/_rels/document.xml.rels": MINIMAL_PARTS["word/_rels/document.xml.rels"]!.replace(
            "</Relationships>",
            `<Relationship Id="rIdH" Type="${R}/header" Target="header9.xml"/></Relationships>`,
        ),
        "[Content_Types].xml": MINIMAL_PARTS["[Content_Types].xml"]!.replace(
            "</Types>",
            `<Override PartName="/word/header9.xml" ContentType="${contentType}"/></Types>`,
        ),
    });

    it("registers in the content-types group for DOCX/DOCM only", () => {
        const check = createDefaultCheckRegistry().get(PID);
        expect(check).toMatchObject({ id: PID, groups: ["content-types"], formats: ["docx"] });
        // It reads the content-types stream and the relationships parts, so
        // the XML syntax gate must precede it.
        expect(check?.requires ?? []).toContain("xml-syntax");
    });

    it("is Word-blocking for DOCX and DOCM", () => {
        expect(isWordBlockingIssue({ code: PID, severity: "error", message: "", path: "word/styles.xml" }, new Set(), new Set())).toBe(
            true,
        );
    });

    it.each([
        {
            label: "the main document's Override carrying the main type plus `; charset=utf-8`",
            parts: withOverride({ ...MINIMAL_PARTS }, "word/document.xml", `${CT.document}; charset=utf-8`),
            part: "word/document.xml",
            segment: "officeDocument",
            source: "override" as const,
            contentType: `${CT.document}; charset=utf-8`,
        },
        {
            label: "the styles Override carrying the styles type plus `; charset=utf-8`",
            parts: withOverride({ ...MINIMAL_PARTS }, "word/styles.xml", `${CT.styles}; charset=utf-8`),
            part: "word/styles.xml",
            segment: "styles",
            source: "override" as const,
            contentType: `${CT.styles}; charset=utf-8`,
        },
        {
            label: "styles resolved through the Default to the styles type plus `; charset=utf-8`",
            parts: typedByDefault("word/styles.xml", `${CT.styles}; charset=utf-8`),
            part: "word/styles.xml",
            segment: "styles",
            source: "default" as const,
            contentType: `${CT.styles}; charset=utf-8`,
        },
        {
            label: "the settings Override carrying the settings type plus `; charset=utf-8` (s17b)",
            parts: docPartSpecimen("word/settings.xml", `${CT.settings}; charset=utf-8`),
            part: "word/settings.xml",
            segment: "settings",
            source: "override" as const,
            contentType: `${CT.settings}; charset=utf-8`,
        },
        {
            label: "the fontTable Override carrying the fontTable type plus `; charset=utf-8` (s17b)",
            parts: docPartSpecimen("word/fontTable.xml", `${CT.fontTable}; charset=utf-8`),
            part: "word/fontTable.xml",
            segment: "fontTable",
            source: "override" as const,
            contentType: `${CT.fontTable}; charset=utf-8`,
        },
        {
            label: "the webSettings Override carrying the webSettings type plus `; charset=utf-8` (s17b)",
            parts: docPartSpecimen("word/webSettings.xml", `${CT.webSettings}; charset=utf-8`),
            part: "word/webSettings.xml",
            segment: "webSettings",
            source: "override" as const,
            contentType: `${CT.webSettings}; charset=utf-8`,
        },
        {
            label: "a referenced header's Override carrying the header type plus `; charset=utf-8` (s17b)",
            parts: headerSpecimen(`${CT.header}; charset=utf-8`, true),
            part: "word/header9.xml",
            segment: "header",
            source: "override" as const,
            contentType: `${CT.header}; charset=utf-8`,
        },
        {
            label: "the main document's Override carrying the main type plus `;charset=utf-8`, no space (s17c)",
            parts: withOverride({ ...MINIMAL_PARTS }, "word/document.xml", `${CT.document};charset=utf-8`),
            part: "word/document.xml",
            segment: "officeDocument",
            source: "override" as const,
            contentType: `${CT.document};charset=utf-8`,
        },
        {
            label: "the main document's Override carrying the main type plus `; foo=bar` (s17c)",
            parts: withOverride({ ...MINIMAL_PARTS }, "word/document.xml", `${CT.document}; foo=bar`),
            part: "word/document.xml",
            segment: "officeDocument",
            source: "override" as const,
            contentType: `${CT.document}; foo=bar`,
        },
        {
            label: "the styles Override carrying the styles type plus a trailing space (s17c)",
            parts: withOverride({ ...MINIMAL_PARTS }, "word/styles.xml", `${CT.styles} `),
            part: "word/styles.xml",
            segment: "styles",
            source: "override" as const,
            contentType: `${CT.styles} `,
        },
        {
            label: "the styles Override carrying the styles type with a leading space",
            parts: withOverride({ ...MINIMAL_PARTS }, "word/styles.xml", ` ${CT.styles}`),
            part: "word/styles.xml",
            segment: "styles",
            source: "override" as const,
            contentType: ` ${CT.styles}`,
        },
    ] as const)("reports $label", async ({ parts, part, segment, source, contentType }) => {
        expect(await diagnoseInexact(ValidationContext.fromEntries(parts))).toEqual([
            expectedInexactFinding(part, segment, "error", contentType, { source }),
        ]);
    });

    it.each([
        [
            "the main document's Override in upper case (s17c)",
            withOverride({ ...MINIMAL_PARTS }, "word/document.xml", CT.document.toUpperCase()),
        ],
        ["the styles Override in upper case (s17c)", withOverride({ ...MINIMAL_PARTS }, "word/styles.xml", CT.styles.toUpperCase())],
        ["the settings Override with its exact type (s17b)", docPartSpecimen("word/settings.xml", CT.settings)],
        ["the fontTable Override with its exact type (s17b)", docPartSpecimen("word/fontTable.xml", CT.fontTable)],
        ["the webSettings Override with its exact type (s17b)", docPartSpecimen("word/webSettings.xml", CT.webSettings)],
        ["a referenced header's Override with its exact type (s17b)", headerSpecimen(CT.header, true)],
    ] as const)("stays silent for $label", async (_label, parts) => {
        expect(await diagnoseInexact(ValidationContext.fromEntries(parts))).toEqual([]);
    });

    it.each([
        {
            label: "ar-main-override-proper-charset",
            parts: withOverride({ ...MINIMAL_PARTS }, "word/document.xml", `${CT.document}; charset=utf-8`),
            part: "word/document.xml",
            segment: "officeDocument",
            source: "override" as const,
            contentType: `${CT.document}; charset=utf-8`,
        },
        {
            label: "ar-styles-override-proper-charset",
            parts: withOverride({ ...MINIMAL_PARTS }, "word/styles.xml", `${CT.styles}; charset=utf-8`),
            part: "word/styles.xml",
            segment: "styles",
            source: "override" as const,
            contentType: `${CT.styles}; charset=utf-8`,
        },
        {
            label: "ar-settings-override-proper-charset (s17b)",
            parts: docPartSpecimen("word/settings.xml", `${CT.settings}; charset=utf-8`),
            part: "word/settings.xml",
            segment: "settings",
            source: "override" as const,
            contentType: `${CT.settings}; charset=utf-8`,
        },
        {
            label: "ar-fonttable-override-proper-charset (s17b)",
            parts: docPartSpecimen("word/fontTable.xml", `${CT.fontTable}; charset=utf-8`),
            part: "word/fontTable.xml",
            segment: "fontTable",
            source: "override" as const,
            contentType: `${CT.fontTable}; charset=utf-8`,
        },
        {
            label: "ar-websettings-override-proper-charset (s17b)",
            parts: docPartSpecimen("word/webSettings.xml", `${CT.webSettings}; charset=utf-8`),
            part: "word/webSettings.xml",
            segment: "webSettings",
            source: "override" as const,
            contentType: `${CT.webSettings}; charset=utf-8`,
        },
        {
            label: "ar-header-override-proper-charset (s17b)",
            parts: headerSpecimen(`${CT.header}; charset=utf-8`, true),
            part: "word/header9.xml",
            segment: "header",
            source: "override" as const,
            contentType: `${CT.header}; charset=utf-8`,
        },
        {
            label: "ar-main-override-proper-charset-nospace (s17c)",
            parts: withOverride({ ...MINIMAL_PARTS }, "word/document.xml", `${CT.document};charset=utf-8`),
            part: "word/document.xml",
            segment: "officeDocument",
            source: "override" as const,
            contentType: `${CT.document};charset=utf-8`,
        },
        {
            label: "ar-main-override-proper-other-param (s17c)",
            parts: withOverride({ ...MINIMAL_PARTS }, "word/document.xml", `${CT.document}; foo=bar`),
            part: "word/document.xml",
            segment: "officeDocument",
            source: "override" as const,
            contentType: `${CT.document}; foo=bar`,
        },
        {
            label: "ar-styles-override-trailing-space (s17c)",
            parts: withOverride({ ...MINIMAL_PARTS }, "word/styles.xml", `${CT.styles} `),
            part: "word/styles.xml",
            segment: "styles",
            source: "override" as const,
            contentType: `${CT.styles} `,
        },
    ] as const)(
        "word-valid and strict both fail the $label specimen",
        async ({ parts, part, segment, source, contentType }) => {
            await withTempDir(async (dir) => {
                const file = path.join(dir, "specimen.docx");
                writeFileSync(file, await buildDocx(parts));
                const wanted = expectedInexactFinding(part, segment, "error", contentType, { source });
                const wordValid = await validate(file, { profile: "word-valid" });
                expect(wordValid.issues).toContainEqual(wanted);
                expect(wordValid.valid).toBe(false);
                const strict = await validate(file, { profile: "strict" });
                expect(strict.issues).toContainEqual(wanted);
                expect(strict.valid).toBe(false);
            });
        },
        60_000,
    );

    it("warns under lenient without failing the package", async () => {
        await withTempDir(async (dir) => {
            const file = path.join(dir, "lenient.docx");
            writeFileSync(file, await buildDocx(withOverride({ ...MINIMAL_PARTS }, "word/styles.xml", `${CT.styles}; charset=utf-8`)));
            const result = await validate(file, { profile: "lenient" });
            expect(inexactFindings(result.issues)).toEqual([
                expectedInexactFinding("word/styles.xml", "styles", "warning", `${CT.styles}; charset=utf-8`),
            ]);
            expect(result.valid).toBe(true);
            expect(result.metadata?.executedChecks).toContain(PID);
        });
    }, 60_000);

    it("keeps the un-parameterized controls valid under every profile", async () => {
        await withTempDir(async (dir) => {
            const file = path.join(dir, "control.docx");
            writeFileSync(file, await buildDocx(MINIMAL_PARTS));
            for (const profile of ["word-valid", "lenient", "strict"] as const) {
                const result = await validate(file, { profile });
                expect(inexactFindings(result.issues), profile).toEqual([]);
                expect(result.valid, profile).toBe(true);
            }
        });
    }, 60_000);

    it("never reports one part twice: each defect class belongs to one check", async () => {
        // A generic media type with a parameter is the generic check's
        // finding alone; the styles type with a parameter is this check's
        // finding alone.
        const generic = withOverride(genericDefault("word/styles.xml"), "word/styles.xml", "application/xml; charset=utf-8");
        expect(await diagnoseInexact(ValidationContext.fromEntries(generic))).toEqual([]);
        const parameterized = withOverride(genericDefault("word/styles.xml"), "word/styles.xml", `${CT.styles}; charset=utf-8`);
        expect(findings(await diagnose(ValidationContext.fromEntries(parameterized)))).toEqual([]);
    });

    it("keeps a related-but-unused header silent under word-valid and reports it under strict (#516, #550)", async () => {
        await withTempDir(async (dir) => {
            const file = path.join(dir, "unused-header.docx");
            writeFileSync(file, await buildDocx(headerSpecimen(`${CT.header}; charset=utf-8`, false)));
            const wordValid = await validate(file, { profile: "word-valid" });
            expect(inexactFindings(wordValid.issues)).toEqual([]);
            expect(wordValid.valid).toBe(true);
            const strict = await validate(file, { profile: "strict" });
            expect(inexactFindings(strict.issues)).toEqual([
                expectedInexactFinding("word/header9.xml", "header", "error", `${CT.header}; charset=utf-8`),
            ]);
            expect(strict.valid).toBe(false);
        });
    }, 60_000);

    it("reports the unprobed Part 1 types under strict, citing clause 11.3, and stays silent under word-valid (#506)", async () => {
        await withTempDir(async (dir) => {
            const file = path.join(dir, "numbering.docx");
            writeFileSync(
                file,
                await buildDocx({
                    ...MINIMAL_PARTS,
                    "word/numbering.xml": `<w:numbering xmlns:w="${W}"/>`,
                    "word/_rels/document.xml.rels": MINIMAL_PARTS["word/_rels/document.xml.rels"]!.replace(
                        "</Relationships>",
                        `<Relationship Id="rId9" Type="${R}/numbering" Target="numbering.xml"/></Relationships>`,
                    ),
                    "[Content_Types].xml": MINIMAL_PARTS["[Content_Types].xml"]!.replace(
                        "</Types>",
                        `<Override PartName="/word/numbering.xml" ContentType="${CT.numbering}; charset=utf-8"/></Types>`,
                    ),
                }),
            );
            const strict = await validate(file, { profile: "strict" });
            expect(inexactFindings(strict.issues)).toEqual([
                expectedInexactFinding("word/numbering.xml", "numbering", "error", `${CT.numbering}; charset=utf-8`),
            ]);
            expect(strict.valid).toBe(false);
            const wordValid = await validate(file, { profile: "word-valid" });
            expect(inexactFindings(wordValid.issues)).toEqual([]);
        });
    }, 60_000);

    it("drops a Default-typed main under strict only when ct-undeclared-part errors at the same path (#535)", async () => {
        await withTempDir(async (dir) => {
            const declared = path.join(dir, "declared.docx");
            writeFileSync(declared, await buildDocx(typedByDefault("word/document.xml", `${CT.document}; charset=utf-8`)));
            const wordValid = await validate(declared, { profile: "word-valid" });
            expect(wordValid.issues).toContainEqual(
                expectedInexactFinding("word/document.xml", "officeDocument", "error", `${CT.document}; charset=utf-8`, {
                    source: "default",
                }),
            );
            expect(wordValid.valid).toBe(false);
            // The declarable `document` root without an Override is
            // ct-undeclared-part's rejection; it is not reported twice.
            const strict = await validate(declared, { profile: "strict" });
            expect(inexactFindings(strict.issues)).toEqual([]);
            expect(strict.issues.some((issue) => issue.code === "ct-undeclared-part" && issue.path === "word/document.xml")).toBe(true);

            // A renamed main the base pass skips keeps its error (#535).
            const renamed: Parts = { ...MINIMAL_PARTS };
            delete renamed["word/document.xml"];
            delete renamed["word/_rels/document.xml.rels"];
            renamed["word/document.bin"] = MINIMAL_PARTS["word/document.xml"]!;
            renamed["word/_rels/document.bin.rels"] = MINIMAL_PARTS["word/_rels/document.xml.rels"]!.replace(
                'Target="styles.xml"',
                'Target="/word/styles.xml"',
            );
            renamed["_rels/.rels"] = MINIMAL_PARTS["_rels/.rels"]!.replace("word/document.xml", "word/document.bin");
            renamed["[Content_Types].xml"] = (MINIMAL_PARTS["[Content_Types].xml"] as string)
                .replace(`<Override PartName="/word/document.xml" ContentType="${CT.document}"/>`, "")
                .replace("</Types>", `<Default Extension="bin" ContentType="${CT.document}; charset=utf-8"/></Types>`);
            const binFile = path.join(dir, "renamed.docx");
            writeFileSync(binFile, await buildDocx(renamed));
            const renamedStrict = await validate(binFile, { profile: "strict" });
            expect(renamedStrict.issues).toContainEqual(
                expect.objectContaining({ code: PID, path: "word/document.bin", severity: "error" }),
            );
        });
    }, 60_000);
});
