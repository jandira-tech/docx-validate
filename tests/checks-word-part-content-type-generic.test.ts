/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

/**
 * The `word-part-content-type-generic` check (#432): a part Word reads, typed
 * only by `<Default Extension="xml" ContentType="application/xml"/>`.
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
 * and a warning under lenient. Strict already rejects the main document
 * through `ct-undeclared-part`, so a full strict run does not report the same
 * defect twice.
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

const ID = WORD_PART_CONTENT_TYPE_GENERIC;

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

type Parts = Record<string, string | Uint8Array>;

const CT = {
    document: "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml",
    styles: "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml",
    settings: "application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml",
    fontTable: "application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml",
    webSettings: "application/vnd.openxmlformats-officedocument.wordprocessingml.webSettings+xml",
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

const expectedFinding = (part: string, segment: string, severity: "error" | "warning") => ({
    code: ID,
    severity,
    path: part,
    message:
        `${part}, named by ${segment === "officeDocument" ? "the package officeDocument relationship" : `the main document's ${segment} relationship`}, ` +
        `resolves only through the Default for extension "xml" to "application/xml"; it needs its own Override ` +
        `"${segment === "officeDocument" ? CT.document : CT[segment as keyof typeof CT]}". Word 16.115 offers to recover the package.`,
});

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
