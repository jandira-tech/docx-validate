/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

import { promises as fs } from "node:fs";
import path from "node:path";
import { describe, expect, it } from "vitest";
import { withTempDir } from "../src/lib/run-cli";
import { ValidationContext } from "../src/lib/validation-context";
import { buildDocx, MINIMAL_PARTS } from "./support/minimal-docx";
import { document, rels, relation, R, RS } from "./support/word-rule-test-xml";
import { validate } from "../src/scripts/office/validate";
import {
    wordDuplicateThemeRelationshipCheck,
    wordDuplicateFontTableRelationshipCheck,
    wordDuplicateWebSettingsRelationshipCheck,
    wordDuplicateSettingsRelationshipCheck,
    wordDuplicateStylesRelationshipCheck,
    wordDuplicateNumberingRelationshipCheck,
} from "../src/scripts/office/checks/word-duplicate-relationships";

describe("direct relationship duplicate counts (#434)", () => {
    it.each([
        [wordDuplicateThemeRelationshipCheck, "theme"],
        [wordDuplicateFontTableRelationshipCheck, "fontTable"],
        [wordDuplicateWebSettingsRelationshipCheck, "webSettings"],
        [wordDuplicateSettingsRelationshipCheck, "settings"],
        [wordDuplicateStylesRelationshipCheck, "styles"],
        [wordDuplicateNumberingRelationshipCheck, "numbering"],
    ] as const)("ignores a nested second %s / %s relationship", async (check, type) => {
        const ctx = ValidationContext.fromEntries({
            "word/_rels/document.xml.rels": rels(relation("direct", type) + `<x:Wrap xmlns:x="urn:x">${relation("nested", type)}</x:Wrap>`),
        });
        expect(await check.run(ctx)).toEqual([]);
    });
});

describe("word-duplicate-theme-relationship", () => {
    it("reports the Word-rejected shape with unchanged issue details", async () => {
        const ctx = ValidationContext.fromEntries({
            "word/_rels/document.xml.rels": rels(relation("rId1", "theme") + relation("rId2", "theme")),
        });
        expect(await wordDuplicateThemeRelationshipCheck.run(ctx)).toEqual([
            {
                severity: "error",
                code: "word-duplicate-theme-relationship",
                path: "word/_rels/document.xml.rels",
                message:
                    "word/_rels/document.xml.rels contains 2 theme relationships (rId1, rId2). Microsoft Word offers to recover the file. Keep one theme relationship.",
            },
        ]);
    });
    it("still reads the relationships through an mc:Ignorable naming the package-relationships namespace (#493)", async () => {
        const MC = "http://schemas.openxmlformats.org/markup-compatibility/2006";
        const PR = "http://schemas.openxmlformats.org/package/2006/relationships";
        // The check reads the stream after Markup Compatibility processing;
        // that processor applies the relationships vocabulary, so an
        // Ignorable naming the stream's own namespace hides nothing (#493).
        const ctx = ValidationContext.fromEntries({
            "word/_rels/document.xml.rels":
                `<Relationships xmlns="${PR}" xmlns:mc="${MC}" xmlns:pr="${PR}" mc:Ignorable="pr">` +
                relation("rId1", "theme") +
                relation("rId2", "theme") +
                `</Relationships>`,
        });
        expect(await wordDuplicateThemeRelationshipCheck.run(ctx)).toEqual([
            {
                severity: "error",
                code: "word-duplicate-theme-relationship",
                path: "word/_rels/document.xml.rels",
                message:
                    "word/_rels/document.xml.rels contains 2 theme relationships (rId1, rId2). Microsoft Word offers to recover the file. Keep one theme relationship.",
            },
        ]);
    });
    it("accepts the repaired control", async () => {
        const ctx = ValidationContext.fromEntries({
            "word/_rels/document.xml.rels": rels(relation("rId1", "theme") + relation("other", "image") + `<Relationship/>`),
            "word/document.xml": document(""),
        });
        expect(await wordDuplicateThemeRelationshipCheck.run(ctx)).toEqual([]);
    });
    it("skips absent, irrelevant and unparseable parts", async () => {
        expect(await wordDuplicateThemeRelationshipCheck.run(ValidationContext.fromEntries({}))).toEqual([]);
        expect(
            await wordDuplicateThemeRelationshipCheck.run(
                ValidationContext.fromEntries({ "word/_rels/document.xml.rels": "<broken", "customXml/item.xml": document("") }),
            ),
        ).toEqual([]);
    });
});
describe("word-duplicate-font-table-relationship", () => {
    it("reports the Word-rejected shape with unchanged issue details", async () => {
        const ctx = ValidationContext.fromEntries({
            "word/_rels/document.xml.rels": rels(relation("rId1", "fontTable") + relation("rId2", "fontTable")),
        });
        expect(await wordDuplicateFontTableRelationshipCheck.run(ctx)).toEqual([
            {
                severity: "error",
                code: "word-duplicate-font-table-relationship",
                path: "word/_rels/document.xml.rels",
                message:
                    "word/_rels/document.xml.rels contains 2 font table relationships (rId1, rId2). Microsoft Word offers to recover the file. Keep one font table relationship.",
            },
        ]);
    });
    it("accepts the repaired control", async () => {
        const ctx = ValidationContext.fromEntries({
            "word/_rels/document.xml.rels": rels(relation("rId1", "fontTable") + relation("other", "image") + `<Relationship/>`),
            "word/document.xml": document(""),
        });
        expect(await wordDuplicateFontTableRelationshipCheck.run(ctx)).toEqual([]);
    });
    it("skips absent, irrelevant and unparseable parts", async () => {
        expect(await wordDuplicateFontTableRelationshipCheck.run(ValidationContext.fromEntries({}))).toEqual([]);
        expect(
            await wordDuplicateFontTableRelationshipCheck.run(
                ValidationContext.fromEntries({ "word/_rels/document.xml.rels": "<broken", "customXml/item.xml": document("") }),
            ),
        ).toEqual([]);
    });
});
describe("word-duplicate-web-settings-relationship", () => {
    it("reports the Word-rejected shape with unchanged issue details", async () => {
        const ctx = ValidationContext.fromEntries({
            "word/_rels/document.xml.rels": rels(relation("rId1", "webSettings") + relation("rId2", "webSettings")),
        });
        expect(await wordDuplicateWebSettingsRelationshipCheck.run(ctx)).toEqual([
            {
                severity: "error",
                code: "word-duplicate-web-settings-relationship",
                path: "word/_rels/document.xml.rels",
                message:
                    "word/_rels/document.xml.rels contains 2 web settings relationships (rId1, rId2). Microsoft Word offers to recover the file. Keep one web settings relationship.",
            },
        ]);
    });
    it("accepts the repaired control", async () => {
        const ctx = ValidationContext.fromEntries({
            "word/_rels/document.xml.rels": rels(relation("rId1", "webSettings") + relation("other", "image") + `<Relationship/>`),
            "word/document.xml": document(""),
        });
        expect(await wordDuplicateWebSettingsRelationshipCheck.run(ctx)).toEqual([]);
    });
    it("skips absent, irrelevant and unparseable parts", async () => {
        expect(await wordDuplicateWebSettingsRelationshipCheck.run(ValidationContext.fromEntries({}))).toEqual([]);
        expect(
            await wordDuplicateWebSettingsRelationshipCheck.run(
                ValidationContext.fromEntries({ "word/_rels/document.xml.rels": "<broken", "customXml/item.xml": document("") }),
            ),
        ).toEqual([]);
    });
});
describe("word-duplicate-settings-relationship", () => {
    it("reports the Word-rejected shape with unchanged issue details", async () => {
        const ctx = ValidationContext.fromEntries({
            "word/_rels/document.xml.rels": rels(relation("rId1", "settings") + relation("rId2", "settings")),
        });
        expect(await wordDuplicateSettingsRelationshipCheck.run(ctx)).toEqual([
            {
                severity: "error",
                code: "word-duplicate-settings-relationship",
                path: "word/_rels/document.xml.rels",
                message:
                    "word/_rels/document.xml.rels contains 2 settings relationships (rId1, rId2). Microsoft Word offers to recover the file. Keep one settings relationship.",
            },
        ]);
    });
    it("accepts the repaired control", async () => {
        const ctx = ValidationContext.fromEntries({
            "word/_rels/document.xml.rels": rels(relation("rId1", "settings") + relation("other", "image") + `<Relationship/>`),
            "word/document.xml": document(""),
        });
        expect(await wordDuplicateSettingsRelationshipCheck.run(ctx)).toEqual([]);
    });
    it("skips absent, irrelevant and unparseable parts", async () => {
        expect(await wordDuplicateSettingsRelationshipCheck.run(ValidationContext.fromEntries({}))).toEqual([]);
        expect(
            await wordDuplicateSettingsRelationshipCheck.run(
                ValidationContext.fromEntries({ "word/_rels/document.xml.rels": "<broken", "customXml/item.xml": document("") }),
            ),
        ).toEqual([]);
    });
});

const relationOfType = (id: string, type: string, target = "part.xml"): string =>
    `<Relationship Id="${id}" Type="${type}" Target="${target}"/>`;

describe("duplicate checks count exact Types (session 47, #708)", () => {
    const EXACT_TYPE_CHECKS = [
        ["theme", wordDuplicateThemeRelationshipCheck],
        ["fontTable", wordDuplicateFontTableRelationshipCheck],
        ["webSettings", wordDuplicateWebSettingsRelationshipCheck],
        ["settings", wordDuplicateSettingsRelationshipCheck],
    ] as const;

    // A type counts only as the namespace plus the one segment — Word reads
    // Types exactly (session 27) — so session 47 opened a suffix-only Type
    // clean beside a real relationship of the segment, as #694's styles and
    // numbering checks already read it.
    it.each(EXACT_TYPE_CHECKS)("draws nothing from one %s relationship beside a `custom/%s` type", async (segment, check) => {
        const ctx = ValidationContext.fromEntries({
            "word/_rels/document.xml.rels": rels(
                relation("rId2", segment, "", `${segment}.xml`) + relationOfType("rId3", `${R}/custom/${segment}`, `${segment}2.xml`),
            ),
        });
        expect(await check.run(ctx)).toEqual([]);
    });

    it.each(EXACT_TYPE_CHECKS)("draws nothing from one %s relationship beside a foreign %s-suffixed type", async (segment, check) => {
        const ctx = ValidationContext.fromEntries({
            "word/_rels/document.xml.rels": rels(
                relation("rId2", segment, "", `${segment}.xml`) +
                    relationOfType("rId3", `http://example.com/${segment}`, `${segment}2.xml`),
            ),
        });
        expect(await check.run(ctx)).toEqual([]);
    });

    it.each(EXACT_TYPE_CHECKS)("counts the ISO Strict %s type URI", async (segment, check) => {
        const strict = `${RS}/${segment}`;
        const ctx = ValidationContext.fromEntries({
            "word/_rels/document.xml.rels": rels(
                relationOfType("rId2", strict, `${segment}.xml`) + relationOfType("rId3", strict, `${segment}2.xml`),
            ),
        });
        const issues = await check.run(ctx);
        expect(issues).toHaveLength(1);
        expect(issues[0]?.code).toBe(check.id);
    });
});

describe("OPC-equivalent duplicate relationship checks (#398)", () => {
    it.each([
        ["theme", wordDuplicateThemeRelationshipCheck],
        ["fontTable", wordDuplicateFontTableRelationshipCheck],
        ["webSettings", wordDuplicateWebSettingsRelationshipCheck],
        ["settings", wordDuplicateSettingsRelationshipCheck],
        ["styles", wordDuplicateStylesRelationshipCheck],
        ["numbering", wordDuplicateNumberingRelationshipCheck],
    ] as const)("checks %s relationships under every equivalent spelling", async (type, check) => {
        await withTempDir(async (dir) => {
            const run = async (extension: string) => {
                const file = path.join(dir, `${extension}.docx`);
                await fs.writeFile(
                    file,
                    await buildDocx({ [`word/_rels/document.xml.${extension}`]: rels(relation("rId1", type) + relation("rId2", type)) }),
                );
                return check.run(await ValidationContext.fromArchive(await fs.readFile(file)));
            };
            const control = await run("rels");
            expect(control).toHaveLength(1);
            expect(control[0].code).toBe(check.id);
            for (const extension of ["RELS", "Rels", "%72els", "%52ELS"]) {
                expect(await run(extension)).toEqual(
                    control.map((issue) => ({
                        ...issue,
                        path: `word/_rels/document.xml.${extension}`,
                        message: issue.message.replace("document.xml.rels", `document.xml.${extension}`),
                    })),
                );
            }
        });
    });
});

// Session 47's specimens (#708): the minimal package plus one real
// relationship of the segment and one whose Type only ends in the segment,
// each to its own part typed with the segment's content type — the session
// 47 builder's `suffixSpecimen` shape (`docs/plans/evidence/word-probes-2026-10-24.md`).
const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const CONTENT_TYPES = "[Content_Types].xml";
const SIDECAR = "word/_rels/document.xml.rels";
const replaceOnce = (text: string, from: string, to: string): string => {
    if (text.split(from).length !== 2) throw new Error(`expected exactly one ${from}`);
    return text.replace(from, to);
};
const relationship = (id: string, type: string, target: string): string => `<Relationship Id="${id}" Type="${type}" Target="${target}"/>`;
const withOverrides = (...overrides: [string, string][]): string =>
    replaceOnce(
        MINIMAL_PARTS[CONTENT_TYPES]!,
        "</Types>",
        `${overrides.map(([part, type]) => `<Override PartName="/${part}" ContentType="${type}"/>`).join("")}</Types>`,
    );
const SINGLETONS: Record<string, { root: string; ct: string }> = {
    settings: { root: `<w:settings xmlns:w="${W}"/>`, ct: "application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml" },
    fontTable: { root: `<w:fonts xmlns:w="${W}"/>`, ct: "application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml" },
    webSettings: {
        root: `<w:webSettings xmlns:w="${W}"/>`,
        ct: "application/vnd.openxmlformats-officedocument.wordprocessingml.webSettings+xml",
    },
    theme: {
        root: `<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"/>`,
        ct: "application/vnd.openxmlformats-officedocument.theme+xml",
    },
};
const suffixSpecimen = async (segment: string, suffixType: string): Promise<Buffer> => {
    const parts: Record<string, string> = { ...MINIMAL_PARTS };
    parts[`word/${segment}.xml`] = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n${SINGLETONS[segment]!.root}`;
    parts[`word/${segment}2.xml`] = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n${SINGLETONS[segment]!.root}`;
    parts[SIDECAR] = replaceOnce(
        MINIMAL_PARTS[SIDECAR]!,
        "</Relationships>",
        `${relationship("rId2", `${R}/${segment}`, `${segment}.xml`)}${relationship("rId3", suffixType, `${segment}2.xml`)}</Relationships>`,
    );
    parts[CONTENT_TYPES] = withOverrides(
        [`word/${segment}.xml`, SINGLETONS[segment]!.ct],
        [`word/${segment}2.xml`, SINGLETONS[segment]!.ct],
    );
    return buildDocx(parts);
};

describe("session 47 specimens: a suffix-only Type beside a real relationship (#708)", { timeout: 30000 }, () => {
    it.each([
        ["settings", `${R}/custom/settings`, "settings-custom-settings-own"],
        ["settings", "http://example.com/settings", "settings-foreign-settings-own"],
        ["webSettings", `${R}/custom/webSettings`, "websettings-custom-websettings-own"],
        ["fontTable", `${R}/custom/fontTable`, "fonttable-custom-fonttable-own"],
        // Unobserved in Word, but the same rule.
        ["theme", `${R}/custom/theme`, "theme-custom-theme-own"],
    ] as const)("%s draws no duplicate finding", async (segment, suffixType, name) => {
        await withTempDir(async (dir) => {
            const file = path.join(dir, `${name}.docx`);
            await fs.writeFile(file, await suffixSpecimen(segment, suffixType));
            const wordValid = await validate(file, { profile: "word-valid" });
            expect(wordValid.issues.some((issue) => issue.code?.startsWith("word-duplicate-") === true)).toBe(false);
            expect(wordValid.valid).toBe(true);
            const strict = await validate(file, { profile: "strict" });
            expect(strict.issues.some((issue) => issue.code?.startsWith("word-duplicate-") === true)).toBe(false);
        });
    });
});
