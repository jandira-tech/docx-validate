/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */
import { mkdirSync, readFileSync, writeFileSync } from "node:fs";
import path from "node:path";
import { describe, expect, it } from "vitest";
import { ModeConflictError } from "../src/lib/mode-policy";
import { withTempDir } from "../src/lib/run-cli";
import { ValidationContext } from "../src/lib/validation-context";
import { createDefaultCheckRegistry } from "../src/scripts/office/checks";
import { validate } from "../src/scripts/office/validate";
import { buildDocx, MINIMAL_PARTS } from "./support/minimal-docx";
import { MINIMAL_PPTX_PARTS } from "./support/minimal-pptx";

const ID = "ct-override-partname-word-refused";
const REFUSED = [...["21", "27", "28", "29", "2A", "2a", "80", "FF", "E9"].map((hex) => `/word/a%${hex}b.xml`), "/word/caf%C3%A9.xml"];
const CLEAN = [
    ...["01", "7F", "20", "22", "24", "25", "26", "2B", "2C", "3A", "3B", "3C", "3D", "40", "5B", "7B", "2521"].map(
        (hex) => `/word/a%${hex}b.xml`,
    ),
    "/word/a!b.xml",
];
const withOverrides = (parts: Readonly<Record<string, string>>, names: readonly string[]) => ({
    ...parts,
    "[Content_Types].xml": parts["[Content_Types].xml"]!.replace(
        "</Types>",
        names.map((name) => `<Override PartName="${name}" ContentType="application/xml"/>`).join("") + "</Types>",
    ),
});

/** An AlternateContent Override pair: the Choice requires an unknown namespace, so a processor takes the Fallback (#459). */
const withAlternate = (choice: string, fallback: string) => ({
    ...MINIMAL_PARTS,
    "[Content_Types].xml": MINIMAL_PARTS["[Content_Types].xml"]!.replace(/<Types [^>]*>/, (open) =>
        open.replace(/>$/, ` xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x="urn:x:unknown">`),
    ).replace(
        "</Types>",
        `<mc:AlternateContent><mc:Choice Requires="x"><Override PartName="${choice}" ContentType="application/xml"/></mc:Choice>` +
            `<mc:Fallback><Override PartName="${fallback}" ContentType="application/xml"/></mc:Fallback></mc:AlternateContent></Types>`,
    ),
});

describe("Word-refused Override PartNames: deterministic check", () => {
    it("registers after the XML syntax prerequisite in content-types", () => {
        expect(createDefaultCheckRegistry().get(ID)).toMatchObject({
            groups: ["content-types"],
            formats: ["docx", "pptx", "xlsx"],
            requires: ["xml-syntax"],
        });
    });
    it("reports each refused declaration, including duplicates, while leaving clean escapes alone", async () => {
        const check = createDefaultCheckRegistry().get(ID);
        expect(check).toBeDefined();
        const issues = await check!.run(ValidationContext.fromEntries(withOverrides(MINIMAL_PARTS, [...REFUSED, REFUSED[0]!, ...CLEAN])));
        expect(issues).toHaveLength(REFUSED.length + 1);
        for (const [index, issue] of issues.entries()) {
            expect(issue).toMatchObject({ code: ID, path: "[Content_Types].xml" });
            expect(issue.message).toContain(JSON.stringify([...REFUSED, REFUSED[0]!][index]));
            expect(issue.message).toMatch(/Word 16\.115.*recovery/);
        }
    });
    it("skips absent and malformed content-types streams", async () => {
        const check = createDefaultCheckRegistry().get(ID);
        expect(check).toBeDefined();
        for (const entries of [{}, { "[Content_Types].xml": "<Types" }] as readonly Record<string, string>[])
            expect(await check!.run(ValidationContext.fromEntries(entries))).toEqual([]);
    });
    it("quotes the actual stored stream spelling", async () => {
        const check = createDefaultCheckRegistry().get(ID);
        expect(check).toBeDefined();
        const parts = withOverrides(MINIMAL_PARTS, [REFUSED[0]!]);
        const issues = await check!.run(ValidationContext.fromEntries({ "[content_types].XML": parts["[Content_Types].xml"] }));
        expect(issues).toHaveLength(1);
        expect(issues[0]!.path).toBe("[content_types].XML");
    });
    it("judges the stream after Markup Compatibility processing, as ct-override-partname-invalid does (#459)", async () => {
        const check = createDefaultCheckRegistry().get(ID);
        expect(check).toBeDefined();
        const ignored = await check!.run(ValidationContext.fromEntries(withAlternate(REFUSED[0]!, "/word/fallback.xml")));
        expect(ignored).toEqual([]);
        const both = await check!.run(ValidationContext.fromEntries(withAlternate(REFUSED[0]!, REFUSED[1]!)));
        expect(both.map((issue) => issue.message)).toEqual([expect.stringContaining(JSON.stringify(REFUSED[1]))]);
    });
});

describe.each(["packed", "directory"] as const)("Word-refused Override PartNames: %s integration", { timeout: 120_000 }, (input) => {
    it.each(["docx", "docm", "pptx", "xlsx"] as const)("profile and selection parity for %s without changing bytes", async (extension) => {
        if (extension === "xlsx") expect(createDefaultCheckRegistry().get(ID)).toBeDefined();
        await withTempDir(async (dir) => {
            const parts = withOverrides(extension === "pptx" ? MINIMAL_PPTX_PARTS : MINIMAL_PARTS, [
                ...REFUSED,
                ...CLEAN,
                "/word/café.xml",
            ]);
            const target = path.join(dir, `names.${extension}`);
            const bytes = await buildDocx(parts);
            const before = Buffer.from(bytes);
            if (input === "packed") writeFileSync(target, bytes);
            else
                for (const [name, text] of Object.entries(parts)) {
                    mkdirSync(path.dirname(path.join(target, name)), { recursive: true });
                    writeFileSync(path.join(target, name), text);
                }
            for (const profile of ["word-valid", "strict", "lenient"] as const) {
                const selections = [{ checks: [ID] }, { checks: ["content-types"] }, { quick: true }];
                if (profile === "strict")
                    for (const selection of selections)
                        await expect(validate(target, { profile, ...selection })).rejects.toBeInstanceOf(ModeConflictError);
                if (extension === "xlsx" && profile === "strict") continue;
                const full = await validate(target, { profile, ...(extension === "xlsx" ? { checks: ["content-types"] } : {}) });
                const found = full.issues.filter((i) => i.code === ID);
                // #455: "/word/café.xml" is a valid Part IRI under ECMA-376
                // Part 2 §9.1.1.1, so the grammar check stays silent and this
                // check reports its literal non-ASCII instead — one more
                // finding than REFUSED alone.
                expect(found, `${extension} ${profile}`).toHaveLength(REFUSED.length + 1);
                expect(full.issues.filter((i) => i.code === "ct-override-partname-invalid")).toHaveLength(0);
                const severity = (extension === "docx" || extension === "docm") && profile !== "lenient" ? "error" : "warning";
                expect(found.every((i) => i.severity === severity)).toBe(true);
                if (profile !== "strict")
                    for (const selection of selections) {
                        const partial = await validate(target, { profile, ...selection });
                        expect(partial.issues.filter((i) => i.code === ID)).toEqual(found);
                        const errors = (issues: typeof full.issues) =>
                            issues.filter((i) => i.severity === "error").map((i) => `${i.code}@${i.path ?? ""}`);
                        expect(errors(full.issues)).toEqual(expect.arrayContaining(errors(partial.issues)));
                        if (full.valid) expect(partial.valid).toBe(true);
                    }
            }
            expect(bytes).toEqual(before);
            if (input === "packed") expect(readFileSync(target)).toEqual(before);
            else for (const [name, text] of Object.entries(parts)) expect(readFileSync(path.join(target, name), "utf8")).toBe(text);
        });
    });
    it("reports unreserved and Word-refused escapes independently, once each", async () => {
        await withTempDir(async (dir) => {
            const parts = withOverrides({ ...MINIMAL_PARTS, "word/a%41%21b.xml": "<x/>" }, ["/word/a%41%21b.xml"]);
            const target = path.join(dir, "both.docx");
            if (input === "packed") writeFileSync(target, await buildDocx(parts));
            else
                for (const [name, text] of Object.entries(parts)) {
                    mkdirSync(path.dirname(path.join(target, name)), { recursive: true });
                    writeFileSync(path.join(target, name), text);
                }
            const result = await validate(target);
            for (const code of [ID, "ct-override-partname-invalid", "opc-part-name-word-refused", "opc-part-name-escaped-unreserved"])
                expect(
                    result.issues.filter((i) => i.code === code),
                    code,
                ).toHaveLength(1);
        });
    });
    it("judges only the selected AlternateContent branch in full and partial runs (#459)", async () => {
        await withTempDir(async (dir) => {
            const parts = withAlternate(REFUSED[0]!, REFUSED[1]!);
            const target = path.join(dir, "alternate.docx");
            if (input === "packed") writeFileSync(target, await buildDocx(parts));
            else
                for (const [name, text] of Object.entries(parts)) {
                    mkdirSync(path.dirname(path.join(target, name)), { recursive: true });
                    writeFileSync(path.join(target, name), text);
                }
            for (const selection of [{}, { checks: ["content-types"] }, { checks: [ID] }]) {
                const result = await validate(target, { profile: "word-valid", ...selection });
                const found = result.issues.filter((i) => i.code === ID);
                expect(
                    found.map((i) => i.message),
                    JSON.stringify(selection),
                ).toEqual([expect.stringContaining(JSON.stringify(REFUSED[1]))]);
                expect(found[0]!.severity).toBe("error");
            }
        });
    });
});

describe("Word-refused Override PartNames: message scope (#473)", { timeout: 120_000 }, () => {
    it("reports a literal-ucschar PartName the §9.1.1.1 grammar accepts (#455)", async () => {
        await withTempDir(async (dir) => {
            const parts = withOverrides(MINIMAL_PARTS, ["/word/café.xml"]);
            const check = createDefaultCheckRegistry().get(ID);
            expect(check).toBeDefined();
            const issues = await check!.run(ValidationContext.fromEntries(parts));
            expect(issues).toHaveLength(1);
            expect(issues[0]).toMatchObject({ code: ID, severity: "error", path: "[Content_Types].xml" });
            expect(issues[0]!.message).toContain("literal non-ASCII characters");
            // The grammar accepting the name is what un-suppresses the reason.
            const target = path.join(dir, "part-iri.docx");
            writeFileSync(target, await buildDocx(parts));
            for (const profile of ["word-valid", "strict"] as const) {
                const found = (await validate(target, { profile })).issues.filter((i) => i.code === ID);
                expect(found, profile).toHaveLength(1);
                expect(found[0]!.severity, profile).toBe("error");
                expect(found[0]!.message, profile).toContain("literal non-ASCII characters");
            }
            const pptx = path.join(dir, "part-iri.pptx");
            writeFileSync(pptx, await buildDocx(withOverrides(MINIMAL_PPTX_PARTS, ["/ppt/café.xml"])));
            const pptxFinding = (await validate(pptx, { profile: "strict" })).issues.filter((i) => i.code === ID);
            expect(pptxFinding).toHaveLength(1);
            expect(pptxFinding[0]!.severity).toBe("warning");
        });
    });

    it("scopes the recovery claim to a DOCX package for a PPTX finding", async () => {
        await withTempDir(async (dir) => {
            const target = path.join(dir, "scope.pptx");
            writeFileSync(target, await buildDocx(withOverrides(MINIMAL_PPTX_PARTS, ["/ppt/a%21b.xml"])));
            const found = (await validate(target, { profile: "strict" })).issues.filter((i) => i.code === ID);
            expect(found).toHaveLength(1);
            expect(found[0]!.message).toContain("Word 16.115 offers recovery for this spelling in a DOCX package");
            expect(found[0]!.message).not.toMatch(/offers recovery\.$/);
        });
    });

    it("keeps naming the Word 16.115 recovery observation for a DOCX finding", async () => {
        await withTempDir(async (dir) => {
            const target = path.join(dir, "scope.docx");
            writeFileSync(target, await buildDocx(withOverrides(MINIMAL_PARTS, ["/word/a%21b.xml"])));
            const found = (await validate(target, { profile: "strict" })).issues.filter((i) => i.code === ID);
            expect(found).toHaveLength(1);
            expect(found[0]!.message).toMatch(/Word 16\.115 offers recovery/);
        });
    });
});
