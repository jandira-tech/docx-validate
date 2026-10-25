/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

import { mkdir, writeFile } from "node:fs/promises";
import path from "node:path";
import { describe, expect, it } from "vitest";

import { withTempDir } from "../src/lib/run-cli";
import type { ValidationIssue } from "../src/lib/types";
import { validate } from "../src/scripts/office/validate";
import { buildDocx, MINIMAL_PARTS } from "./support/minimal-docx";
import {
    STORY_FOOTER_XML,
    STORY_HEADER_XML,
    glossaryStoryReferenceParts,
    storyContentType,
    storyReferenceParts,
} from "./support/story-content-type-specimens";

const R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const MAIN = "word/document.xml";
const RELS = "word/_rels/document.xml.rels";
const PACKAGE_RELS = "_rels/.rels";
const CT = "[Content_Types].xml";
const HEADER = `<Relationship Id="rIdX" Type="${R}/header" Target="header9.xml"/>`;
const FOOTER = `<Relationship Id="rIdX" Type="${R}/footer" Target="footer9.xml"/>`;
const DUP_Z =
    `<Relationship Id="rIdZ" Type="${R}/hyperlink" Target="https://a.example/" TargetMode="External"/>` +
    `<Relationship Id="rIdZ" Type="${R}/hyperlink" Target="https://b.example/" TargetMode="External"/>`;
const clean = (): Record<string, string> =>
    storyReferenceParts("headerReference", "header", "header9.xml", STORY_HEADER_XML, storyContentType("header"));

// Copy Session 29's builder, including its correctly typed, unreferenced footer.
// Evidence: docs/plans/evidence/word-probes-2026-10-24.md, s29a/s29b.
const withFooterPart = (relationships: string): Record<string, string> => {
    const parts = clean();
    parts["word/footer9.xml"] = STORY_FOOTER_XML;
    parts[CT] = parts[CT]!.replace(
        "</Types>",
        `<Override PartName="/word/footer9.xml" ContentType="${storyContentType("footer")}"/></Types>`,
    );
    parts[RELS] = parts[RELS]!.replace(HEADER, relationships);
    return parts;
};
const packageDuplicate = (identical = false): Record<string, string> => {
    const parts = clean();
    parts[PACKAGE_RELS] = parts[PACKAGE_RELS]!.replace(
        "</Relationships>",
        `${identical ? DUP_Z.replace("https://b.example/", "https://a.example/") : DUP_Z}</Relationships>`,
    );
    return parts;
};
const relocated = (strictType: boolean, orphanDuplicate = false): Record<string, string> => {
    const parts = { ...MINIMAL_PARTS };
    parts["word/main.xml"] = parts[MAIN]!;
    parts["word/_rels/main.xml.rels"] = parts[RELS]!;
    delete parts[MAIN];
    delete parts[RELS];
    parts[CT] = parts[CT]!.replace("/word/document.xml", "/word/main.xml");
    parts[PACKAGE_RELS] = parts[PACKAGE_RELS]!.replace("word/document.xml", "word/main.xml");
    if (strictType) {
        parts[PACKAGE_RELS] = parts[PACKAGE_RELS]!.replace(
            `${R}/officeDocument`,
            "http://purl.oclc.org/ooxml/officeDocument/relationships/officeDocument",
        );
    }
    const sidecar = orphanDuplicate ? RELS : "word/_rels/main.xml.rels";
    if (orphanDuplicate) {
        parts[MAIN] = MINIMAL_PARTS[MAIN]!;
        parts[RELS] = MINIMAL_PARTS[RELS]!;
    }
    parts[sidecar] = parts[sidecar]!.replace("</Relationships>", `${DUP_Z}</Relationships>`);
    return parts;
};
// Session 30's reached-glossary shape (#617): session 26's clean glossary
// footer with the rIdZ pair in the glossary's own sidecar. `reached: false`
// drops the main document's glossaryDocument relationship, so the glossary
// stays typed by its Override but no story root reads it. Evidence:
// docs/plans/evidence/word-probes-2026-10-24.md, s30a.
const glossaryDuplicate = (reached = true): Record<string, string> => {
    const parts = glossaryStoryReferenceParts("footerReference", "footer", "footer9.xml", STORY_FOOTER_XML, storyContentType("footer"));
    parts["word/glossary/_rels/document.xml.rels"] = parts["word/glossary/_rels/document.xml.rels"]!.replace(
        "</Relationships>",
        `${DUP_Z}</Relationships>`,
    );
    if (reached) return parts;
    parts[RELS] = parts[RELS]!.replace(`<Relationship Id="rIdG" Type="${R}/glossaryDocument" Target="glossary/document.xml"/>`, "");
    return parts;
};
// Session 31's builder (#618): `toBin` renames `<dir>/document.xml` to
// `document.bin`, with its sidecar, Override and the relationship naming it,
// so the story root exists under a name the XML-only walk never visits, while
// Word still reads it and its sidecar. Evidence:
// docs/plans/evidence/word-probes-2026-10-24.md, s31a.
const replaceOnce = (text: string, from: string, to: string): string => {
    if (!text.includes(from)) throw new Error(`missing ${from}`);
    return text.replace(from, to);
};
const withDup = (parts: Record<string, string>, rels: string): Record<string, string> => ({
    ...parts,
    [rels]: replaceOnce(parts[rels]!, "</Relationships>", `${DUP_Z}</Relationships>`),
});
const toBin = (parts: Record<string, string>, dir: string, owner: string, target: string): Record<string, string> => {
    const out: Record<string, string> = { ...parts };
    out[`${dir}/document.bin`] = out[`${dir}/document.xml`]!;
    out[`${dir}/_rels/document.bin.rels`] = out[`${dir}/_rels/document.xml.rels`]!;
    delete out[`${dir}/document.xml`];
    delete out[`${dir}/_rels/document.xml.rels`];
    out[CT] = replaceOnce(out[CT]!, `PartName="/${dir}/document.xml"`, `PartName="/${dir}/document.bin"`);
    out[owner] = replaceOnce(out[owner]!, `Target="${target}"`, `Target="${target.replace(/xml$/, "bin")}"`);
    return out;
};
const binMain = (): Record<string, string> => toBin({ ...MINIMAL_PARTS }, "word", PACKAGE_RELS, "word/document.xml");
const binGlossary = (): Record<string, string> =>
    toBin(
        glossaryStoryReferenceParts("footerReference", "footer", "footer9.xml", STORY_FOOTER_XML, storyContentType("footer")),
        "word/glossary",
        RELS,
        "glossary/document.xml",
    );

interface Specimen {
    name: string;
    parts: () => Record<string, string>;
    duplicatePath?: string;
    blocking: boolean;
    /** Exact legacy lenient error codes, recorded before the fix. */
    lenientErrors?: readonly string[];
}
const SCHEMA_ERRORS = ["xsd-error", "xsd-summary"];
const SPECIMENS: readonly Specimen[] = [
    { name: "control-minimal", parts: () => ({ ...MINIMAL_PARTS }), blocking: false, lenientErrors: [] },
    {
        name: "headerref-dup-id-footer-first-header-last",
        parts: () => withFooterPart(FOOTER + HEADER),
        duplicatePath: RELS,
        blocking: true,
        lenientErrors: ["rels-id-duplicate", ...SCHEMA_ERRORS],
    },
    {
        name: "headerref-dup-id-header-first-footer-last",
        parts: () => withFooterPart(HEADER + FOOTER),
        duplicatePath: RELS,
        blocking: true,
        lenientErrors: ["rels-id-duplicate", "rels-id-mismatch", ...SCHEMA_ERRORS],
    },
    {
        name: "headerref-dup-id-header-twice",
        parts: () => withFooterPart(HEADER + HEADER),
        duplicatePath: RELS,
        blocking: true,
        lenientErrors: ["rels-unreferenced", "rels-id-duplicate", ...SCHEMA_ERRORS],
    },
    {
        name: "dup-id-unreferenced-hyperlinks",
        parts: () => withFooterPart(HEADER + DUP_Z),
        duplicatePath: RELS,
        blocking: true,
        lenientErrors: ["rels-unreferenced", "rels-id-duplicate", ...SCHEMA_ERRORS],
    },
    {
        name: "dup-id-package-rels-unreferenced-hyperlinks",
        parts: () => packageDuplicate(),
        duplicatePath: PACKAGE_RELS,
        blocking: true,
        lenientErrors: SCHEMA_ERRORS,
    },
    {
        name: "dup-id-package-rels-identical-hyperlinks",
        parts: () => packageDuplicate(true),
        duplicatePath: PACKAGE_RELS,
        blocking: true,
        lenientErrors: SCHEMA_ERRORS,
    },
    {
        name: "dup-id-header-part-rels-unreferenced-hyperlinks",
        parts: () => ({
            ...clean(),
            "word/_rels/header9.xml.rels": `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${DUP_Z}</Relationships>`,
        }),
        duplicatePath: "word/_rels/header9.xml.rels",
        blocking: false,
        lenientErrors: ["rels-id-duplicate", ...SCHEMA_ERRORS],
    },
    {
        name: "headerref-header-rel-with-footer-part",
        parts: () => withFooterPart(HEADER),
        blocking: false,
        lenientErrors: ["rels-unreferenced"],
    },
    {
        name: "dup-id-relocated-main-transitional",
        parts: () => relocated(false),
        duplicatePath: "word/_rels/main.xml.rels",
        blocking: true,
    },
    { name: "dup-id-relocated-main-strict", parts: () => relocated(true), duplicatePath: "word/_rels/main.xml.rels", blocking: true },
    { name: "dup-id-non-main-document-sidecar", parts: () => relocated(false, true), duplicatePath: RELS, blocking: false },
    {
        name: "dup-id-glossary-sidecar",
        parts: () => glossaryDuplicate(),
        duplicatePath: "word/glossary/_rels/document.xml.rels",
        blocking: true,
        lenientErrors: ["rels-id-duplicate", ...SCHEMA_ERRORS],
    },
    {
        name: "dup-id-unreached-glossary-sidecar",
        parts: () => glossaryDuplicate(false),
        duplicatePath: "word/glossary/_rels/document.xml.rels",
        blocking: false,
        lenientErrors: ["rels-unreferenced", "rels-id-duplicate", ...SCHEMA_ERRORS],
    },
    {
        name: "dup-id-bin-main-sidecar",
        parts: () => withDup(binMain(), "word/_rels/document.bin.rels"),
        duplicatePath: "word/_rels/document.bin.rels",
        blocking: true,
        lenientErrors: ["rels-id-duplicate", ...SCHEMA_ERRORS],
    },
    {
        name: "dup-id-bin-glossary-sidecar",
        parts: () => withDup(binGlossary(), "word/glossary/_rels/document.bin.rels"),
        duplicatePath: "word/glossary/_rels/document.bin.rels",
        blocking: true,
        lenientErrors: ["rels-id-duplicate", ...SCHEMA_ERRORS],
    },
    { name: "bin-main", parts: binMain, blocking: false, lenientErrors: [] },
    { name: "bin-glossary", parts: binGlossary, blocking: false, lenientErrors: [] },
];
const errorKeys = (issues: readonly ValidationIssue[]): string[] =>
    issues.filter((issue) => issue.severity === "error").map((issue) => `${issue.code}@${issue.path}`);

describe.each(["packed", "directory"] as const)("duplicate relationship Id scope, %s (#615, #617, #618)", { timeout: 120_000 }, (input) => {
    it.each(SPECIMENS)("$name: full profile findings and partial error-subset parity", async (specimen) => {
        await withTempDir(async (dir) => {
            const target = path.join(dir, `${specimen.name}.docx`);
            const parts = specimen.parts();
            if (input === "packed") await writeFile(target, await buildDocx(parts));
            else {
                for (const [part, xml] of Object.entries(parts)) {
                    const file = path.join(target, part);
                    await mkdir(path.dirname(file), { recursive: true });
                    await writeFile(file, xml);
                }
            }
            const word = await validate(target, { profile: "word-valid" });
            const strict = await validate(target, { profile: "strict" });
            expect.soft(word.valid).toBe(!specimen.blocking);
            const strictErrors = errorKeys(strict.issues);
            for (const key of errorKeys(word.issues)) expect.soft(strictErrors).toContain(key);
            if (specimen.duplicatePath) expect.soft(strict.valid).toBe(false);

            for (const profile of ["word-valid", "lenient", "strict"] as const) {
                const full = profile === "word-valid" ? word : profile === "strict" ? strict : await validate(target, { profile });
                const duplicates = full.issues.filter((issue) => issue.code === "rels-id-duplicate");
                const expectedPath = profile === "lenient" && specimen.duplicatePath === PACKAGE_RELS ? undefined : specimen.duplicatePath;
                expect
                    .soft(
                        duplicates.map((issue) => [issue.path, issue.severity]),
                        profile,
                    )
                    .toEqual(expectedPath ? [[expectedPath, profile === "word-valid" && !specimen.blocking ? "warning" : "error"]] : []);
                if (profile === "strict") continue;
                if (profile === "lenient" && specimen.lenientErrors) {
                    expect
                        .soft(
                            full.issues
                                .filter((issue) => issue.severity === "error")
                                .map((issue) => issue.code)
                                .sort(),
                        )
                        .toEqual([...specimen.lenientErrors].sort());
                    expect.soft(full.valid).toBe(specimen.lenientErrors.length === 0);
                }
                const selections =
                    profile === "lenient"
                        ? [{ checks: ["xml-syntax"] }, { checks: ["xml"] }, { quick: true }]
                        : [{ checks: ["opc-relationships-structure"] }, { checks: ["relationships"] }, { quick: true }];
                for (const selection of selections) {
                    const partial = await validate(target, { profile, ...selection });
                    if (full.valid) expect.soft(partial.valid, `${profile} ${JSON.stringify(selection)}`).toBe(true);
                    for (const key of errorKeys(partial.issues)) expect.soft(errorKeys(full.issues)).toContain(key);
                    // The legacy relationship-id walk is outside the named-check registry.
                    expect.soft(partial.issues.filter((issue) => issue.code === "rels-id-duplicate")).toEqual([]);
                }
            }
        });
    });
});
