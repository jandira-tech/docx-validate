/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

import { mkdirSync, readFileSync, readdirSync, writeFileSync } from "node:fs";
import path from "node:path";
import { describe, expect, it } from "vitest";
import { withTempDir } from "../src/lib/run-cli";
import type { ValidationIssue } from "../src/lib/types";
import { validate } from "../src/scripts/office/validate";
import { BaseSchemaValidator } from "../src/scripts/office/validators/base";
import { buildDocx, MINIMAL_PARTS } from "./support/minimal-docx";

const R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const RS = "http://purl.oclc.org/ooxml/officeDocument/relationships";
const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const WS = "http://purl.oclc.org/ooxml/wordprocessingml/main";
const ESCAPED_NAME = "opc-part-name-escaped-unreserved";
const MAIN = "word/document.xml";
const SIDECAR = "word/_rels/document.xml.rels";
const TEXT = "<w:r><w:t>Hello</w:t></w:r>";
const docH = MINIMAL_PARTS[MAIN]!.replace("<w:document ", `<w:document xmlns:r="${R}" `).replace(
    TEXT,
    `<w:hyperlink r:id="rId2">${TEXT}</w:hyperlink>`,
);
const relsH = MINIMAL_PARTS[SIDECAR]!.replace(
    "</Relationships>",
    `<Relationship Id="rId2" Type="${R}/hyperlink" Target="https://example.com/" TargetMode="External"/></Relationships>`,
);
const hyperlinkParts = (extension: string): Record<string, string> => ({
    ...Object.fromEntries(Object.entries(MINIMAL_PARTS).filter(([name]) => name !== SIDECAR)),
    [MAIN]: docH,
    [`word/_rels/document.xml.${extension}`]: relsH,
});
const picture = `<w:r><w:drawing><wp:inline xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing">
    <wp:extent cx="95250" cy="95250"/><wp:docPr id="1" name="Picture 1"/>
    <a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
    <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
    <pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">
    <pic:nvPicPr><pic:cNvPr id="0" name="image.png"/><pic:cNvPicPr/></pic:nvPicPr>
    <pic:blipFill><a:blip r:embed="rId2"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>
    <pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="95250" cy="95250"/></a:xfrm>
    <a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>
    </pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>`;
const imageParts = (extension: string): Record<string, string> => ({
    ...hyperlinkParts(extension),
    [MAIN]: docH.replace(`<w:hyperlink r:id="rId2">${TEXT}</w:hyperlink>`, picture),
    [`word/_rels/document.xml.${extension}`]: relsH.replace(
        `Type="${R}/hyperlink" Target="https://example.com/" TargetMode="External"`,
        `Type="${R}/image" Target="media/image.svg"`,
    ),
    "[Content_Types].xml": MINIMAL_PARTS["[Content_Types].xml"]!.replace(
        "</Types>",
        '<Default Extension="svg" ContentType="image/svg+xml"/></Types>',
    ),
    "word/media/image.svg": '<svg xmlns="http://www.w3.org/2000/svg" width="1" height="1"><rect width="1" height="1"/></svg>',
});
const errors = (issues: readonly ValidationIssue[]): string[] =>
    issues.filter((issue) => issue.severity === "error").map((issue) => `${issue.code}@${issue.path}`);

describe("relationship references without a sidecar (#624)", () => {
    // Session 33, s33a/s33c: rewrite only document/styles WML, preserving
    // MINIMAL_PARTS' package relationships and content types as observed.
    const noSidecar = (strict: boolean, reference = ""): Record<string, string> => {
        const parts = Object.fromEntries(Object.entries(MINIMAL_PARTS).filter(([name]) => name !== SIDECAR));
        if (strict) {
            for (const part of [MAIN, "word/styles.xml"]) parts[part] = parts[part]!.replaceAll(W, WS);
        }
        if (reference) {
            parts[MAIN] = parts[MAIN]!.replace("<w:document ", `<w:document xmlns:r="${strict ? RS : R}" `).replace(
                "<w:sectPr>",
                `<w:sectPr>${reference}`,
            );
        }
        return parts;
    };
    const footerReference = '<w:footerReference w:type="default" r:id="rId9"/>';

    it.each([
        { strict: true, label: "s33a ISO Strict" },
        { strict: false, label: "session 32 Transitional twin" },
    ])("rejects $label with exactly one missing-sidecar finding", async ({ strict }) => {
        await withTempDir(async (dir) => {
            const file = path.join(dir, "footer-reference.docx");
            writeFileSync(file, await buildDocx(noSidecar(strict, footerReference)));
            const result = await validate(file, { profile: "word-valid" });
            expect(result.valid).toBe(false);
            expect(errors(result.issues)).toEqual([`rels-missing-sidecar@${MAIN}`]);
            expect(result.issues.filter((issue) => issue.code === "rels-missing-sidecar")).toEqual([
                expect.objectContaining({ path: MAIN, severity: "error" }),
            ]);
            // No sidecar belongs to the legacy walk alone, including s33a.
            expect(result.issues.filter((issue) => issue.code === "word-header-footer-relationship")).toEqual([]);
        });
    });

    it("keeps s33c ISO Strict without a sidecar or reference valid", async () => {
        await withTempDir(async (dir) => {
            const file = path.join(dir, "no-reference.docx");
            writeFileSync(file, await buildDocx(noSidecar(true)));
            const result = await validate(file, { profile: "word-valid" });
            expect(result.valid).toBe(true);
            expect(result.issues.filter((issue) => issue.code === "rels-missing-sidecar")).toEqual([]);
        });
    });

    // Same-footing namespace extension, not Word-observed: a Strict embed/link
    // alone on a descendant of a Transitional part still needs a sidecar.
    it.each(["embed", "link"] as const)("reports an ISO Strict r:%s in a Transitional part", async (attribute) => {
        await withTempDir(async (dir) => {
            mkdirSync(path.join(dir, "word"));
            writeFileSync(
                path.join(dir, MAIN),
                `<w:document xmlns:w="${W}"><w:body><w:p><w:r><w:drawing>` +
                    `<a:blip xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:rs="${RS}" rs:${attribute}="rId9"/>` +
                    "</w:drawing></w:r></w:p></w:body></w:document>",
            );
            const result = await new BaseSchemaValidator({ unpackedDir: dir }).validateAllRelationshipIds();
            expect(result.valid).toBe(false);
            expect(result.issues).toEqual([expect.objectContaining({ code: "rels-missing-sidecar", path: MAIN, severity: "error" })]);
        });
    });
});

describe.each(["packed", "directory"] as const)("OPC sidecar lookup, %s (#402)", { timeout: 120_000 }, (input) => {
    // Filesystem and ZIP operations are integration-only. Snapshot every file,
    // including stored names, to detect accidental writes without auto-repair.
    const snapshot = (target: string): unknown =>
        input === "packed"
            ? readFileSync(target)
            : readdirSync(target, { recursive: true, withFileTypes: true })
                  .filter((entry) => entry.isFile())
                  .map((entry) => {
                      const file = path.join(entry.parentPath, entry.name);
                      return [path.relative(target, file), readFileSync(file)];
                  });
    const writePackage = async (dir: string, label: string, parts: Record<string, string>): Promise<string> => {
        const target = path.join(dir, `${label}.docx`);
        if (input === "packed") writeFileSync(target, await buildDocx(parts));
        else
            for (const [name, text] of Object.entries(parts)) {
                const file = path.join(target, name);
                mkdirSync(path.dirname(file), { recursive: true });
                writeFileSync(file, text);
            }
        return target;
    };

    it("resolves the hyperlink triple in full, named and quick runs without changing input", async () => {
        await withTempDir(async (dir) => {
            let strictControl: ValidationIssue[] = [];
            for (const [index, extension] of ["rels", "RELS", "%72els"].entries()) {
                const file = await writePackage(dir, `variant-${index}`, hyperlinkParts(extension));
                const before = snapshot(file);
                const full = await validate(file, { profile: "word-valid" });
                const strict = await validate(file, { profile: "strict" });
                for (const result of [full, strict]) {
                    expect.soft(result.valid, extension).toBe(extension !== "%72els");
                    expect
                        .soft(
                            result.issues.filter((issue) =>
                                ["rels-missing-sidecar", "word-hyperlink-relationship"].includes(issue.code ?? ""),
                            ),
                            extension,
                        )
                        .toEqual([]);
                    if (extension === "%72els") {
                        expect.soft(errors(result.issues), extension).toEqual([`${ESCAPED_NAME}@word/_rels/document.xml.%72els`]);
                    }
                }
                if (extension === "rels") strictControl = strict.issues;
                if (extension === "RELS") {
                    expect.soft(strict.issues).toEqual(
                        strictControl.map((issue) => ({
                            ...issue,
                            path: issue.path?.replace(SIDECAR, "word/_rels/document.xml.RELS"),
                            message: issue.message.replaceAll(SIDECAR, "word/_rels/document.xml.RELS"),
                        })),
                    );
                }
                for (const selection of [
                    { checks: ["word-hyperlink-relationship"] },
                    { checks: ["word-hyperlink-relationship", ESCAPED_NAME] },
                    { quick: true },
                ]) {
                    const partial = await validate(file, { profile: "word-valid", ...selection });
                    expect
                        .soft(
                            partial.issues.filter((issue) => issue.code === "word-hyperlink-relationship"),
                            extension,
                        )
                        .toEqual([]);
                    if (full.valid) expect.soft(partial.valid, extension).toBe(true);
                    expect.soft(errors(full.issues), extension).toEqual(expect.arrayContaining(errors(partial.issues)));
                    if (extension === "%72els") {
                        if ("quick" in selection || selection.checks?.includes(ESCAPED_NAME)) {
                            expect.soft(partial.valid).toBe(false);
                            expect.soft(errors(partial.issues)).toContain(`${ESCAPED_NAME}@word/_rels/document.xml.%72els`);
                        } else {
                            expect.soft(partial.valid).toBe(true);
                            expect.soft(partial.metadata?.omittedChecks).toContain(ESCAPED_NAME);
                        }
                    }
                }
                expect(snapshot(file)).toEqual(before);
            }
        });
    });

    it("resolves image sidecars and preserves missing-sidecar controls", async () => {
        await withTempDir(async (dir) => {
            for (const [index, extension] of ["rels", "RELS", "%72els"].entries()) {
                const file = await writePackage(dir, `variant-${index}`, imageParts(extension));
                const before = snapshot(file);
                for (const profile of ["word-valid", "strict"] as const) {
                    const full = await validate(file, { profile });
                    expect
                        .soft(
                            full.issues.filter((issue) => issue.code === "rels-missing-sidecar"),
                            `${profile}/${extension}`,
                        )
                        .toEqual([]);
                }
                expect(snapshot(file)).toEqual(before);
            }
            // Negative controls retain the existing document error and the
            // Word-tolerated header/footer warning, in full and partial runs.
            for (const story of ["document", "header1", "footer1"]) {
                const part = `word/${story}.xml`;
                const root = story === "header1" ? "hdr" : "ftr";
                const parts =
                    story === "document"
                        ? Object.fromEntries(Object.entries(hyperlinkParts("rels")).filter(([name]) => name !== SIDECAR))
                        : {
                              ...MINIMAL_PARTS,
                              [part]: docH
                                  .replace(`<w:hyperlink r:id="rId2">${TEXT}</w:hyperlink>`, picture)
                                  .replaceAll("w:document", `w:${root}`)
                                  .replace("<w:body>", "")
                                  .replace("</w:body>", "")
                                  .replace(/<w:sectPr>.*<\/w:sectPr>/, ""),
                          };
                const file = await writePackage(dir, story, parts);
                const result = await validate(file, { profile: "word-valid" });
                expect(result.issues).toContainEqual(
                    expect.objectContaining({
                        code: "rels-missing-sidecar",
                        path: part,
                        severity: story === "document" ? "error" : "warning",
                    }),
                );
                expect(result.valid).toBe(story !== "document");
                for (const selection of [{ checks: ["word-hyperlink-relationship"] }, { quick: true }]) {
                    const partial = await validate(file, { profile: "word-valid", ...selection });
                    if (result.valid) expect(partial.valid).toBe(true);
                    expect(errors(result.issues)).toEqual(expect.arrayContaining(errors(partial.issues)));
                }
            }
        });
    });
});
