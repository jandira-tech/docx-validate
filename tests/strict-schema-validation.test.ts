/*
 * Copyright 2026 Jandira Technologies, LLC
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *     http://www.apache.org/licenses/LICENSE-2.0
 */

/**
 * Task 9 — strict-profile schema completeness.
 *
 * Under `profile: "strict"` an ISO OOXML Strict part is schema-validated
 * against the bundled Strict XSDs (namespace + part identity dispatch), not
 * skipped with `xsd-strict-skipped`. Lenient/word-valid keep the documented
 * skip behaviour.
 */

import { existsSync, promises as fs } from "node:fs";
import path from "node:path";

import { describe, expect, it } from "vitest";

import { withTempDir } from "../src/lib/run-cli";
import { validate } from "../src/scripts/office/validate";
import { DOCXSchemaValidator } from "../src/scripts/office/validators/docx";
import { schemaResourceRoots } from "../src/lib/xsd-validator";
import { getElementsByTagNameNSAll, parseXml } from "../src/lib/xml-helpers";
import { defaultSchemasDir, walkFiles } from "../src/scripts/office/validators/base";
import { MINIMAL_PARTS, buildDocx } from "./support/minimal-docx";

const TRANSITIONAL = "http://schemas.openxmlformats.org/";
const STRICT = "http://purl.oclc.org/ooxml/";

/** Rewrite the OOXML namespace URIs of the minimal package to Strict URIs. */
const strictify = (xml: string): string =>
    xml
        .replaceAll(`${TRANSITIONAL}wordprocessingml/2006/main`, `${STRICT}wordprocessingml/main`)
        .replaceAll(`${TRANSITIONAL}officeDocument/2006/relationships`, `${STRICT}officeDocument/relationships`);

const STRICT_PARTS: Readonly<Record<string, string>> = Object.freeze({
    ...MINIMAL_PARTS,
    "_rels/.rels": strictify(MINIMAL_PARTS["_rels/.rels"]!),
    "word/_rels/document.xml.rels": strictify(MINIMAL_PARTS["word/_rels/document.xml.rels"]!),
    "word/document.xml": strictify(MINIMAL_PARTS["word/document.xml"]!),
    "word/styles.xml": strictify(MINIMAL_PARTS["word/styles.xml"]!),
});

const writeDocx = async (dir: string, name: string, parts: Readonly<Record<string, string>>): Promise<string> => {
    const file = path.join(dir, name);
    await fs.writeFile(file, await buildDocx(parts));
    return file;
};

const inject = (xml: string, element: string): string => xml.replace("</w:body>", `${element}</w:body>`);

describe("strict profile — ISO Strict schema dispatch", () => {
    it("schema-validates a clean Strict document without xsd-strict-skipped", async () => {
        await withTempDir(async (dir) => {
            const file = await writeDocx(dir, "strict-clean.docx", STRICT_PARTS);
            const result = await validate(file, { profile: "strict" });
            expect(result.issues.some((i) => i.code === "xsd-strict-skipped")).toBe(false);
            expect(result.valid).toBe(true);
        });
    }, 60_000);

    it("fails a Strict document with an injected schema defect, naming the element", async () => {
        await withTempDir(async (dir) => {
            const parts = {
                ...STRICT_PARTS,
                "word/document.xml": inject(STRICT_PARTS["word/document.xml"]!, "<w:notInSchema/>"),
            };
            const file = await writeDocx(dir, "strict-defect.docx", parts);
            const result = await validate(file, { profile: "strict" });
            expect(result.valid).toBe(false);
            expect(result.issues.some((i) => i.code === "xsd-error" && i.message.includes("notInSchema"))).toBe(true);
        });
    }, 60_000);
});

describe("strict profile — original-baseline subtraction", () => {
    const DEFECTIVE_PARTS: Readonly<Record<string, string>> = Object.freeze({
        ...MINIMAL_PARTS,
        "word/document.xml": MINIMAL_PARTS["word/document.xml"]!.replace("</w:body>", "<w:notInSchema/></w:body>"),
    });

    it("still reports a schema defect that is also present in --original", async () => {
        await withTempDir(async (dir) => {
            const original = await writeDocx(dir, "original.docx", DEFECTIVE_PARTS);
            const modified = await writeDocx(dir, "modified.docx", DEFECTIVE_PARTS);
            const result = await validate(modified, { profile: "strict", original, author: "Test" });
            expect(result.valid).toBe(false);
            expect(result.issues.some((i) => i.code === "xsd-error" && i.message.includes("notInSchema"))).toBe(true);
        });
    }, 60_000);

    it("lenient keeps subtracting defects already present in --original", async () => {
        await withTempDir(async (dir) => {
            const original = await writeDocx(dir, "original.docx", DEFECTIVE_PARTS);
            const modified = await writeDocx(dir, "modified.docx", DEFECTIVE_PARTS);
            const result = await validate(modified, { profile: "lenient", original, author: "Test" });
            expect(result.issues.some((i) => i.code === "xsd-error" && i.message.includes("notInSchema"))).toBe(false);
            expect(result.valid).toBe(true);
        });
    }, 60_000);
});

describe("strict profile — one issue per error occurrence", () => {
    it("keeps distinct locations for repeated identical schema defects", async () => {
        await withTempDir(async (dir) => {
            // Three identical attribute defects on three different lines. The
            // deduplicated message set would collapse them to one; structured
            // reporting must keep one issue per occurrence with its own line.
            const para = (id: string, text: string): string =>
                `<w:p w14:paraId="${id}" w14:textId="${id}" w:bogus="${id}"><w:r><w:t>${text}</w:t></w:r></w:p>`;
            const documentXml = MINIMAL_PARTS["word/document.xml"]!.replace(
                "<w:sectPr>",
                `\n${para("00000002", "a")}\n${para("00000003", "b")}\n${para("00000004", "c")}\n<w:sectPr>`,
            );
            const file = await writeDocx(dir, "repeated-defect.docx", {
                ...MINIMAL_PARTS,
                "word/document.xml": documentXml,
            });
            const result = await validate(file, { profile: "strict" });
            expect(result.valid).toBe(false);
            const occurrences = result.issues.filter((i) => i.code === "xsd-error" && i.message.includes("bogus"));
            expect(occurrences.length).toBe(3);
            const lines = new Set(occurrences.map((i) => i.line));
            expect(lines.size).toBe(3);
        });
    }, 60_000);
});

describe("strict profile — missing schemas never pass silently", () => {
    const unpack = async (dir: string): Promise<string> => {
        const unpacked = path.join(dir, "unpacked");
        for (const [name, content] of Object.entries(MINIMAL_PARTS)) {
            const target = path.join(unpacked, name);
            await fs.mkdir(path.dirname(target), { recursive: true });
            await fs.writeFile(target, content);
        }
        return unpacked;
    };

    it("a schema directory without the required XSDs fails strict with xsd-schema-load-failed", async () => {
        await withTempDir(async (dir) => {
            const unpackedDir = await unpack(dir);
            const emptySchemas = path.join(dir, "empty-schemas");
            await fs.mkdir(emptySchemas, { recursive: true });
            const validator = new DOCXSchemaValidator({
                unpackedDir,
                schemasDir: emptySchemas,
                profile: "strict",
            });
            const result = await validator.validate();
            expect(result.valid).toBe(false);
            expect(result.issues.some((i) => i.code === "xsd-schema-load-failed")).toBe(true);
        });
    }, 60_000);

    it("an unresolvable import reported by the engine stays a strict error", async () => {
        await withTempDir(async (dir) => {
            const unpackedDir = await unpack(dir);
            const validator = new DOCXSchemaValidator({
                unpackedDir,
                profile: "strict",
                xsdValidator: {
                    validate: async () => [
                        {
                            severity: "error",
                            code: "xsd-schema-load-failed",
                            message: "Schema load failed: cannot resolve import",
                        },
                    ],
                },
            });
            const result = await validator.validate();
            expect(result.valid).toBe(false);
            expect(result.issues.some((i) => i.code === "xsd-schema-load-failed")).toBe(true);
        });
    }, 60_000);
});

const CORE_XML_HEAD =
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n` +
    `<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"` +
    ` xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/"` +
    ` xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">`;

const withCoreProps = (body: string): Readonly<Record<string, string>> =>
    Object.freeze({
        ...MINIMAL_PARTS,
        "[Content_Types].xml": MINIMAL_PARTS["[Content_Types].xml"]!.replace(
            "</Types>",
            `<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/></Types>`,
        ),
        "_rels/.rels": MINIMAL_PARTS["_rels/.rels"]!.replace(
            "</Relationships>",
            `<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/></Relationships>`,
        ),
        "docProps/core.xml": `${CORE_XML_HEAD}${body}</cp:coreProperties>`,
    });

describe("strict profile — Dublin Core core properties", () => {
    it("validates a clean core.xml against the bundled DC schemas", async () => {
        await withTempDir(async (dir) => {
            const file = await writeDocx(
                dir,
                "clean-core.docx",
                withCoreProps(`<dc:title>T</dc:title><dcterms:created xsi:type="dcterms:W3CDTF">2026-01-01T00:00:00Z</dcterms:created>`),
            );
            const result = await validate(file, { profile: "strict" });
            expect(result.issues.some((i) => i.code === "xsd-schema-load-failed")).toBe(false);
            expect(result.issues.every((i) => !i.message.includes("dublincore.org") && !i.message.includes("purl.org/dc/"))).toBe(true);
            expect(result.valid).toBe(true);
        });
    }, 60_000);

    it("fails strict on a real Dublin Core defect in core.xml", async () => {
        await withTempDir(async (dir) => {
            const file = await writeDocx(dir, "bad-core.docx", withCoreProps(`<dc:title><dc:notAValidChild/></dc:title>`));
            const result = await validate(file, { profile: "strict" });
            expect(result.valid).toBe(false);
            expect(result.issues.some((i) => i.code === "xsd-error" && i.message.includes("notAValidChild"))).toBe(true);
        });
    }, 60_000);
});

const XSD_NAMESPACE = "http://www.w3.org/2001/XMLSchema";

describe("bundled schema trees resolve offline", () => {
    it("the bundled schemas directory is an allowed libxml2 resource root", () => {
        expect(schemaResourceRoots()).toContain(defaultSchemasDir());
    });

    it("every xs:import/include schemaLocation resolves to a bundled file", async () => {
        const schemasDir = defaultSchemasDir();
        const missing: string[] = [];
        const remote: string[] = [];
        for (const xsdFile of walkFiles(schemasDir, [".xsd"])) {
            const dom = parseXml(await fs.readFile(xsdFile, "utf-8"));
            for (const elem of [
                ...getElementsByTagNameNSAll(dom, XSD_NAMESPACE, "import"),
                ...getElementsByTagNameNSAll(dom, XSD_NAMESPACE, "include"),
                ...getElementsByTagNameNSAll(dom, XSD_NAMESPACE, "redefine"),
            ]) {
                const location = elem.getAttribute("schemaLocation");
                if (!location) continue;
                if (/^[a-zA-Z][a-zA-Z0-9+.-]*:/.test(location)) {
                    remote.push(`${xsdFile}: ${location}`);
                    continue;
                }
                if (!existsSync(path.resolve(path.dirname(xsdFile), location))) {
                    missing.push(`${xsdFile}: ${location}`);
                }
            }
        }
        expect(remote).toEqual([]);
        expect(missing).toEqual([]);
    });
});

describe("strict profile — non-OOXML parts", () => {
    it("a customXml payload in an unknown namespace is not a missing schema", async () => {
        await withTempDir(async (dir) => {
            const parts = {
                ...MINIMAL_PARTS,
                "customXml/item1.xml": `<?xml version="1.0"?><invoice xmlns="http://example.com/bespoke"><line/></invoice>`,
            };
            const file = await writeDocx(dir, "custom-xml.docx", parts);
            const result = await validate(file, { profile: "strict" });
            // No OOXML schema applies to arbitrary customXml payloads — the
            // part is simply not XSD-eligible, so no xsd-* issue (and in
            // particular no schema-load failure) may be reported for it.
            const xsdIssues = result.issues.filter((i) => i.code?.startsWith("xsd-"));
            expect(xsdIssues).toEqual([]);
        });
    }, 60_000);
});

describe("strict profile — no message suppression, Strict-aware markup compatibility", () => {
    it("reports a misplaced w:hyphenationZone instead of suppressing it by message text", async () => {
        await withTempDir(async (dir) => {
            const parts = {
                ...MINIMAL_PARTS,
                "word/document.xml": inject(MINIMAL_PARTS["word/document.xml"]!, '<w:hyphenationZone w:val="1"/>'),
            };
            const file = await writeDocx(dir, "hyphenation-zone-misplaced.docx", parts);
            const result = await validate(file, { profile: "strict" });
            expect(result.valid).toBe(false);
            expect(result.issues.some((i) => i.code === "xsd-error" && i.message.includes("hyphenationZone"))).toBe(true);
        });
    }, 60_000);

    it("treats a Strict part's own namespace as understood for mc:MustUnderstand", async () => {
        await withTempDir(async (dir) => {
            const parts = {
                ...STRICT_PARTS,
                "word/document.xml": STRICT_PARTS["word/document.xml"]!.replace(
                    'mc:Ignorable="w14"',
                    'mc:Ignorable="w14" mc:MustUnderstand="w"',
                ),
            };
            expect(parts["word/document.xml"]).toContain('mc:MustUnderstand="w"');
            const file = await writeDocx(dir, "strict-must-understand.docx", parts);
            const result = await validate(file, { profile: "strict" });
            expect(result.issues.filter((i) => i.code === "mce-must-understand")).toEqual([]);
            expect(result.valid).toBe(true);
        });
    }, 60_000);
});
