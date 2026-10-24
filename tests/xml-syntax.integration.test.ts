/*
 * Copyright 2026 Jandira Technologies, LLC
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *     http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */

/**
 * Integration tests for the built-in XML syntax checker driving the REAL
 * libxml2-wasm engine (`tests/xml-syntax.test.ts` covers the pure mapping).
 *
 * Also covers the wiring through `BaseSchemaValidator.validateXml()` and
 * `validate()`, which are the only production callers of the checker.
 */

import { promises as fs } from "node:fs";
import path from "node:path";

import JSZip from "jszip";
import { describe, expect, it } from "vitest";

import { withTempDir } from "../src/lib/run-cli";
import { parseXml, parseXmlBytes } from "../src/lib/xml-helpers";
import { checkXmlSyntax } from "../src/lib/xml-syntax";
import { xmlSyntaxCheck } from "../src/scripts/office/checks/xml";
import { contextFromDirectory } from "../src/scripts/office/package-context";
import { validate } from "../src/scripts/office/validate";
import { BaseSchemaValidator } from "../src/scripts/office/validators/base";

const W_NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
const CONTENT_TYPES =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
    '<Default Extension="xml" ContentType="application/xml"/>' +
    '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
    "</Types>";
const ROOT_RELS =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>' +
    "</Relationships>";

const ENCODER = new TextEncoder();

const bytes = (text: string): Uint8Array => ENCODER.encode(text);

const writePart = async (root: string, relative: string, content: string | Uint8Array): Promise<string> => {
    const file = path.join(root, relative);
    await fs.mkdir(path.dirname(file), { recursive: true });
    await fs.writeFile(file, content);
    return file;
};

describe("checkXmlSyntax — byte-level well-formedness", () => {
    it("reports the original part path and the 1-based line and column of the defect", async () => {
        // The lone `&` on line 3 is the only defect: a bare ampersand is a
        // well-formedness error (libxml2: "xmlParseEntityRef: no name").
        const part = ['<w:document xmlns:w="urn:w">', "  <w:body>", "    text & more", "  </w:body>", "</w:document>"].join("\n");
        const issues = await checkXmlSyntax(bytes(part), "word/document.xml");
        expect(issues).toHaveLength(1);
        expect(issues[0]?.code).toBe("xml-syntax");
        expect(issues[0]?.path).toBe("word/document.xml");
        expect(issues[0]?.line).toBe(3);
        expect(issues[0]?.column).toBe(11);
        // libxml2's own wording is preserved verbatim (minus its trailing newline).
        expect(issues[0]?.message).toContain("xmlParseEntityRef");
    });
});

describe("checkXmlSyntax — encodings", () => {
    /** UTF-16 bytes with (or without) a leading BOM, in the requested endianness. */
    const utf16 = (text: string, endian: "le" | "be", withBom: boolean): Uint8Array => {
        const out = new Uint8Array((withBom ? 2 : 0) + text.length * 2);
        let offset = 0;
        if (withBom) {
            out[0] = endian === "le" ? 0xff : 0xfe;
            out[1] = endian === "le" ? 0xfe : 0xff;
            offset = 2;
        }
        for (let i = 0; i < text.length; i += 1) {
            const code = text.charCodeAt(i);
            out[offset + i * 2] = endian === "le" ? code & 0xff : code >> 8;
            out[offset + i * 2 + 1] = endian === "le" ? code >> 8 : code & 0xff;
        }
        return out;
    };

    const UTF8_BOM = new Uint8Array([0xef, 0xbb, 0xbf]);
    const concat = (...parts: Uint8Array[]): Uint8Array => Uint8Array.from(parts.flatMap((part) => Array.from(part)));

    const DOC = '<a:doc xmlns:a="urn:a">café — 日本語</a:doc>';
    const TEXT = "café — 日本語";

    it.each([
        ["UTF-8 without BOM", bytes(DOC)],
        ["UTF-8 with BOM", concat(UTF8_BOM, bytes(DOC))],
        ["UTF-16LE with BOM and an encoding declaration", utf16('<?xml version="1.0" encoding="UTF-16"?>' + DOC, "le", true)],
        ["UTF-16BE with BOM and an encoding declaration", utf16('<?xml version="1.0" encoding="UTF-16"?>' + DOC, "be", true)],
        ["UTF-16LE without BOM, declared UTF-16LE", utf16('<?xml version="1.0" encoding="UTF-16LE"?>' + DOC, "le", false)],
        ["UTF-16BE without BOM, declared UTF-16BE", utf16('<?xml version="1.0" encoding="UTF-16BE"?>' + DOC, "be", false)],
    ])("parses %s from bytes", async (_label, input) => {
        expect(await checkXmlSyntax(input, "word/document.xml")).toEqual([]);
    });

    it.each([
        ["UTF-8 without BOM", bytes(DOC)],
        ["UTF-8 with BOM", concat(UTF8_BOM, bytes(DOC))],
        ["UTF-16LE with BOM", utf16('<?xml version="1.0" encoding="UTF-16"?>' + DOC, "le", true)],
        ["UTF-16BE with BOM", utf16('<?xml version="1.0" encoding="UTF-16"?>' + DOC, "be", true)],
        ["UTF-16LE without BOM", utf16('<?xml version="1.0" encoding="UTF-16LE"?>' + DOC, "le", false)],
        ["UTF-16BE without BOM", utf16('<?xml version="1.0" encoding="UTF-16BE"?>' + DOC, "be", false)],
    ])("decodes %s to a DOM carrying the original text", async (_label, input) => {
        const dom = parseXmlBytes(input);
        expect(dom.documentElement?.textContent).toBe(TEXT);
    });
});

describe("checkXmlSyntax — namespaces", () => {
    it("diagnoses an undeclared namespace prefix with the stable xml-undeclared-prefix code", async () => {
        const issues = await checkXmlSyntax(bytes("<w:p>hello</w:p>"), "word/document.xml");
        expect(issues.length).toBeGreaterThan(0);
        expect(issues.every((i) => i.code === "xml-undeclared-prefix")).toBe(true);
        expect(issues.every((i) => i.severity === "error")).toBe(true);
        expect(issues.every((i) => i.path === "word/document.xml")).toBe(true);
        expect(issues[0]?.line).toBe(1);
    });

    it("accepts the same prefix once it is declared", async () => {
        expect(await checkXmlSyntax(bytes('<w:p xmlns:w="urn:w">hello</w:p>'), "word/document.xml")).toEqual([]);
    });
});

describe("checkXmlSyntax — no parser recovery", () => {
    it.each([
        ["a bare ampersand in text", '<a:doc xmlns:a="urn:a">tea & scones</a:doc>'],
        ["an unquoted attribute value", '<a:doc xmlns:a="urn:a" b=c/>'],
    ])("reports %s, which @xmldom accepts in recovery mode, as xml-syntax", async (_label, doc) => {
        // @xmldom recovers silently from both, so the old byte path saw a
        // "valid" document. libxml2 without XML_PARSE_RECOVER does not.
        expect(() => parseXml(doc)).not.toThrow();
        const issues = await checkXmlSyntax(bytes(doc), "word/document.xml");
        expect(issues.length).toBeGreaterThan(0);
        expect(issues.every((i) => i.code === "xml-syntax")).toBe(true);
        expect(issues.every((i) => i.severity === "error")).toBe(true);
        expect(issues.every((i) => i.path === "word/document.xml")).toBe(true);
    });

    it("still reports the controls both parsers already rejected", async () => {
        expect((await checkXmlSyntax(bytes('<a:doc xmlns:a="urn:a"><a:p>text'), "word/document.xml")).map((i) => i.code)).toEqual([
            "xml-syntax",
        ]);
        expect((await checkXmlSyntax(bytes('<a:doc xmlns:a="urn:a" a:id="1" a:id="2"/>'), "word/document.xml")).map((i) => i.code)).toEqual(
            ["xml-syntax"],
        );
    });
});

describe("checkXmlSyntax — external entities", () => {
    it("never loads an external entity declared in a DOCTYPE", async () => {
        const part = '<!DOCTYPE x [<!ENTITY e SYSTEM "file:///etc/hosts">]><x>&e;</x>';
        const issues = await checkXmlSyntax(bytes(part), "word/document.xml");
        expect(issues).toEqual([]);
        expect(JSON.stringify(issues)).not.toContain("localhost");
    });
});

describe("checkXmlSyntax — unsupported encodings", () => {
    it("reports an unsupported encoding declaration as xml-encoding", async () => {
        const issues = await checkXmlSyntax(bytes('<?xml version="1.0" encoding="NotAnEncoding"?><a/>'), "word/document.xml");
        expect(issues.map((i) => i.code)).toEqual(["xml-encoding"]);
        expect(issues[0]?.severity).toBe("error");
    });
});

describe("BaseSchemaValidator.validateXml — original bytes", () => {
    it("reports the package-relative part path with line and column for a malformed part", async () => {
        await withTempDir(async (dir) => {
            await writePart(
                dir,
                "word/document.xml",
                ['<w:document xmlns:w="urn:w">', "  <w:body>", "    text & more", "  </w:body>"].join("\n"),
            );
            const validator = new BaseSchemaValidator({ unpackedDir: dir });
            const result = await validator.validateXml();
            expect(result.valid).toBe(false);
            expect(result.issues).toHaveLength(1);
            expect(result.issues[0]?.code).toBe("xml-syntax");
            expect(result.issues[0]?.path).toBe("word/document.xml");
            expect(result.issues[0]?.line).toBe(3);
            expect(result.issues[0]?.column).toBe(11);
        });
    });

    it("accepts a UTF-16 encoded part that the UTF-8 string read path could not decode", async () => {
        await withTempDir(async (dir) => {
            const text = '<?xml version="1.0" encoding="UTF-16"?><w:document xmlns:w="urn:w"><w:body>café</w:body></w:document>';
            const out = new Uint8Array(2 + text.length * 2);
            out[0] = 0xff;
            out[1] = 0xfe;
            for (let i = 0; i < text.length; i += 1) {
                out[2 + i * 2] = text.charCodeAt(i) & 0xff;
                out[3 + i * 2] = text.charCodeAt(i) >> 8;
            }
            await writePart(dir, "word/document.xml", out);
            const validator = new BaseSchemaValidator({ unpackedDir: dir });
            expect(await validator.validateXml()).toEqual({ valid: true, issues: [] });
        });
    });

    it("flags a malformed part @xmldom recovers from, instead of passing it as valid", async () => {
        await withTempDir(async (dir) => {
            await writePart(
                dir,
                "word/document.xml",
                '<?xml version="1.0"?><w:document xmlns:w="urn:w"><w:body>tea & scones</w:body></w:document>',
            );
            const validator = new BaseSchemaValidator({ unpackedDir: dir });
            const result = await validator.validateXml();
            expect(result.valid).toBe(false);
            expect(result.issues.map((i) => i.code)).toEqual(["xml-syntax"]);
        });
    });

    it.each(["word/COMMENTS.XML", "_RELS/.RELS", "word/Notes.Xml"])(
        "parses %s, whose extension is not lower case, as the xml-syntax check does (#390)",
        async (part) => {
            await withTempDir(async (dir) => {
                await writePart(dir, "word/document.xml", `<w:document ${W_NS}/>`);
                await writePart(dir, part, '<?xml version="1.0"?><w:comments xmlns:w="urn:w">');
                const result = await new BaseSchemaValidator({ unpackedDir: dir }).validateXml();
                const registered = await xmlSyntaxCheck.run(contextFromDirectory(dir));

                expect(result.valid).toBe(false);
                expect(result.issues).toContainEqual(expect.objectContaining({ code: "xml-syntax", path: part }));
                expect(result.issues).toEqual(registered);
            });
        },
    );
});

describe("BOM-prefixed XML is never a failure on its own", () => {
    const BOM = new Uint8Array([0xef, 0xbb, 0xbf]);
    const VALID_PART =
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ' + W_NS + "><w:body><w:p/></w:body></w:document>";

    const bomPart = async (dir: string): Promise<void> => {
        await writePart(dir, "word/document.xml", new Uint8Array([...BOM, ...Array.from(ENCODER.encode(VALID_PART))]));
    };

    it.each(["lenient", "strict", "word-valid"] as const)("reports xml-bom-leading as a warning under the %s profile", async (profile) => {
        await withTempDir(async (dir) => {
            await bomPart(dir);
            const result = await new BaseSchemaValidator({ unpackedDir: dir, profile }).validateNoBom();
            expect(result.issues.map((i) => i.code)).toEqual(["xml-bom-leading"]);
            expect(result.issues.every((i) => i.severity === "warning")).toBe(true);
            expect(result.valid).toBe(true);
        });
    });

    it.each(["lenient", "strict", "word-valid"] as const)(
        "never fails validate() solely because of a BOM under the %s profile",
        async (profile) => {
            await withTempDir(async (dir) => {
                const unpacked = path.join(dir, "unpacked");
                await writePart(unpacked, "[Content_Types].xml", new Uint8Array([...BOM, ...Array.from(ENCODER.encode(CONTENT_TYPES))]));
                await writePart(unpacked, "_rels/.rels", new Uint8Array([...BOM, ...Array.from(ENCODER.encode(ROOT_RELS))]));
                await bomPart(unpacked);
                const zip = new JSZip();
                zip.file("[Content_Types].xml", CONTENT_TYPES);
                zip.file(
                    "word/document.xml",
                    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:document ${W_NS}><w:body/></w:document>`,
                );
                const original = path.join(dir, "original.docx");
                await fs.writeFile(original, await zip.generateAsync({ type: "nodebuffer" }));

                const result = await validate(unpacked, { original, author: "Test", profile });
                const bomIssues = result.issues.filter((i) => i.code === "xml-bom-leading");
                expect(bomIssues.length).toBeGreaterThan(0);
                expect(bomIssues.every((i) => i.severity === "warning")).toBe(true);
                // Nothing else about the part is wrong, so the BOM must not be the
                // reason the run fails. Other checks may still report real defects.
                expect(result.issues.filter((i) => i.severity === "error").map((i) => i.code)).not.toContain("xml-bom-leading");
            });
        },
    );
});

describe("validate() — input is never rewritten", () => {
    /** Byte-for-byte snapshot of every file under `root`. */
    const snapshot = async (root: string): Promise<string[]> => {
        const out: string[] = [];
        const walk = async (dir: string): Promise<void> => {
            for (const entry of await fs.readdir(dir, { withFileTypes: true })) {
                const full = path.join(dir, entry.name);
                if (entry.isDirectory()) {
                    await walk(full);
                } else {
                    out.push(`${path.relative(root, full)}:${(await fs.readFile(full)).toString("hex")}`);
                }
            }
        };
        await walk(root);
        return out.sort();
    };

    it("leaves every byte of the input package unchanged", async () => {
        await withTempDir(async (dir) => {
            const unpacked = path.join(dir, "unpacked");
            // Deliberately non-canonical formatting (odd spacing, BOM, trailing
            // newline): pretty-printing or normalising would show up here.
            await writePart(
                unpacked,
                "[Content_Types].xml",
                new Uint8Array([0xef, 0xbb, 0xbf, ...Array.from(ENCODER.encode(CONTENT_TYPES))]),
            );
            await writePart(unpacked, "_rels/.rels", ROOT_RELS);
            await writePart(
                unpacked,
                "word/document.xml",
                '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:document   ' +
                    `${W_NS}>\n  <w:body>  <w:p><w:r><w:t>keep   my   spacing</w:t></w:r></w:p>  </w:body>\n</w:document>\n`,
            );
            const zip = new JSZip();
            zip.file("[Content_Types].xml", CONTENT_TYPES);
            zip.file(
                "word/document.xml",
                `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:document ${W_NS}><w:body/></w:document>`,
            );
            const original = path.join(dir, "original.docx");
            await fs.writeFile(original, await zip.generateAsync({ type: "nodebuffer" }));

            const before = await snapshot(unpacked);
            const result = await validate(unpacked, { original, author: "Test" });

            expect(await snapshot(unpacked)).toEqual(before);
            // Guard against a vacuous pass: the run really happened.
            expect(result.suffix).toBe(".docx");
            expect(result.issues.some((i) => i.code === "xml-bom-leading")).toBe(true);
        });
    });
});
