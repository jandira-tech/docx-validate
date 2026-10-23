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
 * Unit tests for `src/lib/xml-syntax.ts`.
 *
 * This file holds the PURE half of the module — `mapXmlDiagnostics`, the
 * `ErrorDetail -> ValidationIssue` mapping — exercised with hand-written
 * diagnostic literals so the stable codes are pinned independently of
 * libxml2's message wording and of the wasm engine itself. The byte-level
 * half (`checkXmlSyntax`, `decodeXmlBytes` round trips) lives in
 * `tests/xml-syntax.integration.test.ts`, which drives the real engine.
 */

import { afterEach, describe, expect, it, vi } from "vitest";

import { mapXmlDiagnostics } from "../src/lib/xml-syntax";

describe("mapXmlDiagnostics", () => {
    it("maps a fatal well-formedness diagnostic to xml-syntax with the original part path and 1-based line/column", () => {
        const issues = mapXmlDiagnostics(
            [{ message: "Opening and ending tag mismatch: b line 3 and a\n", level: 3, line: 3, col: 11 }],
            "word/document.xml",
        );
        expect(issues).toEqual([
            {
                severity: "error",
                code: "xml-syntax",
                message: "Opening and ending tag mismatch: b line 3 and a",
                path: "word/document.xml",
                line: 3,
                column: 11,
            },
        ]);
    });

    it("maps an undeclared namespace prefix to the stable xml-undeclared-prefix code", () => {
        const issues = mapXmlDiagnostics(
            [{ message: "Namespace prefix w on p is not defined\n", level: 2, line: 1, col: 5 }],
            "word/document.xml",
        );
        expect(issues).toHaveLength(1);
        expect(issues[0]?.code).toBe("xml-undeclared-prefix");
        expect(issues[0]?.severity).toBe("error");
    });

    it("maps an undeclared prefix on an attribute to the same code", () => {
        const issues = mapXmlDiagnostics(
            [{ message: "Namespace prefix b for c on a is not defined\n", level: 2, line: 2, col: 3 }],
            "word/document.xml",
        );
        expect(issues.map((i) => i.code)).toEqual(["xml-undeclared-prefix"]);
    });

    it("maps an unsupported encoding declaration to xml-encoding", () => {
        const issues = mapXmlDiagnostics(
            [{ message: "Unsupported encoding: NotAnEncoding\n", level: 3, line: 1, col: 45 }],
            "word/document.xml",
        );
        expect(issues.map((i) => i.code)).toEqual(["xml-encoding"]);
    });

    it("maps invalid bytes for the declared encoding to xml-encoding", () => {
        const issues = mapXmlDiagnostics(
            [{ message: "Invalid bytes in character encoding\n", level: 3, line: 1, col: 40 }],
            "word/document.xml",
        );
        expect(issues.map((i) => i.code)).toEqual(["xml-encoding"]);
    });

    it("keeps warning-level diagnostics at warning severity", () => {
        const issues = mapXmlDiagnostics(
            [{ message: "Start tag expected, '<' not found\n", level: 1, line: 0, col: 0 }],
            "[Content_Types].xml",
        );
        expect(issues[0]?.severity).toBe("warning");
        expect(issues[0]?.path).toBe("[Content_Types].xml");
    });

    it("omits line and column when libxml2 does not provide a position", () => {
        const issues = mapXmlDiagnostics([{ message: "Document is empty\n", level: 3, line: 0, col: 0 }], "word/document.xml");
        expect(issues[0]).not.toHaveProperty("line");
        expect(issues[0]).not.toHaveProperty("column");
    });

    it("reports a line without a column when libxml2 has no column", () => {
        const issues = mapXmlDiagnostics(
            [{ message: "Extra content at the end of the document\n", level: 3, line: 7, col: 0 }],
            "word/document.xml",
        );
        expect(issues[0]?.line).toBe(7);
        expect(issues[0]).not.toHaveProperty("column");
    });

    it("maps every diagnostic, preserving order and keeping the part path of the original part", () => {
        const issues = mapXmlDiagnostics(
            [
                { message: "CharRef: invalid hexadecimal value\n", level: 3, line: 3, col: 8 },
                { message: "xmlParseCharRef: invalid xmlChar value 0\n", level: 3, line: 3, col: 8 },
            ],
            "word/document.xml",
        );
        expect(issues).toHaveLength(2);
        expect(issues.every((i) => i.path === "word/document.xml")).toBe(true);
        expect(issues.map((i) => i.message)).toEqual(["CharRef: invalid hexadecimal value", "xmlParseCharRef: invalid xmlChar value 0"]);
    });

    it("returns no issues for an empty diagnostic list", () => {
        expect(mapXmlDiagnostics([], "word/document.xml")).toEqual([]);
    });
});

describe("checkXmlSyntax — engine failures without structured detail", () => {
    afterEach(() => {
        vi.doUnmock("libxml2-wasm");
        vi.resetModules();
    });

    it("reports a single xml-syntax issue against the original part", async () => {
        vi.resetModules();
        vi.doMock("libxml2-wasm", () => ({
            XmlDocument: {
                fromBuffer: () => {
                    throw new Error("engine exploded");
                },
            },
            XmlParseError: class extends Error {},
            ParseOption: { XML_PARSE_NONET: 2048, XML_PARSE_NO_XXE: 8388608 },
        }));
        const { checkXmlSyntax } = await import("../src/lib/xml-syntax");

        expect(await checkXmlSyntax(new Uint8Array([0x3c]), "word/document.xml")).toEqual([
            { severity: "error", code: "xml-syntax", message: "engine exploded", path: "word/document.xml" },
        ]);
    });

    it("surfaces non-fatal parse warnings at warning severity and disposes the document", async () => {
        vi.resetModules();
        const dispose = vi.fn();
        vi.doMock("libxml2-wasm", () => ({
            XmlDocument: {
                fromBuffer: () => ({ warnings: [{ message: "something worth knowing\n", level: 1, line: 0, col: 0 }], dispose }),
            },
            XmlParseError: class extends Error {},
            ParseOption: { XML_PARSE_NONET: 2048, XML_PARSE_NO_XXE: 8388608 },
        }));
        const { checkXmlSyntax } = await import("../src/lib/xml-syntax");

        expect(await checkXmlSyntax(new Uint8Array([0x3c]), "word/document.xml")).toEqual([
            { severity: "warning", code: "xml-syntax", message: "something worth knowing", path: "word/document.xml" },
        ]);
        expect(dispose).toHaveBeenCalledTimes(1);
    });
});
