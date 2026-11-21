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

// cspell:words NOENT
/**
 * Unit tests for the XML resource-ceiling pre-scan
 * (`src/lib/xml-syntax.ts`: `scanXmlCeilings`, `xmlCeilingIssues`, and the
 * guard inside `ValidationContext.xml()`), which refuses attribute bombs,
 * entity bombs and deep nesting in one linear pass before a DOM is built.
 */

import { describe, expect, it } from "vitest";

import { resolveResourceCeilings } from "../src/lib/resource-ceilings";
import { checkXmlSyntax, scanXmlCeilings, xmlCeilingIssues } from "../src/lib/xml-syntax";
import { ValidationContext } from "../src/lib/validation-context";

const encode = (text: string): Uint8Array => new TextEncoder().encode(text);

const nest = (depth: number): string => `${"<a>".repeat(depth)}</a>`.repeat(1) + "</a>".repeat(depth - 1);

describe("scanXmlCeilings", () => {
    it("reads element depth, counting a self-closing tag at the depth it would open", () => {
        const ceilings = resolveResourceCeilings({ maxDepth: 4 });
        expect(scanXmlCeilings(nest(4), ceilings)).toEqual([]);
        const breach = scanXmlCeilings(nest(5), ceilings);
        expect(breach).toHaveLength(1);
        expect(breach[0]).toMatchObject({ kind: "depth", observed: 5, limit: 4 });
    });

    it("skips comments, CDATA, processing instructions and declarations", () => {
        const ceilings = resolveResourceCeilings({ maxDepth: 3, maxEntityDeclarations: 10 });
        const text = `<?xml version="1.0"?><!-- <a><a><a> --><a><![CDATA[<a><a>]]><b/></a>`;
        expect(scanXmlCeilings(text, ceilings)).toEqual([]);
    });

    it("counts attributes per start tag by their quoted values", () => {
        const ceilings = resolveResourceCeilings({ maxAttributes: 3 });
        expect(scanXmlCeilings(`<a x="1" y="2" z="3"/>`, ceilings)).toEqual([]);
        const breach = scanXmlCeilings(`<a x="1" y="2" z="3" w="4"/>`, ceilings);
        expect(breach).toHaveLength(1);
        expect(breach[0]).toMatchObject({ kind: "attributes", observed: 4, limit: 3 });
    });

    it("counts <!ENTITY declarations, even inside the internal subset", () => {
        const ceilings = resolveResourceCeilings({ maxEntityDeclarations: 2 });
        const text = `<!DOCTYPE a [<!ENTITY e1 "x"><!ENTITY e2 "y">]><a>&e1;</a>`;
        expect(scanXmlCeilings(text, ceilings)).toEqual([]);
        const breach = scanXmlCeilings(`${text}<!ENTITY e3 "z">`, ceilings);
        expect(breach.map((b) => b.kind)).toContain("entity-declarations");
    });

    it("reports at most one breach per kind and stops on malformed markup", () => {
        const ceilings = resolveResourceCeilings({ maxDepth: 2, maxAttributes: 1, maxEntityDeclarations: 1 });
        const breaches = scanXmlCeilings(`<a x="1" y="2"><b><c></b>`, ceilings);
        expect(breaches.map((b) => b.kind).sort()).toEqual(["attributes", "depth"]);
    });
});

describe("xmlCeilingIssues / checkXmlSyntax", () => {
    it("maps attribute and entity breaches to named error findings", () => {
        const ceilings = resolveResourceCeilings({ maxAttributes: 2, maxEntityDeclarations: 1 });
        const issues = xmlCeilingIssues(`<a x="1" y="2" z="3"/>`, "word/document.xml", ceilings);
        expect(issues).toHaveLength(1);
        expect(issues[0]).toMatchObject({ severity: "error", code: "xml-attribute-ceiling", path: "word/document.xml" });
        expect(issues[0]!.message).toContain("maxAttributes");
    });

    it("leaves depth to libxml2's own refusal instead of double-reporting", async () => {
        const ceilings = resolveResourceCeilings({ maxDepth: 256 });
        const issues = await checkXmlSyntax(encode(nest(257)), "word/document.xml", ceilings);
        expect(issues.some((issue) => issue.message.includes("Excessive depth"))).toBe(true);
        expect(issues.some((issue) => issue.code === "xml-attribute-ceiling" || issue.code === "xml-entity-declaration-ceiling")).toBe(
            false,
        );
    });

    it("refuses an attribute bomb with the named ceiling through checkXmlSyntax", async () => {
        const attrs = Array.from({ length: 20 }, (_, i) => `a${i}="v"`).join(" ");
        const ceilings = resolveResourceCeilings({ maxAttributes: 5 });
        const issues = await checkXmlSyntax(encode(`<a ${attrs}/>`), "word/document.xml", ceilings);
        expect(issues.some((issue) => issue.code === "xml-attribute-ceiling" && issue.severity === "error")).toBe(true);
    });

    it("refuses an entity bomb with the named ceiling; external entities stay unexpanded regardless", async () => {
        const dtd = Array.from({ length: 5 }, (_, i) => `<!ENTITY e${i} "x">`).join("");
        const xml = `<?xml version="1.0"?><!DOCTYPE a [${dtd}]><a>&e0;</a>`;
        const ceilings = resolveResourceCeilings({ maxEntityDeclarations: 2 });
        const issues = await checkXmlSyntax(encode(xml), "word/document.xml", ceilings);
        expect(issues.some((issue) => issue.code === "xml-entity-declaration-ceiling")).toBe(true);
        // NO_XXE / no NOENT: the SYSTEM entity is never fetched.
        const xxe = await checkXmlSyntax(
            encode(`<!DOCTYPE a [<!ENTITY x SYSTEM "file:///etc/passwd">]><a>&x;</a>`),
            "word/document.xml",
            ceilings,
        );
        expect(JSON.stringify(xxe)).not.toContain("root:");
    });
});

describe("ValidationContext.xml() pre-DOM guard", () => {
    it("never builds a DOM for a part over a ceiling", () => {
        const attrs = Array.from({ length: 10 }, (_, i) => `a${i}="v"`).join(" ");
        const ctx = ValidationContext.fromEntries({ "word/document.xml": `<a ${attrs}/>` }, { ceilings: { maxAttributes: 5 } });
        const parsed = ctx.xml("word/document.xml");
        expect(parsed?.ok).toBe(false);
        expect(parsed?.ok === false && parsed.message).toContain("attributes");
        expect(ctx.parseCount("word/document.xml")).toBe(0);
    });

    it("parses normally under the ceilings", () => {
        const ctx = ValidationContext.fromEntries({ "word/document.xml": `<a x="1"/>` }, { ceilings: { maxAttributes: 5 } });
        expect(ctx.xml("word/document.xml")?.ok).toBe(true);
        expect(ctx.parseCount("word/document.xml")).toBe(1);
    });
});
