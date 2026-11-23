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
import { validate } from "../src/scripts/office/validate";
import { contentTypeTableOf, xmlPartNames } from "../src/scripts/office/checks/xml";
import path from "node:path";

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

describe("review fixes (PR #1281)", () => {
    it("returns a ceiling breach without parsing, so no engine diagnostics follow it (r4651164150)", async () => {
        // Unclosed on purpose: a parse would add an xml-syntax finding.
        const issues = await checkXmlSyntax(
            encode(`<a x="1" y="2" z="3">`),
            "word/document.xml",
            resolveResourceCeilings({ maxAttributes: 2 }),
        );
        expect(issues.map((issue) => issue.code)).toEqual(["xml-attribute-ceiling"]);
    });

    it("reports a caller-lowered depth ceiling the engine would accept (r4651164161, r4649504945)", async () => {
        const issues = await checkXmlSyntax(encode(nest(5)), "word/document.xml", resolveResourceCeilings({ maxDepth: 4 }));
        expect(issues).toHaveLength(1);
        expect(issues[0]).toMatchObject({ severity: "error", code: "xml-depth-ceiling", path: "word/document.xml" });
        expect(issues[0]!.message).toContain("maxDepth");
        // At or under the ceiling: nothing.
        expect(await checkXmlSyntax(encode(nest(4)), "word/document.xml", resolveResourceCeilings({ maxDepth: 4 }))).toEqual([]);
    });

    it("reads a comment in the internal subset as opaque when finding the DOCTYPE end (r4651164153)", () => {
        const breaches = scanXmlCeilings(`<!DOCTYPE a [<!-- [ -->]><a x="1" y="2"/>`, {
            maxDepth: 256,
            maxAttributes: 1,
            maxEntityDeclarations: 64,
        });
        expect(breaches.map((b) => b.kind)).toEqual(["attributes"]);
    });

    it("keeps the DOCTYPE entity search inside the declaration's span (r4651164156)", () => {
        // Entities after the DOCTYPE closes are not in its internal subset;
        // a search past the span would count them (and go quadratic on a run
        // of declarations).
        const text = `<!DOCTYPE a [<!ENTITY e "x">]><a/><!ENTITY f "y">`;
        const ceilings = { maxDepth: 256, maxAttributes: 1024, maxEntityDeclarations: 1 };
        const breaches = scanXmlCeilings(text, ceilings);
        expect(breaches.filter((b) => b.kind === "entity-declarations").map((b) => b.index)).toEqual([text.indexOf("<!ENTITY f")]);
        const many = "<!DOCTYPE a>".repeat(200_000) + "<a/>";
        const started = performance.now();
        scanXmlCeilings(many, ceilings);
        expect(performance.now() - started).toBeLessThan(1_000);
    });

    it("does not count a stray '<' as an element (r4649504999)", () => {
        expect(scanXmlCeilings(`<a>< < < <</a>`, { maxDepth: 1, maxAttributes: 1024, maxEntityDeclarations: 64 })).toEqual([]);
    });

    it("enforces caller ceilings in the lenient validator's XML pass (r4651164145)", async () => {
        const fixture = path.join(import.meta.dirname, "fixtures/working/package.drawing-arc-radius-integer.docx");
        const result = await validate(fixture, { profile: "lenient", ceilings: { maxAttributes: 1 } });
        expect(result.valid).toBe(false);
        expect(result.issues.some((issue) => issue.code === "xml-attribute-ceiling")).toBe(true);
    });

    it("never parses an over-ceiling content-types stream for XML-part discovery (PR #1281 review, xml.ts)", () => {
        const stream = encode(
            `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Override PartName="/word/document.bin" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml" Extension="x"/></Types>`,
        );
        expect(contentTypeTableOf(stream)).toBeDefined();
        const ceilings = resolveResourceCeilings({ maxAttributes: 2 });
        expect(contentTypeTableOf(stream, ceilings)).toBeUndefined();
        const names = ["[Content_Types].xml", "word/document.bin"];
        expect(xmlPartNames(names, () => stream).has("word/document.bin")).toBe(true);
        // Names alone decide, so the .bin part is not judged as XML.
        expect(xmlPartNames(names, () => stream, ceilings).has("word/document.bin")).toBe(false);
    });
});
