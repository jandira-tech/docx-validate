/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

import { describe, expect, it } from "vitest";
import { readdirSync, readFileSync } from "node:fs";
import path from "node:path";

import { applyMarkupCompatibility, STRICT_UNDERSTOOD_NAMESPACES, UNDERSTOOD_NAMESPACES } from "../src/lib/markup-compatibility";
import { defaultSchemasDir } from "../src/scripts/office/validators/base";
import { NS } from "../src/lib/types";
import { parseXml, serializeXml, getElementsByTagNameAll } from "../src/lib/xml-helpers";

const WP = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
const WP14 = "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing";
const process = (xml: string) => applyMarkupCompatibility(parseXml(xml), { understood: UNDERSTOOD_NAMESPACES });
const wrap = (xml: string, attrs = "") =>
    `<w:root xmlns:w="${NS.W}" xmlns:mc="${NS.MC}" xmlns:wp="${WP}" xmlns:wp14="${WP14}" ${attrs}>${xml}</w:root>`;
const alternative = (axis: string) =>
    `<mc:AlternateContent><mc:Choice Requires="wp14"><wp:position${axis} relativeFrom="page"><wp14:pctPos${axis}Offset>45500</wp14:pctPos${axis}Offset></wp:position${axis}></mc:Choice><mc:Fallback><wp:position${axis} relativeFrom="page"><wp:posOffset>4576445</wp:posOffset></wp:position${axis}></mc:Fallback></mc:AlternateContent>`;

describe("markup compatibility", () => {
    it("preserves the application extension payload (#293)", () => {
        const xml = wrap(
            `<a:ext xmlns:a="${NS.A}" uri="test"><e:payload mc:MustUnderstand="e"><mc:AlternateContent><mc:Choice Requires="w"><w:p/></mc:Choice><mc:Fallback><e:backup/></mc:Fallback></mc:AlternateContent></e:payload></a:ext>`,
            'xmlns:e="urn:future"',
        );
        const result = process(xml);
        expect(result.issues).toEqual([]);
        expect(serializeXml(result.document)).toBe(serializeXml(parseXml(xml)));
    });
    it("stops all MCE processing at every recognised extension boundary (#293)", () => {
        for (const namespace of [
            "http://schemas.openxmlformats.org/drawingml/2006/chart",
            NS.A,
            NS.P,
            NS.S,
            "http://purl.oclc.org/ooxml/drawingml/chart",
            "http://purl.oclc.org/ooxml/drawingml/main",
            "http://purl.oclc.org/ooxml/presentationml/main",
            "http://purl.oclc.org/ooxml/spreadsheetml/main",
        ]) {
            const xml = wrap(
                `<x:ext xmlns:x="${namespace}" uri="test" mc:Ignorable="missing"><wp14:payload wp14:flag="keep" mc:MustUnderstand="missing" mc:ProcessContent="w:p:extra"><mc:AlternateContent><mc:Fallback/></mc:AlternateContent></wp14:payload></x:ext>`,
            );
            const result = process(xml);
            expect(result.issues).toEqual([]);
            expect(serializeXml(result.document)).toBe(serializeXml(parseXml(xml)));
        }
    });
    it("still processes an ext element outside recognised extension namespaces (#293)", () => {
        const result = process(wrap('<e:ext xmlns:e="urn:future" mc:MustUnderstand="e"/>'));
        expect(result.issues).toEqual([expect.objectContaining({ severity: "error", code: "mce-must-understand" })]);
    });
    it("reports an unbound Ignorable prefix (#296)", () => {
        const result = process(wrap("<w:p/>", 'mc:Ignorable="missing"'));
        expect(result.issues).toContainEqual(expect.objectContaining({ severity: "error", code: "mce-unbound-prefix" }));
    });
    it("does not diagnose invalid directives inside discarded content (#295, #296)", () => {
        const result = process(
            wrap(
                '<e:discard mc:Ignorable="missing" mc:ProcessContent="missing:child w:p:extra" mc:MustUnderstand="missing"/>',
                'xmlns:e="urn:future" mc:Ignorable="e"',
            ),
        );
        expect(result.issues).toEqual([]);
    });
    it("accepts Unicode ProcessContent names and wildcards (#296)", () => {
        for (const token of ["e:子", "e:*"]) {
            const result = process(wrap(`<e:子><w:p/></e:子>`, `xmlns:e="urn:future" mc:Ignorable="e" mc:ProcessContent="${token}"`));
            expect(result.issues).toEqual([]);
            expect(getElementsByTagNameAll(result.document).map((node) => node.localName)).toEqual(["root", "p"]);
        }
    });
    it("reports an unbound ProcessContent prefix (#296)", () => {
        const result = process(wrap("<w:p/>", 'mc:ProcessContent="missing:child"'));
        expect(result.issues).toContainEqual(expect.objectContaining({ severity: "error", code: "mce-unbound-prefix" }));
    });
    it("rejects malformed ProcessContent tokens (#296)", () => {
        for (const token of ["w:p:extra", "w:", ":p", "w", "w:bad?name", "w:1child"]) {
            const result = process(wrap("<w:p/>", `mc:ProcessContent="${token}"`));
            expect(result.issues, token).toContainEqual(expect.objectContaining({ severity: "error", code: "mce-invalid-token" }));
        }
    });
    it("reports an unbound MustUnderstand prefix without claiming an unsupported capability (#296)", () => {
        const result = process(wrap("<w:p/>", 'mc:MustUnderstand="missing"'));
        expect(result.issues).toEqual([expect.objectContaining({ severity: "error", code: "mce-unbound-prefix" })]);
    });
    it("checks directives on unselected branches without inspecting their contents (#296)", () => {
        for (const attribute of ['mc:Ignorable="missing"', 'mc:ProcessContent="missing:child"', 'mc:MustUnderstand="missing"']) {
            for (const branch of ["Choice", "Fallback"]) {
                const unselected = `<mc:${branch} ${branch === "Choice" ? 'Requires="wp14"' : ""} ${attribute}><w:p mc:MustUnderstand="wp14"/></mc:${branch}>`;
                const chosen = '<mc:Choice Requires="w"><w:p/></mc:Choice>';
                const content = branch === "Choice" ? unselected + chosen : chosen + unselected;
                const result = process(wrap(`<mc:AlternateContent>${content}</mc:AlternateContent>`));
                expect(result.issues).toEqual([expect.objectContaining({ severity: "error", code: "mce-unbound-prefix" })]);
            }
        }
    });
    it("unwraps root AlternateContent (#292)", () => {
        const xml = `<mc:AlternateContent xmlns:mc="${NS.MC}" xmlns:w="${NS.W}"><mc:Choice Requires="w"><w:document><w:body/></w:document></mc:Choice></mc:AlternateContent>`;
        expect(() => process(xml)).not.toThrow();
        expect(process(xml).document.documentElement?.localName).toBe("document");
    });
    it("allows an ignorable AlternateContent child (#294)", () => {
        const result = process(
            wrap(
                '<mc:AlternateContent><mc:Choice Requires="w"><w:p/></mc:Choice><e:newChoice/></mc:AlternateContent>',
                'xmlns:e="urn:future" mc:Ignorable="e"',
            ),
        );
        expect(result.issues.filter((issue) => issue.severity === "error")).toEqual([]);
        expect(getElementsByTagNameAll(result.document).map((node) => node.localName)).toEqual(["root", "p"]);
    });
    it("does not inspect MustUnderstand on a discarded element (#295)", () => {
        const result = process(wrap('<e:discard mc:MustUnderstand="e"/>', 'xmlns:e="urn:future" mc:Ignorable="e"'));
        expect(result.issues.filter((issue) => issue.severity === "error")).toEqual([]);
        expect(getElementsByTagNameAll(result.document).map((node) => node.localName)).toEqual(["root"]);
    });
    it("still enforces MustUnderstand on retained and ProcessContent elements (#295)", () => {
        for (const content of ['<w:p mc:MustUnderstand="e"/>', '<e:keep mc:MustUnderstand="e"><w:p/></e:keep>']) {
            const result = process(wrap(content, 'xmlns:e="urn:future" mc:Ignorable="e" mc:ProcessContent="e:keep"'));
            expect(result.issues).toEqual([expect.objectContaining({ severity: "error", code: "mce-must-understand" })]);
            expect(getElementsByTagNameAll(result.document).map((node) => node.localName)).toEqual(["root", "p"]);
        }
    });
    it("removes non-ignorable Office attributes and aggregates one namespace-aware warning per part", () => {
        const source = parseXml(
            wrap(
                `<w:p wp14:paraId="1" wp14:textId="2"/><w:p wp14:paraId="3"/><w:r xmlns:alias="${WP14}" alias:paraId="4"/><w:p xmlns:w15="${NS.W15}" w15:flag="5"/>`,
            ),
        );
        const before = serializeXml(source);
        const { document, issues } = applyMarkupCompatibility(source, {
            understood: UNDERSTOOD_NAMESPACES,
        });
        expect(
            getElementsByTagNameAll(document)
                .flatMap((node) => Array.from(node.attributes))
                .filter((attr) => attr.namespaceURI?.startsWith("http://schemas.microsoft.com/office/")),
        ).toEqual([]);
        expect(issues).toHaveLength(1);
        expect(issues[0]).toMatchObject({ code: "mce-extension-not-ignorable", severity: "warning" });
        for (const value of [WP14, NS.W15, "wp14:paraId", "wp14:textId", "alias:paraId", "w15:flag", "mc:Ignorable"]) {
            expect(issues[0].message).toContain(value);
        }
        expect(issues[0].message.match(/wp14:paraId/g)).toHaveLength(1);
        expect(issues[0].message).toContain("5 attributes");
        expect(serializeXml(source)).toBe(before);
    });
    it("silently removes Office attributes when their URI is ignorable in scope", () => {
        const { document, issues } = process(wrap(`<w:p xmlns:alias="${WP14}" mc:Ignorable="alias"><w:r wp14:flag="gone"/></w:p>`));
        expect(serializeXml(document)).not.toContain("wp14:flag=");
        expect(issues).toEqual([]);
    });
    it("keeps truly unknown non-ignorable attributes and elements for XSD", () => {
        const { document, issues } = process(
            wrap(
                '<foo:foreign xmlns:foo="urn:example:foo" foo:flag="keep"/><w:p xmlns:near="http://schemas.microsoft.com/officeish/test" near:flag="keep"/>',
            ),
        );
        expect(serializeXml(document)).toContain('foo:flag="keep"');
        expect(serializeXml(document)).toContain('near:flag="keep"');
        expect(getElementsByTagNameAll(document).some((node) => node.namespaceURI === "urn:example:foo")).toBe(true);
        expect(issues).toEqual([]);
    });
    it("keeps fallback anchor positions in schema order", () => {
        const { document, issues } = process(
            wrap(`<wp:simplePos x="0" y="0"/>${alternative("H")}${alternative("V")}<wp:extent cx="1" cy="1"/>`, 'mc:Ignorable="wp14"'),
        );
        expect(Array.from(document.documentElement.childNodes).map((node) => node.nodeName)).toEqual([
            "wp:simplePos",
            "wp:positionH",
            "wp:positionV",
            "wp:extent",
        ]);
        expect(
            getElementsByTagNameAll(document)
                .filter((node) => node.localName === "posOffset")
                .map((node) => node.textContent),
        ).toEqual(["4576445", "4576445"]);
        expect(issues).toEqual([]);
    });
    it("selects the first choice whose namespace URIs are all understood", () => {
        const result = process(
            wrap(
                `<mc:AlternateContent><mc:Choice Requires="wp14"><w:bad/></mc:Choice><mc:Choice xmlns:alias="${NS.W}" Requires="alias wp"><w:chosen/></mc:Choice><mc:Choice Requires="w"><w:later/></mc:Choice><mc:Fallback><w:fallback/></mc:Fallback></mc:AlternateContent>`,
            ),
        );
        expect(getElementsByTagNameAll(result.document).map((node) => node.localName)).toEqual(["root", "chosen"]);
        const empty = process(wrap('<mc:AlternateContent><mc:Choice Requires="wp14"><w:bad/></mc:Choice></mc:AlternateContent>'));
        expect(empty.document.documentElement.childNodes.length).toBe(0);
    });
    it("inherits URI-scoped directives and recursively processes nested alternatives", () => {
        const { document, issues } = process(
            wrap(
                `<w:p xmlns:other="${WP14}" xmlns:p="${WP14}" mc:ProcessContent="p:keep p:*" mc:MustUnderstand="p"><other:keep><mc:AlternateContent><mc:Choice Requires="w">${alternative("H")}</mc:Choice><mc:Fallback><w:bad/></mc:Fallback></mc:AlternateContent></other:keep><other:any><w:kept/></other:any><w:r other:flag="gone" mc:PreserveElements="p:keep" mc:PreserveAttributes="p:flag"/></w:p><wp14:drop><w:lost/></wp14:drop><w:tail/>`,
                'mc:Ignorable="wp14"',
            ),
        );
        expect(getElementsByTagNameAll(document).map((node) => node.localName)).toEqual([
            "root",
            "p",
            "positionH",
            "posOffset",
            "kept",
            "r",
            "tail",
        ]);
        expect(issues).toEqual([expect.objectContaining({ code: "mce-must-understand", severity: "error" })]);
        for (const node of getElementsByTagNameAll(document)) {
            expect(Array.from(node.attributes).some((attr) => attr.namespaceURI === NS.MC || attr.namespaceURI === WP14)).toBe(false);
        }
    });
    it("never selects a Choice without Requires; the Fallback wins, or nothing (#494)", () => {
        for (const [label, choice] of [
            ["absent Requires", "<mc:Choice><w:chosen/></mc:Choice>"],
            ["empty Requires", '<mc:Choice Requires=""><w:chosen/></mc:Choice>'],
            ["blank Requires", '<mc:Choice Requires="   "><w:chosen/></mc:Choice>'],
        ] as const) {
            const withFallback = process(
                wrap(`<mc:AlternateContent>${choice}<mc:Fallback><w:fallback/></mc:Fallback></mc:AlternateContent>`),
            );
            // mc.xsd declares Requires use="required": a Choice without it
            // is malformed and must not win over the well-formed Fallback.
            expect(
                getElementsByTagNameAll(withFallback.document).map((node) => node.localName),
                label,
            ).toEqual(["root", "fallback"]);
            expect(withFallback.issues, label).toContainEqual(expect.objectContaining({ code: "mce-malformed", severity: "error" }));
            const withoutFallback = process(wrap(`<mc:AlternateContent>${choice}</mc:AlternateContent>`));
            expect(withoutFallback.document.documentElement!.childNodes.length, label).toBe(0);
            expect(withoutFallback.issues, label).toContainEqual(expect.objectContaining({ code: "mce-malformed", severity: "error" }));
        }
    });
    it("still selects a later well-formed Choice when an earlier one lacks Requires (#494)", () => {
        const result = process(
            wrap(
                `<mc:AlternateContent><mc:Choice><w:first/></mc:Choice><mc:Choice Requires="w"><w:second/></mc:Choice></mc:AlternateContent>`,
            ),
        );
        expect(getElementsByTagNameAll(result.document).map((node) => node.localName)).toEqual(["root", "second"]);
        expect(result.issues).toContainEqual(expect.objectContaining({ code: "mce-malformed", severity: "error" }));
    });
    it("diagnoses malformed compatibility structures, including unselected choices", () => {
        for (const content of [
            "<mc:Fallback/>",
            "<mc:Choice/>",
            '<mc:Choice Requires=""/>',
            '<mc:Choice Requires="missing"/>',
            '<mc:Choice Requires="w"/><mc:Choice Requires="missing"/>',
            '<mc:Fallback/><mc:Choice Requires="w"/>',
            '<mc:Choice Requires="w"/><mc:Fallback/><mc:Fallback/>',
            '<mc:Choice Requires="w"/><w:unexpected/>',
            '<mc:Choice Requires="w"/><mc:Unexpected/>',
        ]) {
            expect(process(wrap(`<mc:AlternateContent>${content}</mc:AlternateContent>`)).issues, content).toContainEqual(
                expect.objectContaining({ code: "mce-malformed", severity: "error" }),
            );
        }
    });
    it("never mutates the source DOM or unselected branch content", () => {
        const source = parseXml(wrap(alternative("H"), 'mc:Ignorable="wp14"'));
        const before = serializeXml(source);
        const result = applyMarkupCompatibility(source, { understood: UNDERSTOOD_NAMESPACES });
        expect(serializeXml(source)).toBe(before);
        expect(serializeXml(result.document)).not.toContain("AlternateContent");
        expect(serializeXml(source)).toContain("pctPosHOffset");
    });
    it("retains foreign extension payloads even when declared ignorable", () => {
        for (const namespace of ["http://schemas.openxmlformats.org/drawingml/2006/chart", NS.A, NS.P, NS.S]) {
            const { document } = process(
                wrap(
                    `<x:ext xmlns:x="${namespace}" uri="example"><wp14:payload wp14:flag="keep"><wp14:child/></wp14:payload></x:ext>`,
                    'mc:Ignorable="wp14"',
                ),
            );
            expect(getElementsByTagNameAll(document).map((node) => node.localName)).toEqual(["root", "ext", "payload", "child"]);
            expect(serializeXml(document)).toContain('wp14:flag="keep"');
        }
    });
    it("reports, at info level, a root element ignored as a whole (#284)", () => {
        const result = process(
            `<wp14:txbx xmlns:wp14="${WP14}" xmlns:mc="${NS.MC}" xmlns:w="${NS.W}" mc:Ignorable="wp14"><w:p/></wp14:txbx>`,
        );
        expect(result.document.documentElement).toBeNull();
        expect(result.issues).toEqual([
            {
                severity: "info",
                code: "mce-root-ignored",
                message:
                    "Root element wp14:txbx is in an ignorable namespace that is not understood; the part has no content to schema-validate.",
            },
        ]);
    });
    it("retains foreign extension payloads in Strict extension slots", () => {
        for (const namespace of [
            "http://purl.oclc.org/ooxml/drawingml/chart",
            "http://purl.oclc.org/ooxml/drawingml/main",
            "http://purl.oclc.org/ooxml/presentationml/main",
            "http://purl.oclc.org/ooxml/spreadsheetml/main",
        ]) {
            const { document } = applyMarkupCompatibility(
                parseXml(
                    wrap(
                        `<x:ext xmlns:x="${namespace}" uri="example"><wp14:payload wp14:flag="keep"><wp14:child/></wp14:payload></x:ext>`,
                        'mc:Ignorable="wp14"',
                    ),
                ),
                { understood: STRICT_UNDERSTOOD_NAMESPACES },
            );
            expect(getElementsByTagNameAll(document).map((node) => node.localName)).toEqual(["root", "ext", "payload", "child"]);
        }
    });
    it("understands the Strict counterpart of every understood Transitional OOXML namespace", () => {
        const strictTargets = new Set(
            readdirSync(path.join(defaultSchemasDir(), "iso-strict"))
                .filter((name) => name.endsWith(".xsd"))
                .map(
                    (name) =>
                        /targetNamespace="([^"]+)"/.exec(readFileSync(path.join(defaultSchemasDir(), "iso-strict", name), "utf-8"))?.[1],
                ),
        );
        const strictOnly = [...STRICT_UNDERSTOOD_NAMESPACES].filter((uri) => !UNDERSTOOD_NAMESPACES.has(uri));
        const transitionalOoxml = [...UNDERSTOOD_NAMESPACES].filter(
            (uri) => uri.startsWith("http://schemas.openxmlformats.org/") && uri !== NS.MC,
        );
        expect(strictOnly).toHaveLength(transitionalOoxml.length);
        for (const uri of strictOnly) {
            expect(strictTargets.has(uri)).toBe(true);
        }
        for (const uri of UNDERSTOOD_NAMESPACES) {
            expect(STRICT_UNDERSTOOD_NAMESPACES.has(uri)).toBe(true);
        }
        const strictPart = parseXml(
            `<w:document xmlns:w="http://purl.oclc.org/ooxml/wordprocessingml/main" xmlns:mc="${NS.MC}" mc:MustUnderstand="w"/>`,
        );
        expect(applyMarkupCompatibility(strictPart, { understood: STRICT_UNDERSTOOD_NAMESPACES }).issues).toEqual([]);
    });
    it("retains non-ignorable foreign content and understood namespaces", () => {
        // This control needs a truly unknown URI: Office attributes have their
        // own removal rule, independently covered above. Keep the same QNames.
        const source = wrap(
            '<wp14:foreign xmlns:wp14="urn:example:foo" wp14:attribute="kept"/><w:p mc:MustUnderstand="w xml"/>',
            'mc:Ignorable="w"',
        );
        const result = process(source);
        expect(getElementsByTagNameAll(result.document).map((node) => node.localName)).toEqual(["root", "foreign", "p"]);
        expect(serializeXml(result.document)).toContain('wp14:attribute="kept"');
        expect(result.issues).toEqual([]);
    });
    it("resolves ProcessContent QNames where declared and inherits expanded names", () => {
        const result = process(
            wrap(
                `<w:p xmlns:p="${WP14}" mc:ProcessContent="p:keep"><w:r xmlns:p="urn:other"><wp14:keep><w:t>kept</w:t></wp14:keep><wp14:drop><w:t>lost</w:t></wp14:drop></w:r></w:p>`,
                'mc:Ignorable="wp14"',
            ),
        );
        expect(result.document.documentElement.textContent).toBe("kept");
        expect(getElementsByTagNameAll(result.document).some((node) => node.namespaceURI === WP14)).toBe(false);
        expect(result.issues).toEqual([]);
    });
    it("uses additional declared capabilities for standalone extension parts", () => {
        for (const namespace of [NS.W15, NS.W16CEX, NS.W16CID]) {
            const source = parseXml(
                `<part:root xmlns:part="${namespace}" xmlns:mc="${NS.MC}" mc:Ignorable="part"><part:child/></part:root>`,
            );
            const { document, issues } = applyMarkupCompatibility(source, {
                understood: new Set([...UNDERSTOOD_NAMESPACES, namespace]),
            });
            expect(getElementsByTagNameAll(document).map((node) => node.localName)).toEqual(["root", "child"]);
            expect(issues).toEqual([]);
        }
    });
});

describe("markup compatibility: understood extension namespaces (#288)", () => {
    it("keeps the attributes of an understood Office namespace and reports nothing", () => {
        const xml = `<w15:commentsEx xmlns:w15="${NS.W15}"><w15:commentEx w15:paraId="00000001" w15:done="0"/></w15:commentsEx>`;
        const { document, issues } = applyMarkupCompatibility(parseXml(xml), { understood: new Set([...UNDERSTOOD_NAMESPACES, NS.W15]) });
        expect(serializeXml(document)).toBe(serializeXml(parseXml(xml)));
        expect(issues).toEqual([]);
    });
});
