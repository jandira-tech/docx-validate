import { execFile } from "node:child_process";
import { fileURLToPath } from "node:url";
import { promisify } from "node:util";
import { describe, expect, it } from "vitest";
import * as xmlHelpers from "../src/lib/xml-helpers";
import { NS } from "../src/lib/types";

import {
    decodeXmlBytes,
    decodeXmlPart,
    encodeXmlPart,
    getElementsByTagNameAll,
    getElementsByTagNameNSAll,
    parseXml,
    parseXmlBytes,
    prettyXml,
    xmlDeclarationEncoding,
} from "../src/lib/xml-helpers";

const DOC = `<?xml version="1.0"?>
<pkg:document xmlns:pkg="urn:root" xmlns:pkga="urn:a" xmlns:pkgb="urn:b">
    <pkga:body><pkga:p>one</pkga:p><pkga:p>two<pkga:r>run</pkga:r></pkga:p></pkga:body>
    <pkgb:body><pkgb:p>three</pkgb:p></pkgb:body>
</pkg:document>`;

describe("getRelationshipChildren (#434)", () => {
    it("returns only direct Relationship children of a package Relationships root", () => {
        expect(xmlHelpers.getRelationshipChildren).toBeTypeOf("function");
        const dom = parseXml(
            `<p:Relationships xmlns:p="${NS.PR}">\n<!--keep--><?ok yes?><p:Relationship Id="first"/><x:Wrap xmlns:x="urn:x"><p:Relationship Id="nested"/></x:Wrap><Relationship xmlns="urn:x"/><p:Relationship Id="last"/></p:Relationships>`,
        );
        const before = xmlHelpers.serializeXml(dom);
        expect(xmlHelpers.getRelationshipChildren(dom).map((rel) => rel.getAttribute("Id"))).toEqual(["first", "last"]);
        expect(xmlHelpers.serializeXml(dom)).toBe(before);
        for (const xml of [`<Foo xmlns="${NS.PR}"><Relationship/></Foo>`, '<Relationships xmlns="urn:x"/>', "<Relationships/>"]) {
            expect(xmlHelpers.getRelationshipChildren(parseXml(xml))).toEqual([]);
        }
        dom.removeChild(dom.documentElement);
        expect(xmlHelpers.getRelationshipChildren(dom)).toEqual([]);
    });
});

describe("xml-helpers element collectors", () => {
    it("getElementsByTagNameNSAll returns matching descendants in document order", () => {
        const dom = parseXml(DOC);
        const ps = getElementsByTagNameNSAll(dom, "urn:a", "p");
        expect(ps.map((p) => p.textContent)).toEqual(["one", "tworun"]);
    });

    it("root semantics mirror the native API: documentElement included for a Document root, self excluded for an Element root", () => {
        const dom = parseXml(DOC);
        // Document.getElementsByTagName('*') includes the documentElement…
        expect(getElementsByTagNameNSAll(dom, "urn:root", "document")).toHaveLength(1);
        // …but Element.getElementsByTagNameNS never returns the element itself.
        const body = getElementsByTagNameNSAll(dom, "urn:a", "body")[0]!;
        expect(getElementsByTagNameNSAll(body, "urn:a", "body")).toHaveLength(0);
    });

    it("getElementsByTagNameNSAll supports '*' wildcards for namespace and local name", () => {
        const dom = parseXml(DOC);
        // 7 = documentElement + 6 descendants.
        expect(getElementsByTagNameNSAll(dom, "*", "*")).toHaveLength(7);
        expect(getElementsByTagNameNSAll(dom, "*", "p")).toHaveLength(3);
        expect(getElementsByTagNameNSAll(dom, "urn:a", "*")).toHaveLength(4);
    });

    it("getElementsByTagNameAll is exactly the '*','*' form of getElementsByTagNameNSAll", () => {
        const dom = parseXml(DOC);
        expect(getElementsByTagNameAll(dom)).toEqual(getElementsByTagNameNSAll(dom, "*", "*"));
        expect(getElementsByTagNameAll(dom).map((e) => e.nodeName)).toEqual([
            "pkg:document",
            "pkga:body",
            "pkga:p",
            "pkga:p",
            "pkga:r",
            "pkgb:body",
            "pkgb:p",
        ]);
    });

    it("collectors work from a non-document element root and ignore comments/PIs", () => {
        const dom = parseXml(`<?xml version="1.0"?><r xmlns:x="urn:x"><!--c--><?pi?><x:a/><x:b><x:a/></x:b></r>`);
        const aRoot = dom.documentElement!;
        const as = getElementsByTagNameNSAll(aRoot, "urn:x", "a");
        expect(as).toHaveLength(2);
        // 3 = x:a + x:b + nested x:a; the starting element itself is excluded.
        expect(getElementsByTagNameAll(aRoot)).toHaveLength(3);
    });

    it("a Document-rooted query sees every change made to the document since the last query", () => {
        const dom = parseXml(DOC);
        const texts = () => getElementsByTagNameNSAll(dom, "urn:a", "p").map((p) => p.textContent);
        expect(texts()).toEqual(["one", "tworun"]);

        const body = getElementsByTagNameNSAll(dom, "urn:a", "body")[0]!;
        const added = dom.createElementNS("urn:a", "pkga:p");
        added.appendChild(dom.createTextNode("zero"));
        body.insertBefore(added, body.firstChild);
        expect(texts()).toEqual(["zero", "one", "tworun"]);

        body.removeChild(getElementsByTagNameNSAll(dom, "urn:a", "p")[1]!);
        expect(texts()).toEqual(["zero", "tworun"]);

        // A change below another element counts too.
        const run = getElementsByTagNameNSAll(dom, "urn:a", "r")[0]!;
        run.appendChild(dom.createElementNS("urn:a", "pkga:p"));
        expect(getElementsByTagNameNSAll(dom, "urn:a", "p")).toHaveLength(3);
        expect(getElementsByTagNameNSAll(dom, "*", "*")).toHaveLength(8);
    });

    it("a Document-rooted query returns a new array each call, in the order an Element-rooted walk finds", () => {
        const dom = parseXml(DOC);
        const first = getElementsByTagNameNSAll(dom, "*", "*");
        first.length = 0;
        expect(getElementsByTagNameNSAll(dom, "*", "*")).toHaveLength(7);
        getElementsByTagNameNSAll(dom, "urn:a", "p").pop();
        expect(getElementsByTagNameNSAll(dom, "urn:a", "p")).toHaveLength(2);

        const root = dom.documentElement!;
        for (const [ns, local] of [
            ["*", "*"],
            ["*", "p"],
            ["urn:a", "*"],
            ["urn:b", "p"],
            ["urn:a", "missing"],
            ["urn:missing", "p"],
        ] as const) {
            const viaRoot = getElementsByTagNameNSAll(root, ns, local);
            const rootMatches = (ns === "*" || root.namespaceURI === ns) && (local === "*" || root.localName === local);
            expect(getElementsByTagNameNSAll(dom, ns, local)).toEqual(rootMatches ? [root, ...viaRoot] : viaRoot);
        }
    });
});

describe("decodeXmlBytes", () => {
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
            out[offset + i * 2] = endian === "le" ? text.charCodeAt(i) & 0xff : text.charCodeAt(i) >> 8;
            out[offset + i * 2 + 1] = endian === "le" ? text.charCodeAt(i) >> 8 : text.charCodeAt(i) & 0xff;
        }
        return out;
    };

    const utf8 = (text: string): Uint8Array => new TextEncoder().encode(text);
    const BODY = "<a>café — 日本語</a>";

    it("strips a UTF-8 BOM", () => {
        expect(decodeXmlBytes(new Uint8Array([0xef, 0xbb, 0xbf, ...utf8(BODY)]))).toBe(BODY);
    });

    it("decodes UTF-8 without a BOM", () => {
        expect(decodeXmlBytes(utf8(BODY))).toBe(BODY);
    });

    it("honours the UTF-16LE BOM", () => {
        expect(decodeXmlBytes(utf16(BODY, "le", true))).toBe(BODY);
    });

    it("honours the UTF-16BE BOM", () => {
        expect(decodeXmlBytes(utf16(BODY, "be", true))).toBe(BODY);
    });

    it("uses the declared encoding when there is no BOM", () => {
        expect(decodeXmlBytes(utf16(`<?xml version="1.0" encoding="UTF-16LE"?>${BODY}`, "le", false))).toBe(
            `<?xml version="1.0" encoding="UTF-16LE"?>${BODY}`,
        );
        expect(decodeXmlBytes(utf16(`<?xml version="1.0" encoding="UTF-16BE"?>${BODY}`, "be", false))).toBe(
            `<?xml version="1.0" encoding="UTF-16BE"?>${BODY}`,
        );
    });

    it("sniffs the endianness of a BOM-less UTF-16 entity declaring plain UTF-16", () => {
        expect(decodeXmlBytes(utf16(`<?xml version="1.0" encoding="UTF-16"?>${BODY}`, "le", false))).toBe(
            `<?xml version="1.0" encoding="UTF-16"?>${BODY}`,
        );
        expect(decodeXmlBytes(utf16(`<?xml version="1.0" encoding="UTF-16"?>${BODY}`, "be", false))).toBe(
            `<?xml version="1.0" encoding="UTF-16"?>${BODY}`,
        );
    });

    it("reads a BOM-less UTF-16 declaration of any length, not just the first 200 bytes", () => {
        // The declaration may carry arbitrary whitespace (XML 1.0 §2.8, S),
        // so a fixed byte window can cut the encoding name off.
        const decl = `<?xml version="1.0"${" ".repeat(70)}encoding="UTF-16"?>`;
        expect(decodeXmlBytes(utf16(`${decl}${BODY}`, "le", false))).toBe(`${decl}${BODY}`);
        expect(decodeXmlBytes(utf16(`${decl}${BODY}`, "be", false))).toBe(`${decl}${BODY}`);
    });

    it("keeps surrogate pairs and NUL-high-byte characters after a BOM-less UTF-16 declaration intact", () => {
        const text = `<?xml version="1.0" encoding="UTF-16"?><r a="\u{1F600}">\u0100\u4E2D\u{1F600}</r>`;
        expect(decodeXmlBytes(utf16(text, "le", false))).toBe(text);
        expect(decodeXmlBytes(utf16(text, "be", false))).toBe(text);
    });

    it("prefers the BOM over a contradicting declaration", () => {
        expect(decodeXmlBytes(utf16('<?xml version="1.0" encoding="UTF-16BE"?>' + BODY, "le", true))).toBe(
            `<?xml version="1.0" encoding="UTF-16BE"?>${BODY}`,
        );
    });

    it("falls back to UTF-8 for other declarations and for no declaration at all", () => {
        expect(decodeXmlBytes(utf8('<?xml version="1.0" encoding="US-ASCII"?><a>ok</a>'))).toBe(
            '<?xml version="1.0" encoding="US-ASCII"?><a>ok</a>',
        );
        expect(decodeXmlBytes(utf8("<a>ok</a>"))).toBe("<a>ok</a>");
        expect(decodeXmlBytes(utf8("<a>ok</a>"))).not.toBe("");
    });

    it("returns an empty string for empty input", () => {
        expect(decodeXmlBytes(new Uint8Array())).toBe("");
    });

    it("parseXmlBytes builds the DOM from the decoded bytes", () => {
        const dom = parseXmlBytes(new Uint8Array([0xef, 0xbb, 0xbf, ...utf8(BODY)]));
        expect(dom.documentElement?.textContent).toBe("café — 日本語");
    });
});

describe("decodeXmlPart / encodeXmlPart", () => {
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
            out[offset + i * 2] = endian === "le" ? text.charCodeAt(i) & 0xff : text.charCodeAt(i) >> 8;
            out[offset + i * 2 + 1] = endian === "le" ? text.charCodeAt(i) >> 8 : text.charCodeAt(i) & 0xff;
        }
        return out;
    };

    const utf8 = (text: string): Uint8Array => new TextEncoder().encode(text);
    const BODY = "<a>café — 日本語</a>";

    it("reports the encoding and the byte order mark of a UTF-8 part", () => {
        expect(decodeXmlPart(utf8(BODY))).toEqual({ text: BODY, encoding: "utf-8", bom: undefined });
        expect(decodeXmlPart(new Uint8Array([0xef, 0xbb, 0xbf, ...utf8(BODY)]))).toEqual({
            text: BODY,
            encoding: "utf-8",
            bom: "utf-8",
        });
    });

    it("reports the encoding and the byte order mark of a UTF-16 part", () => {
        expect(decodeXmlPart(utf16(BODY, "le", true))).toEqual({ text: BODY, encoding: "utf-16le", bom: "utf-16le" });
        expect(decodeXmlPart(utf16(BODY, "be", true))).toEqual({ text: BODY, encoding: "utf-16be", bom: "utf-16be" });
    });

    it("reports the encoding of a BOM-less UTF-16 part the declaration names", () => {
        const declared = `<?xml version="1.0" encoding="UTF-16"?>${BODY}`;
        expect(decodeXmlPart(utf16(declared, "le", false))).toEqual({ text: declared, encoding: "utf-16le", bom: undefined });
        expect(decodeXmlPart(utf16(declared, "be", false))).toEqual({ text: declared, encoding: "utf-16be", bom: undefined });
    });

    it("refuses an encoding it cannot decode", () => {
        expect(decodeXmlPart(utf8('<?xml version="1.0" encoding="macintosh"?><a>ok</a>'))).toBeUndefined();
        expect(decodeXmlPart(utf8('<?xml version="1.0" encoding="windows-1252"?><a>ok</a>'))).toBeUndefined();
        expect(decodeXmlPart(utf8('<?xml version="1.0" encoding="ISO-8859-1"?><a>ok</a>'))).toBeUndefined();
    });

    it("decodes US-ASCII, whose every byte UTF-8 encodes the same way", () => {
        const ascii = '<?xml version="1.0" encoding="US-ASCII"?><a>ok</a>';
        expect(decodeXmlPart(utf8(ascii))).toEqual({ text: ascii, encoding: "utf-8", bom: undefined });
    });

    it("keeps decodeXmlBytes' UTF-8 fallback for an encoding it cannot decode", () => {
        const macintosh = '<?xml version="1.0" encoding="macintosh"?><a>ok</a>';
        expect(decodeXmlBytes(utf8(macintosh))).toBe(macintosh);
    });

    it("writes a part back in its own encoding, byte order mark included", () => {
        for (const [encoding, bom, bytes] of [
            ["utf-8", undefined, utf8(BODY)],
            ["utf-16le", "utf-16le", utf16(BODY, "le", true)],
            ["utf-16be", "utf-16be", utf16(BODY, "be", true)],
        ] as const) {
            const part = decodeXmlPart(bytes)!;
            expect(part.encoding).toBe(encoding);
            expect(part.bom).toBe(bom);
            expect(encodeXmlPart(part.text, part)).toEqual(bytes);
        }
    });

    it("writes a UTF-8 part without a byte order mark, as before #866", () => {
        const part = decodeXmlPart(new Uint8Array([0xef, 0xbb, 0xbf, ...utf8(BODY)]))!;
        expect(part).toEqual({ text: BODY, encoding: "utf-8", bom: "utf-8" });
        expect(encodeXmlPart(part.text, part)).toEqual(utf8(BODY));
    });

    it("round-trips text through a UTF-16 part without losing a character", () => {
        const part = decodeXmlPart(utf16(BODY, "be", true))!;
        const written = encodeXmlPart(part.text, part);
        expect(decodeXmlPart(written)).toEqual(part);
        expect([...written.subarray(0, 2)]).toEqual([0xfe, 0xff]);
    });

    it("names the encoding a part's declaration must carry", () => {
        expect(xmlDeclarationEncoding({ encoding: "utf-8" })).toBe("UTF-8");
        expect(xmlDeclarationEncoding({ encoding: "utf-8", bom: "utf-8" })).toBe("UTF-8");
        expect(xmlDeclarationEncoding({ encoding: "utf-16le", bom: "utf-16le" })).toBe("UTF-16");
        expect(xmlDeclarationEncoding({ encoding: "utf-16be", bom: "utf-16be" })).toBe("UTF-16");
    });

    it("names the byte order of a UTF-16 part with no byte order mark (#878)", () => {
        expect(xmlDeclarationEncoding({ encoding: "utf-16le" })).toBe("UTF-16LE");
        expect(xmlDeclarationEncoding({ encoding: "utf-16be" })).toBe("UTF-16BE");
    });
});

describe("prettyXml", () => {
    it("preserves inline text after opening and empty tags", () => {
        expect(prettyXml(parseXml("<r><a>text</a><empty/>tail</r>"))).toBe("<r>\n  <a>text\n  </a>\n  <empty/>tail\n</r>");
    });

    it("preserves formatting around comments, processing instructions, and CDATA", () => {
        expect(prettyXml(parseXml("<r><!-- <a -->tail<?pi?><![CDATA[<a>]]>text</r>"))).toBe(
            "<r>\n  <!-- <a -->tail\n  <?pi ?>\n  <![CDATA[<a>]]>text\n</r>",
        );
    });

    it("handles repeated tag starts in a comment within a bounded child process", async () => {
        const moduleUrl = new URL("../src/lib/xml-helpers.ts", import.meta.url).href;
        const script = `
            import { strict as assert } from "node:assert";
            import { parseXml, prettyXml } from ${JSON.stringify(moduleUrl)};
            const comment = "<!-- " + "< ".repeat(200_000) + "> trailing -->";
            const xml = "<root>" + comment + "tail</root>";
            assert.equal(prettyXml(parseXml(xml)), "<root>\\n  " + comment + "\\n  tail\\n</root>");
        `;
        await promisify(execFile)(process.execPath, ["--import", "tsx", "--input-type=module", "--eval", script], {
            cwd: fileURLToPath(new URL("..", import.meta.url)),
            timeout: 5_000,
        });
    }, 10_000);
});
