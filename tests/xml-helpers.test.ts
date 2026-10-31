import { execFile } from "node:child_process";
import { fileURLToPath } from "node:url";
import { promisify } from "node:util";
import { describe, expect, it, vi } from "vitest";
import * as xmlHelpers from "../src/lib/xml-helpers";
import { NS } from "../src/lib/types";

import {
    decodeXmlBytes,
    decodeXmlPart,
    encodeXmlPart,
    getElementsByTagNameAll,
    getElementsByTagNameNSAll,
    getLineNumber,
    getLocalName,
    makeSelect,
    parseXml,
    parseXmlBytes,
    prettyXml,
    serializeXml,
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

    it("walks when the document's mutation counter is not a number", () => {
        const dom = parseXml(DOC);
        // A stable non-number: xmldom still updates the tree, and two queries would otherwise cache the old one.
        Object.defineProperty(dom, "_inc", {
            get: () => "opaque",
            set: () => undefined,
            configurable: true,
        });
        const texts = () => getElementsByTagNameNSAll(dom, "urn:a", "p").map((p) => p.textContent);
        expect(texts()).toEqual(["one", "tworun"]);
        expect(texts()).toEqual(["one", "tworun"]);
        const body = dom.getElementsByTagName("pkga:body")[0]!;
        const added = dom.createElementNS("urn:a", "pkga:p");
        added.appendChild(dom.createTextNode("three"));
        body.appendChild(added);
        expect(texts()).toEqual(["one", "tworun", "three"]);
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

    it("refuses an encoding it cannot decode however long the declaration is (#881)", () => {
        const padding = " ".repeat(240);
        const macintosh = `<?xml version="1.0"${padding} encoding="macintosh"?><a>ok</a>`;
        expect(decodeXmlPart(utf8(macintosh))).toBeUndefined();
        expect(decodeXmlPart(utf8(`<?xml version="1.0"${padding} encoding="windows-1252"?><a>ok</a>`))).toBeUndefined();
        expect(decodeXmlBytes(utf8(macintosh))).toBe(macintosh);
        const utf8Long = `<?xml version="1.0"${padding} encoding="UTF-8"?><a>ok</a>`;
        expect(decodeXmlPart(utf8(utf8Long))).toEqual({ text: utf8Long, encoding: "utf-8", bom: undefined });
    });

    it("decodes a part that opens with no declaration only once (#888)", () => {
        const decode = vi.spyOn(TextDecoder.prototype, "decode");
        try {
            const commented = `<!-- ${"> ".repeat(64)}-->\n<a>ok</a>`;
            const unmarked = "plain text, no markup";
            expect(decodeXmlPart(utf8(commented))).toEqual({ text: commented, encoding: "utf-8", bom: undefined });
            expect(decodeXmlPart(utf8(unmarked))).toEqual({ text: unmarked, encoding: "utf-8", bom: undefined });
            expect(decode.mock.contexts.map((decoder) => (decoder as TextDecoder).encoding)).toEqual(["utf-8", "utf-8"]);
            // A part that opens with a declaration still has it read first.
            decode.mockClear();
            decodeXmlPart(utf8('<?xml version="1.0"?><a>ok</a>'));
            expect(decode.mock.contexts.map((decoder) => (decoder as TextDecoder).encoding)).toEqual(["windows-1252", "utf-8"]);
        } finally {
            decode.mockRestore();
        }
    });

    it("reads a declaration after the leading whitespace its pattern accepts, as before #888", () => {
        // `latin1` decodes as windows-1252, where tab, line feed, vertical
        // tab, form feed, carriage return, space and 0xA0 are what `\s`
        // matches; 0x85 is an ellipsis there, and 0xC2 0xA0 starts "Â".
        const macintosh = utf8('<?XML version="1.0" encoding="macintosh"?><a>ok</a>');
        for (const lead of [[0x09, 0x0a, 0x0d, 0x20], [0x0b, 0x0c], [0xa0]]) {
            expect(decodeXmlPart(new Uint8Array([...lead, ...macintosh]))).toBeUndefined();
        }
        for (const lead of [[0x85], [0xc2, 0xa0]]) {
            expect(decodeXmlPart(new Uint8Array([...lead, ...macintosh]))?.encoding).toBe("utf-8");
        }
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

    it("writes a BOM-less UTF-16 part without a byte order mark", () => {
        const bytes = encodeXmlPart("<a/>", { encoding: "utf-16le" });
        expect([...bytes.subarray(0, 4)]).toEqual([0x3c, 0x00, 0x61, 0x00]);
    });

    it("reads a declaration that never closes as the whole head", () => {
        const open = `<?xml version="1.0" encoding="UTF-8"`;
        expect(decodeXmlPart(utf8(open))).toEqual({ text: open, encoding: "utf-8" });
        expect(decodeXmlPart(utf8(`<?xml version="1.0" encoding="macintosh"`))).toBeUndefined();
    });
});

describe("prettyXml", () => {
    /** Text nodes in document order, including whitespace-only nodes. */
    const textNodes = (node: Node): string[] => {
        if (node.nodeType === 3 || node.nodeType === 4) return [node.nodeValue ?? ""];
        return Array.from(node.childNodes).flatMap(textNodes);
    };

    it.each([" <!--c--> ", " <x/> ", "<!--c-->", "<?mark a > b?>"])(
        "preserves text nodes beside markup in xml:space=preserve (#1219): %s",
        (content) => {
            const xml = `<w:t xmlns:w="urn:word" xml:space="preserve">${content}</w:t>`;
            const before = parseXml(xml);
            const pretty = prettyXml(before);
            expect(textNodes(parseXml(pretty).documentElement)).toEqual(textNodes(before.documentElement));
            expect(pretty).toBe(xml);
        },
    );

    it.each(["t", "w:t", "a:t", "other:delText", "instrText", "other:delInstrText"])(
        "preserves text-bearing local names without xml:space (#1219): %s",
        (name) => {
            const xml = `<${name} xmlns:w="urn:word" xmlns:a="urn:drawing" xmlns:other="urn:other"> <!--c--><?mark a > b?><x/> </${name}>`;
            const before = parseXml(xml);
            const pretty = prettyXml(before);
            expect(textNodes(parseXml(pretty).documentElement)).toEqual(textNodes(before.documentElement));
            expect(pretty).toBe(xml);
        },
    );

    it.each(["\u00a0<!--c-->", "<!--c-->\u00a0", "&#160;<!--c-->", "\u2003<x/>", " \t\n\r"])(
        "uses XML whitespace when preserving direct text (#1219): %j",
        (content) => {
            const before = parseXml(`<r>${content}</r>`);
            const pretty = prettyXml(before);
            expect(textNodes(parseXml(pretty).documentElement)).toEqual(textNodes(before.documentElement));
            expect(pretty).toBe(xmlHelpers.serializeXml(before));
        },
    );

    it("inherits xml:space when printing a subtree (#1219)", () => {
        const before = parseXml('<r xml:space="preserve"><wrap><target> <!--c--><x/> </target></wrap></r>');
        const target = getElementsByTagNameNSAll(before, "*", "target")[0]!;
        const pretty = prettyXml(target);
        expect(textNodes(parseXml(pretty).documentElement)).toEqual(textNodes(target));
        expect(pretty).toBe("<target> <!--c--><x/> </target>");
    });

    it("resets inherited xml:space at the nearer default ancestor (#1219)", () => {
        const before = parseXml(
            '<r xml:space="preserve"><wrap xml:space="default"><target> <a/><!--c--><?mark a > b?><b/> </target></wrap></r>',
        );
        const target = getElementsByTagNameNSAll(before, "*", "target")[0]!;
        expect(prettyXml(target)).toBe("<target>\n  <a/>\n  <!--c-->\n  <?mark a > b?>\n  <b/>\n</target>");
    });

    it("keeps a preserved container and its descendants verbatim (#1219)", () => {
        const xml = '<r xml:space="preserve"> <wrap> <!--c--><x/> </wrap> </r>';
        const before = parseXml(xml);
        const pretty = prettyXml(before);
        expect(textNodes(parseXml(pretty).documentElement)).toEqual(textNodes(before.documentElement));
        expect(pretty).toBe(xml);
    });

    it.each([
        "<!DOCTYPE r [<!ELEMENT r ANY>]>",
        '<!DOCTYPE r [<!ELEMENT r ANY><!ENTITY e "a > b [ ]">]>',
        '<!DOCTYPE r SYSTEM "a>b.dtd" [<!-- ] > [ --><?mark ] >?><!ELEMENT r ANY>]>',
    ])("keeps the complete DOCTYPE subset opaque (#1219): %s", (doctype) => {
        const before = parseXml(`${doctype}<r/>`);
        const pretty = prettyXml(before);
        const after = parseXml(pretty);
        expect(xmlHelpers.serializeXml(after.doctype!)).toBe(xmlHelpers.serializeXml(before.doctype!));
        expect(textNodes(after.documentElement)).toEqual(textNodes(before.documentElement));
        expect(pretty).toBe(`${xmlHelpers.serializeXml(before.doctype!)}\n<r/>`);
    });

    it("indents element-only Word containers with comments and PIs (#1219)", () => {
        const xml = '<w:body xmlns:w="urn:word"> <w:p> <!--c--><?mark a > b?><w:r/> </w:p> </w:body>';
        expect(prettyXml(parseXml(xml))).toBe(
            '<w:body xmlns:w="urn:word">\n  <w:p>\n    <!--c-->\n    <?mark a > b?>\n    <w:r/>\n  </w:p>\n</w:body>',
        );
    });

    it.each([" a  b ", "\t a\n  b\t "])("keeps preserve-space text verbatim (#1219): %j", (text) => {
        const xml = `<w:t xmlns:w="urn:word" xml:space="preserve">${text}</w:t>`;
        expect(prettyXml(parseXml(xml))).toBe(xml);
    });

    it("keeps plain text and its closing tag together (#1219)", () => {
        expect(prettyXml(parseXml("<r><t>Hello</t></r>"))).toBe("<r>\n  <t>Hello</t>\n</r>");
    });

    it.each([" ", "  ", "\t\n  "])("keeps a whitespace-only text child (#1219): %j", (text) => {
        const xml = `<t xml:space="preserve">${text}</t>`;
        expect(prettyXml(parseXml(xml))).toBe(xml);
    });

    it.each([
        "<r> before  <b> bold </b>  after </r>",
        "<r><b/>tail</r>",
        "<r>head<b/></r>",
        "<r> head <b><c/></b> tail </r>",
        "<r>before<b/> <c/>after</r>",
    ])("adds no whitespace inside mixed content (#1219): %s", (xml) => {
        expect(prettyXml(parseXml(xml))).toBe(xml);
    });

    it.each(["<r><![CDATA[ a  > b <c/>\n ]]></r>", "<r> before <![CDATA[<a>]]> after </r>"])(
        "keeps CDATA byte-identical (#1219): %s",
        (xml) => {
            expect(prettyXml(parseXml(xml))).toBe(xml);
        },
    );

    it("keeps comments, PIs, and entity escapes byte-identical in mixed content (#1219)", () => {
        const xml = "<r> &amp; &lt; <!-- a > b <c --><?pi a > b?> &gt; </r>";
        expect(prettyXml(parseXml(xml))).toBe(xml);
    });

    it("keeps nested element-only indentation and drops old indentation (#1219)", () => {
        expect(prettyXml(parseXml("<r> \n <a>\n <b/> <c/>\n </a>\t<d/>\n</r>"), "    ")).toBe(
            "<r>\n    <a>\n        <b/>\n        <c/>\n    </a>\n    <d/>\n</r>",
        );
    });

    it("preserves inline text after opening and empty tags", () => {
        expect(prettyXml(parseXml("<r><a>text</a><empty/>tail</r>"))).toBe("<r><a>text</a><empty/>tail</r>");
    });

    it("preserves formatting around comments, processing instructions, and CDATA", () => {
        expect(prettyXml(parseXml("<r><!-- <a -->tail<?pi?><![CDATA[<a>]]>text</r>"))).toBe(
            "<r><!-- <a -->tail<?pi ?><![CDATA[<a>]]>text</r>",
        );
    });

    it("handles repeated tag starts in a comment within a bounded child process", async () => {
        const moduleUrl = new URL("../src/lib/xml-helpers.ts", import.meta.url).href;
        const script = `
            import { strict as assert } from "node:assert";
            import { parseXml, prettyXml } from ${JSON.stringify(moduleUrl)};
            const comment = "<!-- " + "< ".repeat(200_000) + "> trailing -->";
            const xml = "<root>" + comment + "tail</root>";
            assert.equal(prettyXml(parseXml(xml)), xml);
        `;
        await promisify(execFile)(process.execPath, ["--import", "tsx", "--input-type=module", "--eval", script], {
            cwd: fileURLToPath(new URL("..", import.meta.url)),
            timeout: 5_000,
        });
    }, 10_000);

    it("indents element-only content", () => {
        // #1219: prettyXml inserts whitespace into element text (`x` becomes `x\n  `). Pin only structure.
        expect(prettyXml(parseXml("<r><a><b/></a></r>"))).toBe("<r>\n  <a>\n    <b/>\n  </a>\n</r>");
    });

    it("copies mixed content with a comment holding angle brackets verbatim (#1219)", () => {
        expect(prettyXml(parseXml("<r><!--<>-->tail</r>"))).toBe("<r><!--<>-->tail</r>");
        expect(prettyXml(parseXml("<r><!-- > -->tail</r>"))).toBe("<r><!-- > -->tail</r>");
    });

    it("drops whitespace-only text", () => {
        expect(prettyXml(parseXml("<r> <a/></r>"))).toBe("<r>\n  <a/>\n</r>");
    });

    it("pretty-prints a document with no root as an empty string", () => {
        const dom = parseXml("<r/>");
        dom.removeChild(dom.documentElement);
        expect(prettyXml(dom)).toBe("");
    });
});

describe("parseXml", () => {
    it("strips a leading UTF-8 BOM before parsing", () => {
        expect(parseXml("\uFEFF<a/>").documentElement.localName).toBe("a");
    });

    it("wraps a recoverable parse error and keeps the parser's message", () => {
        expect(() => parseXml("\u0000<a/>")).toThrow(/^parseXml failed: Unexpected content outside root element: '\0'$/);
    });
});

describe("serializeXml", () => {
    it("names the encoding it is asked to, adding a declaration or rewriting the one the document carries", () => {
        expect(serializeXml(parseXml("<r/>"), "UTF-8")).toBe('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<r/>');
        expect(serializeXml(parseXml('<?xml version="1.0" encoding="UTF-16"?><r/>'), "UTF-8")).toBe(
            '<?xml version="1.0" encoding="UTF-8"?><r/>',
        );
        const utf16le = serializeXml(parseXml("<r/>"), "UTF-16LE");
        expect(utf16le).toBe('<?xml version="1.0" encoding="UTF-16LE" standalone="yes"?>\n<r/>');
        const utf16leBytes = Buffer.from(utf16le, "utf16le");
        expect(utf16leBytes.subarray(0, 2)).toEqual(Buffer.from([0x3c, 0x00]));
        expect(utf16leBytes.includes(Buffer.from('encoding="UTF-16LE"', "utf16le"))).toBe(true);
        expect(serializeXml(parseXml('<?xml version="1.0" encoding="UTF-8"?><r/>'), "UTF-16LE")).toBe(
            '<?xml version="1.0" encoding="UTF-16LE"?><r/>',
        );
    });
});

describe("makeSelect", () => {
    it("selects with the OOXML prefixes, and with a caller's extra prefix", () => {
        const dom = parseXml(`<w:document xmlns:w="${NS.W}"><w:body><w:p/></w:body></w:document>`);
        const select = makeSelect();
        expect(select("//w:p", dom)).toHaveLength(1);
        const extra = makeSelect({ z: "urn:z" });
        const zdom = parseXml(`<z:r xmlns:z="urn:z"><z:t/></z:r>`);
        expect(extra("//z:t", zdom)).toHaveLength(1);
        expect(extra("//w:p", dom)).toHaveLength(1);
    });
});

describe("getLineNumber", () => {
    it("returns the parser's line, or 0 when none was recorded", () => {
        expect(getLineNumber(null)).toBe(0);
        expect(getLineNumber(undefined)).toBe(0);
        const dom = parseXml("<r>\n<a/>\n</r>");
        expect(getLineNumber(dom.documentElement)).toBe(1);
        expect(getLineNumber(dom.getElementsByTagName("a")[0]!)).toBe(2);
        expect(getLineNumber(dom.createTextNode("t"))).toBe(0);
    });
});

describe("getLocalName", () => {
    it("returns the element's local name, lower-cased when asked", () => {
        const dom = parseXml(`<W:Document xmlns:W="${NS.W}"/>`);
        expect(getLocalName(dom.documentElement)).toBe("Document");
        expect(getLocalName(dom.documentElement, true)).toBe("document");
        expect(getLocalName(dom.documentElement, false)).toBe("Document");
    });
});
