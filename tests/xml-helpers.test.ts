import { execFile } from "node:child_process";
import { fileURLToPath } from "node:url";
import { promisify } from "node:util";
import { describe, expect, it } from "vitest";

import {
    decodeXmlBytes,
    getElementsByTagNameAll,
    getElementsByTagNameNSAll,
    parseXml,
    parseXmlBytes,
    prettyXml,
} from "../src/lib/xml-helpers";

const DOC = `<?xml version="1.0"?>
<pkg:document xmlns:pkg="urn:root" xmlns:pkga="urn:a" xmlns:pkgb="urn:b">
    <pkga:body><pkga:p>one</pkga:p><pkga:p>two<pkga:r>run</pkga:r></pkga:p></pkga:body>
    <pkgb:body><pkgb:p>three</pkgb:p></pkgb:body>
</pkg:document>`;

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
