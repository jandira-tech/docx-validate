import { describe, expect, it } from "vitest";

import { getElementsByTagNameAll, getElementsByTagNameNSAll, parseXml } from "../src/lib/xml-helpers";

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
