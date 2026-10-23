/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

/**
 * Unit tests for `textboxContentNamespaceCheck` (issue #244, jubarte lane
 * family k12). Three textbox models exist and each pins its content
 * element's namespace: `wp:txbx` takes `wp:txbxContent` (the same
 * wordprocessingDrawing namespace), while `wps:txbx` and the VML
 * `v:textbox` take `w:txbxContent`. A body in any other namespace —
 * Word's Strict-writer `wne:txbxContent` quirk is the observed defect —
 * is an advisory; Word tolerates it.
 */

import { describe, expect, it } from "vitest";

import { ValidationContext } from "../src/lib/validation-context";
import { textboxContentNamespaceCheck } from "../src/scripts/office/checks/textbox-content-namespace";

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const W_STRICT = "http://purl.oclc.org/ooxml/wordprocessingml/main";
const WP = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
const WP_STRICT = "http://purl.oclc.org/ooxml/drawingml/wordprocessingDrawing";
const WPS = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape";
const VML = "urn:schemas-microsoft-com:vml";
const WNE = "http://schemas.microsoft.com/office/word/2006/wordml";

const NS = `xmlns:w="${W}" xmlns:wp="${WP}" xmlns:wps="${WPS}" xmlns:v="${VML}" xmlns:wne="${WNE}"`;
const documentWith = (body: string): string =>
    `<?xml version="1.0" encoding="UTF-8"?><w:document ${NS}><w:body>${body}</w:body></w:document>`;

const run = (entries: Readonly<Record<string, string>>) => textboxContentNamespaceCheck.run(ValidationContext.fromEntries(entries));

describe("textbox-content-namespace check", () => {
    it("reports a wne-namespace text body inside wp:txbx and names the required wp:txbxContent", async () => {
        const xml = documentWith(
            `<wp:wsp><wp:txbx><wne:txbxContent><w:p><w:r><w:t>boxed</w:t></w:r></w:p></wne:txbxContent></wp:txbx></wp:wsp>`,
        );
        const issues = await run({ "word/document.xml": xml });
        expect(issues).toHaveLength(1);
        expect(issues[0]).toMatchObject({ severity: "warning", code: "textbox-content-namespace", path: "word/document.xml" });
        expect(issues[0]!.message).toContain(WNE);
        expect(issues[0]!.message).toContain("txbxContent");
        expect(issues[0]!.message).toContain(WP);
    });

    it("also reports a WML-namespace text body inside wp:txbx", async () => {
        // w:txbxContent is the wps/VML model's element; inside a wp:txbx it
        // does not satisfy CT_TextboxInfo either.
        const xml = documentWith(`<wp:wsp><wp:txbx><w:txbxContent><w:p/></w:txbxContent></wp:txbx></wp:wsp>`);
        const issues = await run({ "word/document.xml": xml });
        expect(issues).toHaveLength(1);
        expect(issues[0]!.message).toContain(W);
        expect(issues[0]!.message).toContain(WP);
    });

    it("accepts the correct body in each of the three textbox models", async () => {
        const xml = documentWith(
            `<wp:wsp><wp:txbx><wp:txbxContent><w:p/></wp:txbxContent></wp:txbx></wp:wsp>` +
                `<wps:wsp><wps:txbx><w:txbxContent><w:p/></w:txbxContent></wps:txbx></wps:wsp>` +
                `<w:p><w:r><w:pict><v:shape><v:textbox><w:txbxContent><w:p/></w:txbxContent></v:textbox></v:shape></w:pict></w:r></w:p>`,
        );
        expect(await run({ "word/document.xml": xml })).toEqual([]);
    });

    it("checks the strict wordprocessingDrawing namespace too", async () => {
        const good =
            `<?xml version="1.0"?><w:document xmlns:w="${W_STRICT}" xmlns:wp="${WP_STRICT}"><w:body>` +
            `<wp:wsp><wp:txbx><wp:txbxContent><w:p/></wp:txbxContent></wp:txbx></wp:wsp></w:body></w:document>`;
        expect(await run({ "word/document.xml": good })).toEqual([]);
        const bad = good
            .replace(/<wp:txbxContent>/, `<wne:txbxContent xmlns:wne="${WNE}">`)
            .replace(/<\/wp:txbxContent>/, "</wne:txbxContent>");
        const issues = await run({ "word/document.xml": bad });
        expect(issues).toHaveLength(1);
        expect(issues[0]!.message).toContain(WP_STRICT);
    });

    it("reports the wps model's body in the wrong namespace and VML's likewise", async () => {
        const xml = documentWith(
            `<wps:wsp><wps:txbx><wps:txbxContent><w:p/></wps:txbxContent></wps:txbx></wps:wsp>` +
                `<w:p><w:r><w:pict><v:shape><v:textbox><v:txbxContent/></v:textbox></v:shape></w:pict></w:r></w:p>`,
        );
        const issues = await run({ "word/document.xml": xml });
        expect(issues).toHaveLength(2);
        expect(issues[0]!.message).toContain(WPS);
        expect(issues[0]!.message).toContain(W);
        expect(issues[1]!.message).toContain(VML);
        expect(issues[1]!.message).toContain(W);
    });

    it("reports each occurrence separately and scans every story part", async () => {
        const site = `<wp:wsp><wp:txbx><wne:txbxContent><w:p/></wne:txbxContent></wp:txbx></wp:wsp>`;
        const issues = await run({
            "word/document.xml": documentWith(site + site),
            "word/header1.xml": `<?xml version="1.0"?><w:hdr ${NS}>${site}</w:hdr>`,
            "word/styles.xml": `<?xml version="1.0"?><w:styles xmlns:w="${W}">${site}</w:styles>`,
        });
        expect(issues).toHaveLength(3);
        expect(issues.map((issue) => issue.path)).toEqual(["word/document.xml", "word/document.xml", "word/header1.xml"]);
    });

    it("skips a part that does not parse instead of throwing", async () => {
        const issues = await run({ "word/document.xml": `<w:document ${NS}><w:body><wp:wsp>` });
        expect(issues).toEqual([]);
    });
});
