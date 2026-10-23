/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

import { describe, expect, it } from "vitest";
import { ValidationContext } from "../src/lib/validation-context";
import { document, rels, relation } from "./support/word-rule-test-xml";
import { wordHyperlinkRelationshipCheck } from "../src/scripts/office/checks/word-hyperlink";

describe("word-hyperlink-relationship", () => {
    it("reports the Word-rejected shape with unchanged issue details", async () => {
        const ctx = ValidationContext.fromEntries({ "word/document.xml": document('<w:p><w:hyperlink r:id="rId1"/></w:p>') });
        expect(await wordHyperlinkRelationshipCheck.run(ctx)).toEqual([
            {
                severity: "error",
                code: "word-hyperlink-relationship",
                path: "word/document.xml",
                message:
                    "<w:hyperlink> r:id=\"rId1\" does not reference an external hyperlink relationship (no Relationship with Id 'rId1'). Microsoft Word offers to recover the file. Use a Relationship whose Type ends in /hyperlink and whose TargetMode is External, or remove the r:id.",
            },
        ]);
    });
    it("accepts the repaired control", async () => {
        const ctx = ValidationContext.fromEntries({
            "word/document.xml": document(`<w:p><w:hyperlink r:id="rId1"/><w:hyperlink/></w:p>`),
            "word/_rels/document.xml.rels": rels(relation("rId1", "hyperlink", "External")),
        });
        expect(await wordHyperlinkRelationshipCheck.run(ctx)).toEqual([]);
    });
    it("skips absent, irrelevant and unparseable parts", async () => {
        expect(await wordHyperlinkRelationshipCheck.run(ValidationContext.fromEntries({}))).toEqual([]);
        expect(
            await wordHyperlinkRelationshipCheck.run(
                ValidationContext.fromEntries({ "word/document.xml": "<broken", "customXml/item.xml": document("") }),
            ),
        ).toEqual([]);
    });
});
it.each([
    ["image", "External", "Type"],
    ["hyperlink", "", "TargetMode"],
])("explains an incompatible relationship %s/%s", async (type, mode, reason) => {
    const ctx = ValidationContext.fromEntries({
        "word/document.xml": document(`<w:hyperlink r:id="rId1"/>`),
        "word/_rels/document.xml.rels": rels(relation("rId1", type, mode)),
    });
    expect((await wordHyperlinkRelationshipCheck.run(ctx))[0]?.message).toContain(reason);
});
it("reads strict relationship ids, missing attributes, and unparseable sidecars", async () => {
    const ctx = ValidationContext.fromEntries({
        "word/document.xml": document(`<w:hyperlink rs:id="rId1"/><w:hyperlink/>`),
        "word/_rels/document.xml.rels": rels(`<Relationship/><Relationship Id="rId1"/>`),
    });
    expect((await wordHyperlinkRelationshipCheck.run(ctx))[0]?.message).toContain("Type '(none)'");
    const invalid = ValidationContext.fromEntries({
        "word/document.xml": document(`<w:hyperlink r:id="rId1"/>`),
        "word/_rels/document.xml.rels": "<Relationships",
    });
    expect(await wordHyperlinkRelationshipCheck.run(invalid)).toEqual([]);
});
