/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

import { describe, expect, it } from "vitest";
import { ValidationContext } from "../src/lib/validation-context";
import { document, rels, relation, R } from "./support/word-rule-test-xml";
import { wordAttachedToolbarUtf8Check } from "../src/scripts/office/checks/word-attached-toolbar";

describe("word-attached-toolbar-utf8", () => {
    it("reports the Word-rejected shape with unchanged issue details", async () => {
        const ctx = ValidationContext.fromEntries({
            "word/fontTable.xml": document(`<wne:toolbarData r:id="rId1"/>`),
            "word/_rels/fontTable.xml.rels": rels(relation("rId1", "attachedToolbars", "", "attachedToolbars.bin")),
            "word/attachedToolbars.bin": new Uint8Array([0x00, 0xef, 0xbf, 0xbd]),
        });
        expect(await wordAttachedToolbarUtf8Check.run(ctx)).toEqual([
            {
                severity: "error",
                code: "word-attached-toolbar-utf8",
                path: "word/attachedToolbars.bin",
                message:
                    "word/attachedToolbars.bin is referenced by wne:toolbarData and contains the UTF-8 replacement sequence EF BF BD. Microsoft Word offers to recover the file. One sequence is enough. Replace each sequence with the original byte.",
            },
        ]);
    });
    it("accepts the repaired control", async () => {
        const ctx = ValidationContext.fromEntries({
            "word/fontTable.xml": document(`<wne:toolbarData r:id="rId1"/>`),
            "word/_rels/fontTable.xml.rels": rels(relation("rId1", "attachedToolbars", "", "attachedToolbars.bin")),
            "word/attachedToolbars.bin": new Uint8Array([0x00, 0x10]),
            "word/unreferenced.bin": new Uint8Array([0xef, 0xbf, 0xbd]),
        });
        expect(await wordAttachedToolbarUtf8Check.run(ctx)).toEqual([]);
    });
    it("skips absent, irrelevant and unparseable parts", async () => {
        expect(await wordAttachedToolbarUtf8Check.run(ValidationContext.fromEntries({}))).toEqual([]);
        expect(
            await wordAttachedToolbarUtf8Check.run(
                ValidationContext.fromEntries({ "word/fontTable.xml": "<broken", "customXml/item.xml": document("") }),
            ),
        ).toEqual([]);
    });
});
it.each(["/word/attachedToolbars.bin#toolbar", "attachedToolbars.bin"])(
    "resolves a referenced target %s and strict relationship ids",
    async (target) => {
        const ctx = ValidationContext.fromEntries({
            "word/fontTable.xml": document(`<wne:toolbarData rs:id="rId1"/>`),
            "word/_rels/fontTable.xml.rels": rels(relation("rId1", "attachedToolbars", "", target)),
            "word/attachedToolbars.bin": new Uint8Array([0xef, 0xbf, 0xbd]),
        });
        expect(await wordAttachedToolbarUtf8Check.run(ctx)).toHaveLength(1);
    },
);
it.each(["../../escape.bin", "#fragment", "missing.bin"])("leaves unavailable or escaping target %s alone", async (target) => {
    const ctx = ValidationContext.fromEntries({
        "word/fontTable.xml": document(`<wne:toolbarData r:id="rId1"/>`),
        "word/_rels/fontTable.xml.rels": rels(relation("rId1", "attachedToolbars", "", target)),
    });
    expect(await wordAttachedToolbarUtf8Check.run(ctx)).toEqual([]);
});
it("skips absent ids, unrelated types, absent targets, and unparseable sidecars", async () => {
    for (const content of [
        undefined,
        "<Relationships",
        rels(relation("rId1", "image")),
        rels(`<Relationship Id="rId1" Type="${R}/attachedToolbars"/>`),
    ]) {
        const ctx = ValidationContext.fromEntries({
            "word/fontTable.xml": document(`<wne:toolbarData/><wne:toolbarData r:id="rId1"/>`),
            ...(content === undefined ? {} : { "word/_rels/fontTable.xml.rels": content }),
        });
        expect(await wordAttachedToolbarUtf8Check.run(ctx)).toEqual([]);
    }
});
