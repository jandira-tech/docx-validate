/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

import { describe, expect, it, vi } from "vitest";
import { ValidationContext } from "../src/lib/validation-context";
import { document, rels, relation, R } from "./support/word-rule-test-xml";
import { wordAttachedToolbarUtf8Check } from "../src/scripts/office/checks/word-attached-toolbar";

describe("word-attached-toolbar-utf8", () => {
    it("does not discover toolbar bytes through a nested relationship (#434)", async () => {
        const ctx = ValidationContext.fromEntries({
            "word/fontTable.xml": document('<wne:toolbarData r:id="nested"/>'),
            "word/_rels/fontTable.xml.rels": rels(
                `<x:Wrap xmlns:x="urn:x">${relation("nested", "attachedToolbars", "", "attachedToolbars.bin")}</x:Wrap>`,
            ),
            "word/attachedToolbars.bin": new Uint8Array([0xef, 0xbf, 0xbd]),
        });
        expect(await wordAttachedToolbarUtf8Check.run(ctx)).toEqual([]);
    });

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
    it("reports every referenced part that carries the sequence, once each (Jules on #323)", async () => {
        const ctx = ValidationContext.fromEntries({
            "word/fontTable.xml": document(`<wne:toolbarData r:id="rId1"/><wne:toolbarData r:id="rId2"/><wne:toolbarData r:id="rId3"/>`),
            "word/_rels/fontTable.xml.rels": rels(
                relation("rId1", "attachedToolbars", "", "attachedToolbars.bin") +
                    relation("rId2", "attachedToolbars", "", "attachedToolbars2.bin") +
                    relation("rId3", "attachedToolbars", "", "attachedToolbars.bin"),
            ),
            // A second part naming the first toolbar again adds no finding.
            "word/settings.xml": document(`<wne:toolbarData r:id="rId1"/>`),
            "word/_rels/settings.xml.rels": rels(relation("rId1", "attachedToolbars", "", "attachedToolbars.bin")),
            "word/attachedToolbars.bin": new Uint8Array([0xef, 0xbf, 0xbd]),
            "word/attachedToolbars2.bin": new Uint8Array([0x00, 0xef, 0xbf, 0xbd]),
        });
        const issues = await wordAttachedToolbarUtf8Check.run(ctx);
        expect(issues.map((issue) => issue.path).sort()).toEqual(["word/attachedToolbars.bin", "word/attachedToolbars2.bin"]);
    });
    it.each([
        ["a backslash", "a\\tb.bin"],
        ["a case variant", "A/TB.bin"],
    ])("reports the entry %s target names, once, under its real name (#350)", async (_label, spelling) => {
        const ctx = ValidationContext.fromEntries({
            "word/fontTable.xml": document(`<wne:toolbarData r:id="rId1"/><wne:toolbarData r:id="rId2"/>`),
            "word/_rels/fontTable.xml.rels": rels(
                relation("rId1", "attachedToolbars", "", spelling) + relation("rId2", "attachedToolbars", "", "a/tb.bin"),
            ),
            "word/a/tb.bin": new Uint8Array([0xef, 0xbf, 0xbd]),
        });
        const issues = await wordAttachedToolbarUtf8Check.run(ctx);
        expect(issues.map((issue) => issue.path)).toEqual(["word/a/tb.bin"]);
    });
    it("finds the entry a case-variant target alone names (#350)", async () => {
        const ctx = ValidationContext.fromEntries({
            "word/fontTable.xml": document(`<wne:toolbarData r:id="rId1"/>`),
            "word/_rels/fontTable.xml.rels": rels(relation("rId1", "attachedToolbars", "", "A/TB.bin")),
            "word/a/tb.bin": new Uint8Array([0xef, 0xbf, 0xbd]),
        });
        const issues = await wordAttachedToolbarUtf8Check.run(ctx);
        expect(issues.map((issue) => issue.path)).toEqual(["word/a/tb.bin"]);
    });
    it("looks up a repeated reference without rescanning every entry (#358)", async () => {
        const filler = Object.fromEntries(Array.from({ length: 300 }, (_, i) => [`word/media/image${i}.png`, new Uint8Array([0])]));
        const references = 60;
        const ctx = ValidationContext.fromEntries({
            ...filler,
            "word/fontTable.xml": document(`<wne:toolbarData r:id="rId1"/>`.repeat(references)),
            "word/_rels/fontTable.xml.rels": rels(relation("rId1", "attachedToolbars", "", "A/TB.bin")),
            "word/a/tb.bin": new Uint8Array([0x00]),
        });
        const names = ctx.partNames;
        let reads = 0;
        const counted = new Proxy(names, {
            get: (target, key, receiver) => {
                if (typeof key === "string" && /^\d+$/.test(key)) {
                    reads += 1;
                }
                return Reflect.get(target, key, receiver) as unknown;
            },
        });
        vi.spyOn(ctx, "partNames", "get").mockReturnValue(counted);

        expect(await wordAttachedToolbarUtf8Check.run(ctx)).toEqual([]);
        // Building one index reads each entry a bounded number of times; a
        // scan per reference would read them `references` times over.
        expect(reads).toBeLessThan(5 * names.length);
    });
    it("skips an External relationship, which names no package part (#359)", async () => {
        const ctx = ValidationContext.fromEntries({
            "word/fontTable.xml": document(`<wne:toolbarData r:id="rId1"/><wne:toolbarData r:id="rId2"/>`),
            "word/_rels/fontTable.xml.rels": rels(
                relation("rId1", "attachedToolbars", "External", "A/TB.bin") + relation("rId2", "attachedToolbars", "External", "a/tb.bin"),
            ),
            "word/a/tb.bin": new Uint8Array([0xef, 0xbf, 0xbd]),
        });
        expect(await wordAttachedToolbarUtf8Check.run(ctx)).toEqual([]);
    });
    it("resolves only the path of a target with a query (#359)", async () => {
        const named = ValidationContext.fromEntries({
            "word/fontTable.xml": document(`<wne:toolbarData r:id="rId1"/>`),
            "word/_rels/fontTable.xml.rels": rels(relation("rId1", "attachedToolbars", "", "a/tb.bin?x")),
            "word/a/tb.bin%3Fx": new Uint8Array([0xef, 0xbf, 0xbd]),
        });
        expect(await wordAttachedToolbarUtf8Check.run(named)).toEqual([]);

        const path = ValidationContext.fromEntries({
            "word/fontTable.xml": document(`<wne:toolbarData r:id="rId1"/>`),
            "word/_rels/fontTable.xml.rels": rels(relation("rId1", "attachedToolbars", "", "a/tb.bin?x")),
            "word/a/tb.bin": new Uint8Array([0xef, 0xbf, 0xbd]),
        });
        expect((await wordAttachedToolbarUtf8Check.run(path)).map((issue) => issue.path)).toEqual(["word/a/tb.bin"]);
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
