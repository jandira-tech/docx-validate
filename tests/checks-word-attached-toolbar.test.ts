/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

import { describe, expect, it, vi } from "vitest";
import { ValidationContext } from "../src/lib/validation-context";
import { document, rels, relation } from "./support/word-rule-test-xml";
import { wordAttachedToolbarUtf8Check } from "../src/scripts/office/checks/word-attached-toolbar";

const MS = "http://schemas.microsoft.com/office/2006/relationships/attachedToolbars";
// Word accepts only the exact Microsoft attachedToolbars Type (#716), so a
// toolbar relationship spells it in full — `relation` builds Office types.
const toolbarRelation = (id: string, mode = "", target = "part.xml"): string =>
    `<Relationship Id="${id}" Type="${MS}" Target="${target}"${mode ? ` TargetMode="${mode}"` : ""}/>`;

describe("word-attached-toolbar-utf8", () => {
    it("does not discover toolbar bytes through a nested relationship (#434)", async () => {
        const ctx = ValidationContext.fromEntries({
            "word/fontTable.xml": document('<wne:toolbarData r:id="nested"/>'),
            "word/_rels/fontTable.xml.rels": rels(
                `<x:Wrap xmlns:x="urn:x">${toolbarRelation("nested", "", "attachedToolbars.bin")}</x:Wrap>`,
            ),
            "word/attachedToolbars.bin": new Uint8Array([0xef, 0xbf, 0xbd]),
        });
        expect(await wordAttachedToolbarUtf8Check.run(ctx)).toEqual([]);
    });

    it("reports the Word-rejected shape with unchanged issue details", async () => {
        const ctx = ValidationContext.fromEntries({
            "word/fontTable.xml": document(`<wne:toolbarData r:id="rId1"/>`),
            "word/_rels/fontTable.xml.rels": rels(toolbarRelation("rId1", "", "attachedToolbars.bin")),
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
            "word/_rels/fontTable.xml.rels": rels(toolbarRelation("rId1", "", "attachedToolbars.bin")),
            "word/attachedToolbars.bin": new Uint8Array([0x00, 0x10]),
            "word/unreferenced.bin": new Uint8Array([0xef, 0xbf, 0xbd]),
        });
        expect(await wordAttachedToolbarUtf8Check.run(ctx)).toEqual([]);
    });
    it("reports every referenced part that carries the sequence, once each (Jules on #323)", async () => {
        const ctx = ValidationContext.fromEntries({
            "word/fontTable.xml": document(`<wne:toolbarData r:id="rId1"/><wne:toolbarData r:id="rId2"/><wne:toolbarData r:id="rId3"/>`),
            "word/_rels/fontTable.xml.rels": rels(
                toolbarRelation("rId1", "", "attachedToolbars.bin") +
                    toolbarRelation("rId2", "", "attachedToolbars2.bin") +
                    toolbarRelation("rId3", "", "attachedToolbars.bin"),
            ),
            // A second part naming the first toolbar again adds no finding.
            "word/settings.xml": document(`<wne:toolbarData r:id="rId1"/>`),
            "word/_rels/settings.xml.rels": rels(toolbarRelation("rId1", "", "attachedToolbars.bin")),
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
            "word/_rels/fontTable.xml.rels": rels(toolbarRelation("rId1", "", spelling) + toolbarRelation("rId2", "", "a/tb.bin")),
            "word/a/tb.bin": new Uint8Array([0xef, 0xbf, 0xbd]),
        });
        const issues = await wordAttachedToolbarUtf8Check.run(ctx);
        expect(issues.map((issue) => issue.path)).toEqual(["word/a/tb.bin"]);
    });
    it("finds the entry a case-variant target alone names (#350)", async () => {
        const ctx = ValidationContext.fromEntries({
            "word/fontTable.xml": document(`<wne:toolbarData r:id="rId1"/>`),
            "word/_rels/fontTable.xml.rels": rels(toolbarRelation("rId1", "", "A/TB.bin")),
            "word/a/tb.bin": new Uint8Array([0xef, 0xbf, 0xbd]),
        });
        const issues = await wordAttachedToolbarUtf8Check.run(ctx);
        expect(issues.map((issue) => issue.path)).toEqual(["word/a/tb.bin"]);
    });
    it.each(["rels", "RELS", undefined])(
        "looks up repeated references with sidecar %s without rescanning every entry (#358, #758)",
        async (extension) => {
            const filler = Object.fromEntries(Array.from({ length: 300 }, (_, i) => [`word/media/image${i}.png`, new Uint8Array([0])]));
            const references = 60;
            const ctx = ValidationContext.fromEntries({
                ...filler,
                "word/fontTable.xml": document(`<wne:toolbarData r:id="rId1"/>`.repeat(references)),
                ...(extension === undefined
                    ? { "word/settings.xml": document(`<wne:toolbarData r:id="rId1"/>`) }
                    : { [`word/_rels/fontTable.xml.${extension}`]: rels(toolbarRelation("rId1", "", "A/TB.bin")) }),
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
            // Building the indexes reads each entry a bounded number of
            // times — the part-name index, plus the keyMapCustomizations
            // reach that builds its own (#777) — where a scan per
            // reference would read them `references` times over.
            expect(reads).toBeLessThan(8 * names.length);
        },
    );
    it("skips an External relationship, which names no package part (#359)", async () => {
        const ctx = ValidationContext.fromEntries({
            "word/fontTable.xml": document(`<wne:toolbarData r:id="rId1"/><wne:toolbarData r:id="rId2"/>`),
            "word/_rels/fontTable.xml.rels": rels(
                toolbarRelation("rId1", "External", "A/TB.bin") + toolbarRelation("rId2", "External", "a/tb.bin"),
            ),
            "word/a/tb.bin": new Uint8Array([0xef, 0xbf, 0xbd]),
        });
        expect(await wordAttachedToolbarUtf8Check.run(ctx)).toEqual([]);
    });
    it("resolves only the path of a target with a query (#359)", async () => {
        const named = ValidationContext.fromEntries({
            "word/fontTable.xml": document(`<wne:toolbarData r:id="rId1"/>`),
            "word/_rels/fontTable.xml.rels": rels(toolbarRelation("rId1", "", "a/tb.bin?x")),
            "word/a/tb.bin%3Fx": new Uint8Array([0xef, 0xbf, 0xbd]),
        });
        expect(await wordAttachedToolbarUtf8Check.run(named)).toEqual([]);

        const path = ValidationContext.fromEntries({
            "word/fontTable.xml": document(`<wne:toolbarData r:id="rId1"/>`),
            "word/_rels/fontTable.xml.rels": rels(toolbarRelation("rId1", "", "a/tb.bin?x")),
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
            "word/_rels/fontTable.xml.rels": rels(toolbarRelation("rId1", "", target)),
            "word/attachedToolbars.bin": new Uint8Array([0xef, 0xbf, 0xbd]),
        });
        expect(await wordAttachedToolbarUtf8Check.run(ctx)).toHaveLength(1);
    },
);
// `../../escape.bin` names `/escape.bin`, which is not stored (#991).
it.each(["../../escape.bin", "#fragment", "missing.bin"])("leaves target %s, which names no stored part, alone", async (target) => {
    const ctx = ValidationContext.fromEntries({
        "word/fontTable.xml": document(`<wne:toolbarData r:id="rId1"/>`),
        "word/_rels/fontTable.xml.rels": rels(toolbarRelation("rId1", "", target)),
    });
    expect(await wordAttachedToolbarUtf8Check.run(ctx)).toEqual([]);
});
it("skips absent ids, unrelated types, absent targets, and unparseable sidecars", async () => {
    for (const content of [undefined, "<Relationships", rels(relation("rId1", "image")), rels(`<Relationship Id="rId1" Type="${MS}"/>`)]) {
        const ctx = ValidationContext.fromEntries({
            "word/fontTable.xml": document(`<wne:toolbarData/><wne:toolbarData r:id="rId1"/>`),
            ...(content === undefined ? {} : { "word/_rels/fontTable.xml.rels": content }),
        });
        expect(await wordAttachedToolbarUtf8Check.run(ctx)).toEqual([]);
    }
});

// #991: the toolbar resolver reads a target as the read model and both
// file-reference walks read it. A target whose last segment is `.` or `..`,
// or that ends in `/`, names a folder, which names no part (#962, #975), so
// the stored part its path spells is not the toolbar; a climb past the
// package root stops at the root (#984, #986), so `word/`'s
// `../../attachedToolbars.bin` names the stored `/attachedToolbars.bin`. An
// attachedToolbars climb is unprobed in Word; the reading follows the six
// relationship types sessions 63, 64 and 74 to 78 probed.
describe("toolbar targets read as the read model reads them (#991)", () => {
    const replacement = new Uint8Array([0xef, 0xbf, 0xbd]);
    const findings = async (target: string, stored: string) =>
        (
            await wordAttachedToolbarUtf8Check.run(
                ValidationContext.fromEntries({
                    "word/fontTable.xml": document(`<wne:toolbarData r:id="rId1"/>`),
                    "word/_rels/fontTable.xml.rels": rels(toolbarRelation("rId1", "", target)),
                    [stored]: replacement,
                }),
            )
        ).map((issue) => [issue.code, issue.path]);

    it.each(["attachedToolbars.bin/.", "attachedToolbars.bin/x/..", "attachedToolbars.bin/", "/word/attachedToolbars.bin/."])(
        "reads no toolbar through the folder target %s beside the stored word/attachedToolbars.bin",
        async (target) => {
            expect(await findings(target, "word/attachedToolbars.bin")).toEqual([]);
        },
    );

    it.each(["../../attachedToolbars.bin", "../../../attachedToolbars.bin", "/../attachedToolbars.bin"])(
        "reads the root toolbar the climbing target %s names",
        async (target) => {
            expect(await findings(target, "attachedToolbars.bin")).toEqual([["word-attached-toolbar-utf8", "attachedToolbars.bin"]]);
        },
    );
});
