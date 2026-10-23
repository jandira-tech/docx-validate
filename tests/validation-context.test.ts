/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

import JSZip from "jszip";
import { describe, expect, it } from "vitest";

import { ValidationContext } from "../src/lib/validation-context";
import { poisonEntry } from "./support/poison-entry";

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const DOCUMENT = `<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="${W}"><w:body><w:p/></w:body></w:document>`;

describe("ValidationContext", () => {
    it("normalizes part names: no leading slash, forward slashes, sorted", () => {
        const ctx = ValidationContext.fromEntries({
            "/word/document.xml": DOCUMENT,
            "[Content_Types].xml": "<Types/>",
            "word\\styles.xml": "<x/>",
        });
        expect(ctx.partNames).toEqual(["[Content_Types].xml", "word/document.xml", "word/styles.xml"]);
        expect(ctx.has("/word/document.xml")).toBe(true);
        expect(ctx.has("word/document.xml")).toBe(true);
        expect(ctx.has("word/missing.xml")).toBe(false);
        expect(ctx.source).toBe("memory");
        expect(ctx.archiveBytes).toBeUndefined();
    });

    it("returns the part's bytes and undefined for a missing part", () => {
        const ctx = ValidationContext.fromEntries({ "word/document.xml": DOCUMENT, "word/media/a.png": new Uint8Array([1, 2, 3]) });
        expect(ctx.bytes("word/media/a.png")).toEqual(new Uint8Array([1, 2, 3]));
        expect(new TextDecoder().decode(ctx.bytes("word/document.xml"))).toBe(DOCUMENT);
        expect(ctx.bytes("word/none.xml")).toBeUndefined();
        expect(ctx.text("word/none.xml")).toBeUndefined();
        expect(ctx.xml("word/none.xml")).toBeUndefined();
    });

    it("reads a lazily loaded part once, however often it is requested", () => {
        let loads = 0;
        const ctx = ValidationContext.fromEntries([
            [
                "word/document.xml",
                () => {
                    loads += 1;
                    return new TextEncoder().encode(DOCUMENT);
                },
            ],
        ]);
        expect(loads).toBe(0);
        ctx.bytes("word/document.xml");
        ctx.text("word/document.xml");
        ctx.xml("word/document.xml");
        ctx.xml("word/document.xml");
        expect(loads).toBe(1);
        expect(ctx.readCount("word/document.xml")).toBe(1);
    });

    it("parses an XML part once and hands every caller the same document", () => {
        const ctx = ValidationContext.fromEntries({ "word/document.xml": DOCUMENT });
        const first = ctx.xml("word/document.xml");
        const second = ctx.xml("/word/document.xml");
        expect(first?.ok).toBe(true);
        expect(second).toBe(first);
        expect(ctx.parseCount("word/document.xml")).toBe(1);
    });

    it("reports malformed XML as a failed parse instead of throwing", () => {
        const ctx = ValidationContext.fromEntries({ "word/document.xml": "<w:document><w:body>" });
        const parsed = ctx.xml("word/document.xml");
        expect(parsed?.ok).toBe(false);
        expect(parsed && !parsed.ok ? parsed.message : "").not.toBe("");
        expect(ctx.xml("word/document.xml")).toBe(parsed);
        expect(ctx.parseCount("word/document.xml")).toBe(1);
    });

    it("decodes UTF-8 with a BOM and UTF-16 parts by their byte order mark", () => {
        const utf8Bom = new Uint8Array([0xef, 0xbb, 0xbf, ...new TextEncoder().encode(DOCUMENT)]);
        const utf16le = new Uint8Array([0xff, 0xfe, ...Array.from(DOCUMENT).flatMap((ch) => [ch.charCodeAt(0), 0])]);
        const utf16be = new Uint8Array([0xfe, 0xff, ...Array.from(DOCUMENT).flatMap((ch) => [0, ch.charCodeAt(0)])]);
        const ctx = ValidationContext.fromEntries({ "a.xml": utf8Bom, "b.xml": utf16le, "c.xml": utf16be });
        for (const part of ["a.xml", "b.xml", "c.xml"]) {
            expect(ctx.text(part)).toBe(DOCUMENT);
            expect(ctx.xml(part)?.ok).toBe(true);
        }
    });

    it("decodes a UTF-16 part without a byte order mark by its XML declaration", () => {
        const declared = `<?xml version="1.0" encoding="UTF-16"?>${DOCUMENT.replace(/^<\?xml[^>]*\?>/, "")}`;
        const utf16le = new Uint8Array(Array.from(declared).flatMap((ch) => [ch.charCodeAt(0), 0]));
        const ctx = ValidationContext.fromEntries({ "word/document.xml": utf16le });
        expect(ctx.text("word/document.xml")).toBe(declared);
        expect(ctx.xml("word/document.xml")?.ok).toBe(true);
    });

    it("lists the Word story parts that carry user text", () => {
        const ctx = ValidationContext.fromEntries({
            "word/document.xml": DOCUMENT,
            "word/comments.xml": "<x/>",
            "word/footnotes.xml": "<x/>",
            "word/endnotes.xml": "<x/>",
            "word/header1.xml": "<x/>",
            "word/footer12.xml": "<x/>",
            "word/styles.xml": "<x/>",
            "word/settings.xml": "<x/>",
            "word/media/image1.png": new Uint8Array([0]),
        });
        expect(ctx.storyParts()).toEqual([
            "word/comments.xml",
            "word/document.xml",
            "word/endnotes.xml",
            "word/footer12.xml",
            "word/footnotes.xml",
            "word/header1.xml",
        ]);
    });

    it("loads a packed archive and keeps its original bytes", async () => {
        const zip = new JSZip();
        zip.file("word/document.xml", DOCUMENT);
        zip.folder("word/empty-dir");
        const bytes = await zip.generateAsync({ type: "uint8array" });
        const ctx = await ValidationContext.fromArchive(bytes);
        expect(ctx.source).toBe("archive");
        expect(ctx.archiveBytes).toBe(bytes);
        expect(ctx.partNames).toEqual(["word/document.xml"]);
        expect(ctx.text("word/document.xml")).toBe(DOCUMENT);
    });

    it("rejects bytes that are not a ZIP archive", async () => {
        await expect(ValidationContext.fromArchive(new Uint8Array([1, 2, 3]))).rejects.toThrow();
    });

    it("keeps caches per context, so one run cannot see another run's documents", () => {
        const a = ValidationContext.fromEntries({ "word/document.xml": DOCUMENT });
        const b = ValidationContext.fromEntries({ "word/document.xml": DOCUMENT.replace("<w:p/>", "<w:p/><w:p/>") });
        const docA = a.xml("word/document.xml");
        const docB = b.xml("word/document.xml");
        expect(docA).not.toBe(docB);
        expect(b.text("word/document.xml")).toContain("<w:p/><w:p/>");
        expect(a.text("word/document.xml")).not.toContain("<w:p/><w:p/>");
    });
});

describe("ValidationContext: the content Word consumes (#291)", () => {
    const MC = "http://schemas.openxmlformats.org/markup-compatibility/2006";
    const WPS = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape";
    const alternate = (requires: string) =>
        `<w:document xmlns:w="${W}" xmlns:mc="${MC}" xmlns:wps="${WPS}" xmlns:e="urn:future"><w:body>` +
        `<mc:AlternateContent><mc:Choice Requires="${requires}"><w:tbl/></mc:Choice><mc:Fallback><w:p/></mc:Fallback></mc:AlternateContent>` +
        `</w:body></w:document>`;
    const bodyChildren = (document: Document): string[] =>
        Array.from(document.documentElement!.firstChild!.childNodes)
            .filter((node) => node.nodeType === 1)
            .map((node) => (node as Element).nodeName);

    it("selects the Fallback of a Choice Word does not understand and leaves the raw DOM untouched", () => {
        const ctx = ValidationContext.fromEntries({ "word/document.xml": alternate("e") });
        const effective = ctx.effectiveXml("word/document.xml");
        expect(effective?.ok).toBe(true);
        expect(bodyChildren((effective as { document: Document }).document)).toEqual(["w:p"]);
        const raw = ctx.xml("word/document.xml");
        expect(bodyChildren((raw as { document: Document }).document)).toEqual(["mc:AlternateContent"]);
    });

    it("selects a Choice that requires a namespace Word understands", () => {
        const ctx = ValidationContext.fromEntries({ "word/document.xml": alternate("wps") });
        const effective = ctx.effectiveXml("word/document.xml");
        expect(bodyChildren((effective as { document: Document }).document)).toEqual(["w:tbl"]);
    });

    it("derives the view from the one cached parse", () => {
        const ctx = ValidationContext.fromEntries({ "word/document.xml": alternate("e") });
        const first = ctx.effectiveXml("word/document.xml");
        ctx.xml("word/document.xml");
        expect(ctx.effectiveXml("word/document.xml")).toBe(first);
        expect(ctx.parseCount("word/document.xml")).toBe(1);
    });

    it("passes a parse failure through and has no view of a missing part", () => {
        const ctx = ValidationContext.fromEntries({ "word/document.xml": "<w:document><w:body>" });
        expect(ctx.effectiveXml("word/document.xml")?.ok).toBe(false);
        expect(ctx.effectiveXml("word/missing.xml")).toBeUndefined();
    });

    it("keeps the compatibility findings of that view, with the part's path", () => {
        const ctx = ValidationContext.fromEntries({
            "word/document.xml": alternate("e").replace("<mc:Choice", "<w:p/><mc:Choice"),
            "word/header1.xml": `<w:hdr xmlns:w="${W}" xmlns:mc="${MC}" xmlns:e="urn:future" mc:MustUnderstand="e"/>`,
        });
        expect(ctx.compatibilityIssues("word/document.xml")).toEqual([
            expect.objectContaining({ code: "mce-malformed", severity: "error", path: "word/document.xml" }),
        ]);
        expect(ctx.compatibilityIssues("word/header1.xml")).toEqual([
            expect.objectContaining({ code: "mce-must-understand", severity: "error", path: "word/header1.xml" }),
        ]);
        expect(ctx.parseCount("word/document.xml")).toBe(1);
    });

    it("finds nothing to report where Word understands the required namespace", () => {
        const W14 = "http://schemas.microsoft.com/office/word/2010/wordml";
        const ctx = ValidationContext.fromEntries({
            "word/document.xml": `<w:document xmlns:w="${W}" xmlns:mc="${MC}" xmlns:w14="${W14}" mc:MustUnderstand="w14"><w:body/></w:document>`,
        });
        expect(ctx.compatibilityIssues("word/document.xml")).toEqual([]);
        expect(ctx.compatibilityIssues("word/missing.xml")).toEqual([]);
        expect(ValidationContext.fromEntries({ "word/document.xml": "<w:document>" }).compatibilityIssues("word/document.xml")).toEqual([]);
    });
});

describe("ValidationContext: a view with unparseable parts (#307)", () => {
    it("reports the listed parts unparseable, shares every other parse, and leaves the original untouched", () => {
        const ctx = ValidationContext.fromEntries({ "word/document.xml": DOCUMENT, "word/styles.xml": DOCUMENT });
        const view = ctx.withUnparseable(new Map([["/word/styles.xml", "not well-formed"]]));
        expect(view.xml("word/styles.xml")).toEqual({ ok: false, message: "not well-formed" });
        expect(view.effectiveXml("word/styles.xml")?.ok).toBe(false);
        expect(view.compatibilityIssues("word/styles.xml")).toEqual([]);
        expect(view.xml("word/document.xml")).toBe(ctx.xml("word/document.xml"));
        expect(ctx.parseCount("word/document.xml")).toBe(1);
        expect(ctx.xml("word/styles.xml")?.ok).toBe(true);
        expect(view.partNames).toEqual(ctx.partNames);
        expect(view.source).toBe(ctx.source);
    });
});

describe("ValidationContext: two entries naming one part (#366)", () => {
    // `\` reads as `/` and a leading `/` is dropped, so each pair names one
    // part. The entry stored under the part name itself is kept, whatever the
    // order, so a packed archive and its extracted directory keep the same bytes.
    const PAIRS = [
        ["_rels\\.rels", "_rels/.rels"],
        ["/word/document.xml", "word/document.xml"],
    ] as const;
    const ORDERS = PAIRS.flatMap(([other, kept]) => [
        [`${kept} last`, kept, [other, kept]],
        [`${kept} first`, kept, [kept, other]],
    ]) as [string, string, string[]][];

    it.each(ORDERS)("fromEntries keeps the entry stored under the part name (%s)", (_order, kept, names) => {
        const ctx = ValidationContext.fromEntries(names.map((name) => [name, name] as const));
        expect(ctx.partNames).toEqual([kept]);
        expect(ctx.text(kept)).toBe(kept);
    });

    // A ZIP item that starts with `/` names no part, so an archive refuses it (#372).
    it.each(ORDERS.filter(([, , names]) => names.every((name) => !name.startsWith("/"))))(
        "fromArchive keeps the entry stored under the part name, whatever the ZIP order (%s)",
        async (_order, kept, names) => {
            const zip = new JSZip();
            for (const name of names) {
                zip.file(name, name);
            }
            const ctx = await ValidationContext.fromArchive(await zip.generateAsync({ type: "uint8array" }));
            expect(ctx.text(kept)).toBe(kept);
        },
    );

    it.each([
        ["in order", ["/word\\a.xml", "word\\a.xml"]],
        ["reversed", ["word\\a.xml", "/word\\a.xml"]],
    ])("of two other spellings, keeps the smaller in code-unit order (%s)", (_order, names) => {
        const ctx = ValidationContext.fromEntries(names.map((name) => [name, name] as const));
        expect(ctx.text("word/a.xml")).toBe("/word\\a.xml");
    });
});

describe("ValidationContext.fromArchive: an unreadable entry (#369)", () => {
    const archiveWith = async (poisoned: string): Promise<Uint8Array> => {
        const zip = new JSZip();
        zip.file("word/document.xml", DOCUMENT);
        zip.file("word\\document.xml", DOCUMENT);
        return poisonEntry(await zip.generateAsync({ type: "uint8array", compression: "DEFLATE" }), poisoned);
    };

    it("rejects the archive when the entry it keeps cannot be read", async () => {
        await expect(ValidationContext.fromArchive(await archiveWith("word/document.xml"))).rejects.toThrow();
    });

    it("rejects the archive when an entry it drops for a duplicate name cannot be read, as extraction does", async () => {
        await expect(ValidationContext.fromArchive(await archiveWith("word\\document.xml"))).rejects.toThrow();
    });
});

describe("ValidationContext.fromArchive: an item name that places no part (#372)", () => {
    it.each(["/word/bonus.xml", "\\word\\bonus.xml", "C:/bonus.xml", "c:\\bonus.xml"])("rejects the archive holding %j", async (name) => {
        const zip = new JSZip();
        zip.file("word/document.xml", DOCUMENT);
        zip.file(name, "<b/>");
        await expect(ValidationContext.fromArchive(await zip.generateAsync({ type: "uint8array" }))).rejects.toThrow(
            /names no package part/,
        );
    });

    it("still reads a part name given in code with a leading slash", () => {
        expect(ValidationContext.fromEntries({ "/word/document.xml": DOCUMENT }).partNames).toEqual(["word/document.xml"]);
    });
});

describe("ValidationContext.fromArchive: a stored item name a view would rename (#378)", () => {
    it.each(["/../word/document.xml", "./word/x.xml", "word/./x.xml", "word//x.xml", "docProps\\nested\\..\\custom.xml"])(
        "rejects the archive holding %j",
        async (name) => {
            const zip = new JSZip();
            zip.file("word/styles.xml", "<s/>");
            zip.file(name, "<x/>");
            await expect(ValidationContext.fromArchive(await zip.generateAsync({ type: "uint8array" }))).rejects.toThrow(
                /names no package part/,
            );
        },
    );

    it("still loads a name with dots inside a segment", async () => {
        const zip = new JSZip();
        zip.file("word/a..b.xml", "<x/>");
        zip.file("word/.hidden.xml", "<y/>");
        const ctx = await ValidationContext.fromArchive(await zip.generateAsync({ type: "uint8array" }));
        expect(ctx.partNames).toEqual(["word/.hidden.xml", "word/a..b.xml"]);
    });
});
