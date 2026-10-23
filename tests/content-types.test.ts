/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

/**
 * Unit tests for the pure OPC content-type resolver
 * (`src/lib/content-types.ts`) — ECMA-376 Part 2 §10.1.2.4: an exact
 * `<Override PartName>` match beats a `<Default Extension>`, part names
 * compare ASCII case-insensitively on their percent-decoded form, and
 * extensions compare ASCII case-insensitively.
 */

import { describe, expect, it } from "vitest";

import {
    canonicalPartName,
    CONTENT_TYPES_NAMESPACE,
    findEquivalentPart,
    resolvePartName,
    normalizePartName,
    parseContentTypes,
    partExtension,
    resolveContentType,
} from "../src/lib/content-types";
import { parseXml } from "../src/lib/xml-helpers";

const typesXml = (inner: string): Document => parseXml(`<?xml version="1.0"?><Types xmlns="${CONTENT_TYPES_NAMESPACE}">${inner}</Types>`);

const overrideEntry = (partName: string, contentType = "image/png"): string =>
    `<Override PartName="${partName}" ContentType="${contentType}"/>`;

const defaultEntry = (extension: string, contentType = "image/png"): string =>
    `<Default Extension="${extension}" ContentType="${contentType}"/>`;

describe("normalizePartName", () => {
    it("strips the leading package slash", () => {
        expect(normalizePartName("/word/document.xml")).toBe("word/document.xml");
    });

    it("strips repeated leading slashes", () => {
        expect(normalizePartName("///word/document.xml")).toBe("word/document.xml");
    });

    it("ASCII-lowercases for comparison (OPC part-name equivalence)", () => {
        expect(normalizePartName("/Word/Media/IMAGE1.JPEG")).toBe("word/media/image1.jpeg");
    });

    it("does not Unicode-fold non-ASCII letters", () => {
        // Turkish İ (U+0130) lowercases to a two-code-point sequence under
        // String#toLowerCase; ASCII-only lowering must leave it untouched.
        expect(normalizePartName("/word/İmage.png")).toBe("word/İmage.png");
    });

    it("decodes percent-escapes (%20 → space)", () => {
        expect(normalizePartName("/word/document%20with%20space.xml")).toBe("word/document with space.xml");
    });

    it("keeps malformed escapes raw instead of throwing", () => {
        expect(normalizePartName("/word/a%zzb.xml")).toBe("word/a%zzb.xml");
    });

    it("decodes before stripping, so an encoded leading slash is still stripped", () => {
        expect(normalizePartName("%2Fword/document.xml")).toBe("word/document.xml");
    });
});

describe("partExtension", () => {
    it("returns the substring after the last dot, lowercased", () => {
        expect(partExtension("word/media/IMAGE1.JPEG")).toBe("jpeg");
    });

    it("treats _rels/.rels as extension 'rels' (OPC last-dot rule, not POSIX dotfile)", () => {
        expect(partExtension("_rels/.rels")).toBe("rels");
    });

    it("returns '' for extensionless parts", () => {
        expect(partExtension("word/media/image_3094991244")).toBe("");
    });

    it("returns '' for a trailing dot", () => {
        expect(partExtension("word/media/x.")).toBe("");
    });
});

describe("canonicalPartName", () => {
    it("decodes escapes of unreserved characters only", () => {
        expect(canonicalPartName("/word/media/a%7Ebin")).toBe("word/media/a~bin");
        expect(canonicalPartName("word/media/a~bin")).toBe("word/media/a~bin");
        expect(canonicalPartName("word/media/a%7ebin")).toBe("word/media/a~bin");
    });

    it("keeps escapes of reserved characters distinct from their literals", () => {
        expect(canonicalPartName("/word/media/a%40b.bin")).toBe("word/media/a%40b.bin");
        expect(canonicalPartName("word/media/a@b.bin")).toBe("word/media/a@b.bin");
        expect(canonicalPartName("/word/a%2Fb.xml")).toBe("word/a%2fb.xml");
        expect(canonicalPartName("word/a/b.xml")).toBe("word/a/b.xml");
    });

    it("encodes characters outside pchar as uppercase-hex UTF-8 escapes", () => {
        expect(canonicalPartName("word/media/my img.jpeg")).toBe("word/media/my%20img.jpeg");
        expect(canonicalPartName("/word/media/my%20img.jpeg")).toBe("word/media/my%20img.jpeg");
        expect(canonicalPartName("word/média.png")).toBe("word/m%c3%a9dia.png");
        expect(canonicalPartName("word/media/a?b#c.xml")).toBe("word/media/a%3fb%23c.xml");
    });

    it("encodes a percent not followed by two hex digits", () => {
        expect(canonicalPartName("/word/a%zzb.xml")).toBe("word/a%25zzb.xml");
        expect(canonicalPartName("/word/a%.xml")).toBe("word/a%25.xml");
    });

    it("uppercases escape hex and then compares ASCII case-insensitively", () => {
        expect(canonicalPartName("/Word/Media/IMAGE1.JPEG")).toBe("word/media/image1.jpeg");
        expect(canonicalPartName("a%2fb")).toBe(canonicalPartName("a%2Fb"));
    });
});

describe("findEquivalentPart (#344)", () => {
    const names = ["word/document.xml", "word/media/a~b.png", "word/a%2Fb.xml"];

    it("prefers the exact entry", () => {
        expect(findEquivalentPart([...names, "WORD/document.xml"], "WORD/document.xml")).toBe("WORD/document.xml");
    });

    it.each([
        ["WORD/document.xml", "word/document.xml"],
        ["/Word/Document.XML", "word/document.xml"],
        ["word/%64ocument.xml", "word/document.xml"],
        ["word/media/a%7Eb.png", "word/media/a~b.png"],
        ["word/a%2fb.xml", "word/a%2Fb.xml"],
    ])("finds the entry %s names under OPC equivalence", (target, entry) => {
        expect(findEquivalentPart(names, target)).toBe(entry);
    });

    it.each(["word%2Fdocument.xml", "word/other.xml", ""])("finds nothing for %j", (target) => {
        expect(findEquivalentPart(names, target)).toBeUndefined();
    });

    it("picks the same equivalent entry whatever order the names come in (#349)", () => {
        const entries = ["word/document.xml", "WORD/document.xml", "word/%64ocument.xml"];
        expect(findEquivalentPart(entries, "Word/Document.xml")).toBe("WORD/document.xml");
        expect(findEquivalentPart([...entries].reverse(), "Word/Document.xml")).toBe("WORD/document.xml");
    });
});

describe("resolvePartName (#349)", () => {
    const names = ["word/document.xml", "word/media/a~b.png"];

    it.each([
        ["word\\document.xml", "word/document.xml"],
        ["/word/document.xml", "word/document.xml"],
        ["word/./media/../document.xml", "word/document.xml"],
        ["WORD\\Document.XML", "word/document.xml"],
        ["word/media/a%7Eb.png", "word/media/a~b.png"],
    ])("resolves %j to the entry it names", (target, entry) => {
        expect(resolvePartName(names, target)).toBe(entry);
    });

    it.each([
        ["word%2Fdocument.xml", "word%2Fdocument.xml"],
        ["word\\other.xml", "word/other.xml"],
    ])("falls back to the normalized name %j when no entry is equivalent", (target, normalized) => {
        expect(resolvePartName(names, target)).toBe(normalized);
    });
});

describe("parseContentTypes", () => {
    it("indexes Overrides by normalized part name", () => {
        const table = parseContentTypes(typesXml(overrideEntry("/Word/Media/IMAGE1.JPEG", "image/jpeg")));
        expect(table.overrides.get("word/media/image1.jpeg")).toBe("image/jpeg");
    });

    it("indexes Defaults by ASCII-lowercased extension", () => {
        const table = parseContentTypes(typesXml(defaultEntry("JPEG", "image/jpeg")));
        expect(table.defaults.get("jpeg")).toBe("image/jpeg");
    });

    it("ignores declarations outside the content-types namespace", () => {
        const dom = parseXml(
            `<?xml version="1.0"?><other:Types xmlns:other="urn:example">` +
                `<other:Override PartName="/word/media/image1.jpeg" ContentType="image/jpeg"/>` +
                `<other:Default Extension="jpeg" ContentType="image/jpeg"/>` +
                `</other:Types>`,
        );
        const table = parseContentTypes(dom);
        expect(table.overrides.size).toBe(0);
        expect(table.defaults.size).toBe(0);
    });

    it("skips entries with a missing or empty PartName/Extension", () => {
        const table = parseContentTypes(
            typesXml(
                `<Override ContentType="image/jpeg"/><Override PartName="" ContentType="image/jpeg"/><Default ContentType="image/jpeg"/><Default Extension="" ContentType="image/jpeg"/>`,
            ),
        );
        expect(table.overrides.size).toBe(0);
        expect(table.defaults.size).toBe(0);
    });

    it("keeps the first declaration for duplicate part names, deterministically", () => {
        const table = parseContentTypes(
            typesXml(overrideEntry("/word/media/image1.jpeg", "image/jpeg") + overrideEntry("/word/media/image1.jpeg", "image/png")),
        );
        expect(table.overrides.get("word/media/image1.jpeg")).toBe("image/jpeg");
    });

    it("keeps the first declaration for duplicate extensions, deterministically", () => {
        const table = parseContentTypes(typesXml(defaultEntry("jpeg", "image/jpeg") + defaultEntry("JPEG", "image/png")));
        expect(table.defaults.get("jpeg")).toBe("image/jpeg");
    });

    it("indexes an Override with a missing ContentType as coverage with an empty value", () => {
        const table = parseContentTypes(typesXml(`<Override PartName="/word/media/image1.jpeg"/>`));
        expect(table.overrides.get("word/media/image1.jpeg")).toBe("");
    });

    it("indexes a Default with a missing ContentType as coverage with an empty value", () => {
        const table = parseContentTypes(typesXml(`<Default Extension="jpeg"/>`));
        expect(table.defaults.get("jpeg")).toBe("");
    });
});

describe("resolveContentType", () => {
    it("keeps escapes of reserved characters distinct from their literal form", () => {
        // Issue #297: OPC compares part names as case-insensitive ASCII
        // strings, and RFC 3986 does not make an escaped reserved character
        // equivalent to its literal — an Override for /word/media/a%40b.bin
        // must not swallow the distinct part word/media/a@b.bin.
        const table = parseContentTypes(
            typesXml(overrideEntry("/word/media/a%40b.bin", "application/custom") + defaultEntry("bin", "application/octet-stream")),
        );
        expect(resolveContentType(table, "word/media/a@b.bin")).toEqual({ contentType: "application/octet-stream", source: "default" });
        expect(resolveContentType(table, "/word/media/a%40b.bin")).toEqual({ contentType: "application/custom", source: "override" });
    });
    it("gives an exact Override precedence over a Default for the same part", () => {
        const table = parseContentTypes(
            typesXml(overrideEntry("/word/media/x.png", "image/custom-override") + defaultEntry("png", "image/png")),
        );
        expect(resolveContentType(table, "/word/media/x.png")).toEqual({ contentType: "image/custom-override", source: "override" });
    });

    it("resolves a part covered only by an exact Override (decoded path input)", () => {
        const table = parseContentTypes(typesXml(overrideEntry("/word/media/image1.jpeg", "image/jpeg")));
        expect(resolveContentType(table, "word/media/image1.jpeg")).toEqual({ contentType: "image/jpeg", source: "override" });
    });

    it("resolves OPC-equivalent case variants of an Override PartName", () => {
        const table = parseContentTypes(typesXml(overrideEntry("/Word/Media/IMAGE1.JPEG", "image/jpeg")));
        expect(resolveContentType(table, "word/media/image1.jpeg")).toEqual({ contentType: "image/jpeg", source: "override" });
    });

    it("resolves escaped Override PartNames against decoded part paths, in both input forms", () => {
        const table = parseContentTypes(typesXml(overrideEntry("/word/media/my%20img.jpeg", "image/jpeg")));
        expect(resolveContentType(table, "word/media/my img.jpeg")).toEqual({ contentType: "image/jpeg", source: "override" });
        expect(resolveContentType(table, "/word/media/my%20img.jpeg")).toEqual({ contentType: "image/jpeg", source: "override" });
    });

    it("falls back to the Default for the part's extension", () => {
        const table = parseContentTypes(typesXml(defaultEntry("jpeg", "image/jpeg")));
        expect(resolveContentType(table, "word/media/any-name.jpeg")).toEqual({ contentType: "image/jpeg", source: "default" });
    });

    it("matches Default extensions ASCII case-insensitively in both directions", () => {
        const upperDefault = parseContentTypes(typesXml(defaultEntry("JPEG", "image/jpeg")));
        expect(resolveContentType(upperDefault, "word/media/x.jpeg")?.source).toBe("default");
        const upperPart = parseContentTypes(typesXml(defaultEntry("jpeg", "image/jpeg")));
        expect(resolveContentType(upperPart, "word/media/x.JPEG")?.source).toBe("default");
    });

    it("covers _rels/.rels via the OPC last-dot extension rule", () => {
        const table = parseContentTypes(typesXml(defaultEntry("rels", "application/vnd.openxmlformats-package.relationships+xml")));
        expect(resolveContentType(table, "_rels/.rels")).toEqual({
            contentType: "application/vnd.openxmlformats-package.relationships+xml",
            source: "default",
        });
    });

    it("returns null when neither Override nor Default covers the part", () => {
        const table = parseContentTypes(typesXml(defaultEntry("rels", "application/vnd.openxmlformats-package.relationships+xml")));
        expect(resolveContentType(table, "word/media/image1.jpeg")).toBeNull();
    });

    it("returns null for an extensionless part even when Defaults exist", () => {
        const table = parseContentTypes(typesXml(defaultEntry("jpeg", "image/jpeg")));
        expect(resolveContentType(table, "word/media/image_3094991244")).toBeNull();
    });

    it("returns null when only an Override for a different part exists", () => {
        const table = parseContentTypes(typesXml(overrideEntry("/word/media/other.jpeg", "image/jpeg")));
        expect(resolveContentType(table, "word/media/image1.jpeg")).toBeNull();
    });
});
