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
    equivalentPartLookup,
    findContentTypesStream,
    findEquivalentPart,
    isContentTypesStreamName,
    isRelationshipsPartName,
    PART_NAME_FOLDER_SEGMENT_DOT,
    partNameViolation,
    resolvePartName,
    normalizePartName,
    parseContentTypes,
    partExtension,
    relationshipTargetFolder,
    relationshipTargetPartName,
    resolveContentType,
    resolveTarget,
} from "../src/lib/content-types";
import { parseXml } from "../src/lib/xml-helpers";
import * as contentTypes from "../src/lib/content-types";
import * as publicApi from "../src/index";

const typesXml = (inner: string): Document => parseXml(`<?xml version="1.0"?><Types xmlns="${CONTENT_TYPES_NAMESPACE}">${inner}</Types>`);

const overrideEntry = (partName: string, contentType = "image/png"): string =>
    `<Override PartName="${partName}" ContentType="${contentType}"/>`;

const defaultEntry = (extension: string, contentType = "image/png"): string =>
    `<Default Extension="${extension}" ContentType="${contentType}"/>`;

describe("duplicate content-type declarations (#431)", () => {
    it("reports each repetition against the first spelling using ASCII/OPC keys, without changing resolution or XML", () => {
        // Namespace access lets RED fail on this assertion rather than an absent-export import error.
        expect(contentTypes).toHaveProperty("duplicateContentTypeDeclarations", expect.any(Function));
        const dom = typesXml(
            defaultEntry("xml", "application/xml") +
                defaultEntry("XML") +
                defaultEntry("xml") +
                overrideEntry("/word/styles.xml", "application/xml") +
                overrideEntry("/WORD/%73tyles.xml") +
                // Reserved escapes remain distinct; non-ASCII extensions do not Unicode-fold.
                overrideEntry("/word/a%40.xml") +
                overrideEntry("/word/a@.xml") +
                overrideEntry("/word/my%20pic.xml") +
                overrideEntry("/word/my pic.xml") +
                defaultEntry("É") +
                defaultEntry("é") +
                '<Default/><Default Extension=""/><Override/><Override PartName=""/>' +
                '<Default xmlns="urn:foreign" Extension="xml"/>' +
                '<Override xmlns="urn:foreign" PartName="/word/styles.xml"/>',
        );
        const before = dom.toString();
        const issues = contentTypes.duplicateContentTypeDeclarations(dom, "[CONTENT_TYPES].XML");
        expect(issues).toHaveLength(4);
        expect(issues).toEqual(
            expect.arrayContaining([
                expect.objectContaining({ message: expect.stringContaining('"XML" repeats "xml"') }),
                expect.objectContaining({ message: expect.stringContaining('"xml" repeats "xml"') }),
                expect.objectContaining({ message: expect.stringContaining('"/WORD/%73tyles.xml" repeats "/word/styles.xml"') }),
                expect.objectContaining({ message: expect.stringContaining('"/word/my pic.xml" repeats "/word/my%20pic.xml"') }),
            ]),
        );
        expect(
            issues.every(
                (issue) => issue.code === "ct-duplicate-declaration" && issue.path === "[CONTENT_TYPES].XML" && issue.severity === "error",
            ),
        ).toBe(true);
        expect(resolveContentType(parseContentTypes(dom), "word/styles.xml")).toEqual({
            source: "override",
            contentType: "application/xml",
        });
        expect(dom.toString()).toBe(before);
        expect(
            contentTypes.duplicateContentTypeDeclarations(
                typesXml(defaultEntry("xml") + overrideEntry("/word/styles.xml")),
                "[Content_Types].xml",
            ),
        ).toEqual([]);
    });
});

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

describe("equivalentPartLookup (#358)", () => {
    const names = ["word/document.xml", "WORD/document.xml", "word/%64ocument.xml", "word/media/a~b.png", "word/a%2Fb.xml"];

    it.each([
        "word/document.xml",
        "Word/Document.xml",
        "word/%64ocument.xml",
        "word/media/a%7Eb.png",
        "word/a%2fb.xml",
        "word%2Fdocument.xml",
        "word/other.xml",
        "",
    ])("chooses what findEquivalentPart chooses for %j, in either order", (target) => {
        expect(equivalentPartLookup(names)(target)).toBe(findEquivalentPart(names, target));
        expect(equivalentPartLookup([...names].reverse())(target)).toBe(findEquivalentPart(names, target));
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

    it.each([
        ["word/document.xml#frag", "word/document.xml"],
        ["word/document.xml?x=1", "word/document.xml"],
        ["WORD/document.xml?x#y", "word/document.xml"],
        ["word/media/a%7Eb.png#p", "word/media/a~b.png"],
    ])("resolves only the path of %j, as System.IO.Packaging does (#361)", (target, entry) => {
        expect(resolvePartName(names, target)).toBe(entry);
    });

    // Word refused a main document named through a final `.` or `..`
    // segment (session 70 s70a/s70c): the folder it leaves names no entry.
    it.each([
        ["word/document.xml/.", "word/document.xml/"],
        ["/word/document.xml/.", "word/document.xml/"],
        ["word/document.xml/x/..", "word/document.xml/"],
        ["word\\document.xml\\.", "word/document.xml/"],
        ["word/document.xml/.?x=1", "word/document.xml/"],
        ["word/document.xml/", "word/document.xml/"],
    ])("keeps the final / of %j, which names a folder and no entry (#962)", (target, folder) => {
        expect(resolvePartName(names, target)).toBe(folder);
    });

    it.each(["#frag", "?x=1", "/#frag"])("names no part for %j, which has no path (#361)", (target) => {
        expect(resolvePartName(names, target)).toBeUndefined();
    });

    it.each(["", "/", ".", "word/..", "../word/document.xml", "word/../../word/document.xml"])(
        "names no part for %j, which is empty or climbs out of the package (#356)",
        (target) => {
            expect(resolvePartName(names, target)).toBeUndefined();
        },
    );
});

describe("relationshipTargetPartName (#947)", () => {
    it.each([
        ["/word/../word/main.xml", "/word/main.xml"],
        ["/word/./main.xml", "/word/main.xml"],
    ])("resolves the dot segments of the package-absolute target %s", (target, expected) => {
        expect(relationshipTargetPartName("word", target)).toBe(expected);
    });

    it("keeps a package-absolute climb past the package root, and clamps it only when asked", () => {
        expect(relationshipTargetPartName(".", "/../word/document.xml")).toBe("/../word/document.xml");
        expect(relationshipTargetPartName(".", "/../word/document.xml", true)).toBe("/word/document.xml");
    });

    it.each(["/", "/word/.."])("keeps the package root spelled / when the package-absolute target %s resolves to it", (target) => {
        expect(relationshipTargetPartName("word", target)).toBe("/");
    });

    it("resolves a relative target as it always has (#830)", () => {
        expect(relationshipTargetPartName(".", "word/x/../main.xml")).toBe("word/main.xml");
        expect(relationshipTargetPartName("word", "../../word/x.xml")).toBe("../word/x.xml");
        expect(relationshipTargetPartName("word", "../../word/x.xml", true)).toBe("word/x.xml");
    });

    it.each([
        ["word", "../..", ".", "../"],
        ["word", "../../..", ".", "../../"],
        [".", "/..", "/", "/../"],
        ["word", "/word/../..", "/", "/../"],
        [".", "/../..", "/", "/../../"],
    ])("clamps a climb that ends in a bare .. at the package root: from %s, %s (#962)", (baseDir, target, clamped, unclamped) => {
        expect(relationshipTargetPartName(baseDir, target, true)).toBe(clamped);
        expect(relationshipTargetPartName(baseDir, target)).toBe(unclamped);
    });

    // Word read the part through each of these (session 70 s70a/s70b/s70c),
    // so an empty segment collapses, as path.posix.join collapses it and
    // RFC 3986 would not (#962).
    it.each([
        [".", "word//main.xml", "word/main.xml"],
        ["word", "/word//main.xml", "/word/main.xml"],
        ["word", "notes//footnotes.xml", "word/notes/footnotes.xml"],
        [".", "word/./main.xml", "word/main.xml"],
    ])("collapses an empty segment and resolves an inner . : from %s, %s (#962)", (baseDir, target, expected) => {
        expect(relationshipTargetPartName(baseDir, target)).toBe(expected);
    });

    // Word refused each of these (session 70 s70a/s70b/s70c): a final `.` or
    // `..` segment leaves a folder, as RFC 3986's remove_dot_segments leaves
    // it with a final `/`, which path.posix.join would drop (#962).
    it.each([
        [".", "word/main.xml/.", "word/main.xml/"],
        ["word", "/word/main.xml/.", "/word/main.xml/"],
        [".", "word/main.xml/x/..", "word/main.xml/"],
        ["word", "notes/footnotes.xml/.", "word/notes/footnotes.xml/"],
        ["word", "notes/footnotes.xml/x/..", "word/notes/footnotes.xml/"],
        [".", "word/main.xml/", "word/main.xml/"],
    ])("keeps the final / of a target naming a folder: from %s, %s (#962)", (baseDir, target, expected) => {
        expect(relationshipTargetPartName(baseDir, target)).toBe(expected);
        expect(relationshipTargetPartName(baseDir, target, true)).toBe(expected);
    });
});

describe("resolveTarget: the root and a climb past it name no part (#970)", () => {
    const lookup = equivalentPartLookup(["word/document.xml", "%2e", "%2e%2e", "%2e%2e/footnotes.xml", "footnotes.xml"]);

    it.each([
        ["word", "../..", true],
        ["word", "../../..", true],
        [".", "/..", true],
        ["word", "../..", false],
        ["word", "../../footnotes.xml", false],
        [".", "/../footnotes.xml", false],
    ] as const)("from %s, %s (clamped: %s) finds no stored encoded-dot spelling", (baseDir, target, clamp) => {
        expect(resolveTarget(lookup, baseDir, target, clamp)).toBeUndefined();
    });

    it("still finds the part a clamped climb stops at", () => {
        expect(resolveTarget(lookup, "word", "../../footnotes.xml", true)).toBe("footnotes.xml");
    });
});

describe("resolveTarget: a target naming a folder names no part (#962, session 70)", () => {
    const lookup = equivalentPartLookup(["word/main.xml", "word/notes/footnotes.xml"]);

    it.each([
        [".", "word/main.xml/."],
        ["word", "/word/main.xml/."],
        [".", "word/main.xml/x/.."],
        [".", "word/main.xml/"],
        ["word", "notes/footnotes.xml/."],
        ["word", "notes/footnotes.xml/x/.."],
    ])("from %s, %s finds no part", (baseDir, target) => {
        expect(resolveTarget(lookup, baseDir, target)).toBeUndefined();
        expect(resolveTarget(lookup, baseDir, target, true)).toBeUndefined();
    });

    it.each([
        [".", "word//main.xml", "word/main.xml"],
        ["word", "/word//main.xml", "word/main.xml"],
        [".", "word/./main.xml", "word/main.xml"],
        ["word", "notes//footnotes.xml", "word/notes/footnotes.xml"],
    ])("from %s, %s still finds %s", (baseDir, target, expected) => {
        expect(resolveTarget(lookup, baseDir, target)).toBe(expected);
    });

    it("finds no part for a folder even through a lookup that lists one", () => {
        const folders = equivalentPartLookup(["word/main.xml/"]);
        expect(resolveTarget(folders, ".", "word/main.xml/.")).toBeUndefined();
        expect(resolveTarget(folders, ".", "word/main.xml/")).toBeUndefined();
    });
});

describe("relationshipTargetFolder: the folder a target names (#975, sessions 70–72)", () => {
    it.each([
        // A final `.` or `..` segment, or a trailing `/`, whatever the type:
        // Word refused each one (sessions 70, 71 and 72).
        ["word", "styles.xml/.", "word/styles.xml/"],
        ["word", "styles.xml/", "word/styles.xml/"],
        ["word", "styles.xml\\", "word/styles.xml/"],
        ["word", "styles.xml/#frag", "word/styles.xml/"],
        ["word", "media/image1.png/x/..", "word/media/image1.png/"],
        ["word", "../customXml/item1.xml/.", "customXml/item1.xml/"],
        ["word", "/word/styles.xml/.", "word/styles.xml/"],
        [".", "docProps/core.xml/.", "docProps/core.xml/"],
        // `.` names the source part's own folder (s72a's image row).
        ["word", ".", "word/"],
        // The package root is a folder too, spelled `/`, however it is reached.
        ["word", "..", "/"],
        ["word", "../..", "/"],
        ["word", "../x/..", "/"],
        [".", ".", "/"],
        [".", "/", "/"],
        [".", "/.", "/"],
    ])("from %s, %s names the folder %s", (baseDir, target, folder) => {
        expect(relationshipTargetFolder(baseDir, target)).toBe(folder);
    });

    it.each([
        ["word", "styles.xml"],
        ["word", "./styles.xml"],
        ["word", "media//image1.png"],
        ["word", "../customXml/item1.xml"],
        ["word", "../../customXml/item1.xml"],
        [".", "/docProps/core.xml"],
        // No part reference at all: a URI scheme, an empty reference.
        ["word", "http://example.com/"],
        ["word", ""],
        ["word", "#frag"],
    ])("from %s, %s names no folder", (baseDir, target) => {
        expect(relationshipTargetFolder(baseDir, target)).toBeUndefined();
    });
});

describe("partNameViolation: ECMA-376 Part 2, 4th edition §9.1.1.1 (#455)", () => {
    const PCHAR_RULE = "must contain only RFC 3986 pchar characters";
    const M1_8_RULE = "must not percent-encode an unreserved character (OPC [M1.8])";

    // A part name shall be an IRI and shall be encoded as either a Part IRI
    // (§9.1.1.1.1, RFC 3987 isegment/ipchar: literal ucschar allowed) or a
    // Part URI (§9.1.1.1.2, RFC 3986 segment/pchar). A name passes when it
    // is either; both forbid percent-encoded (i)unreserved characters [M1.8].
    it.each([
        "/word/café.xml", // Part IRI: a literal ucschar (é is U+00E9, in %xA0-D7FF)
        "/a/%D1%86.xml", // Part URI: ECMA-376 Part 2, 4th edition, Example 9–1
        "/word/caf%C3%A9.xml", // Part URI: the escape encodes non-ASCII, not an unreserved character
        "/word/caf%c3%a9.xml", // the same escape in lowercase hex
        "/word/x¡y.xml", // ucschar starts at U+00A0 (¡ is U+00A1)
        "/word/x\u{D7FF}y.xml", // top of %xA0-D7FF
        "/word/x\u{F900}y.xml", // %xF900-FDCF (CJK ideographs)
        "/word/x\u{FFEF}y.xml", // top of %xFDF0-FFEF
        "/word/x\u{10000}y.xml", // %x10000-1FFFD, astral: two UTF-16 code units
        "/word/x\u{E1000}y.xml", // %xE1000-EFFFD: the last block starts at E1000 (#522)
        "/word/x\u{EFFFD}y.xml", // top of %xE1000-EFFFD
        "/word/cafe\u0301.xml", // a combining mark (U+0301) is ucschar too
        "/word/%E9-é.xml", // a malformed UTF-8 escape is pct-encoded data, not an encoded iunreserved
    ])("%s names a part: a valid Part IRI or Part URI", (name) => {
        expect(partNameViolation(name)).toBeUndefined();
    });

    it.each([
        // A literal ucschar makes the name a Part IRI only, and [M1.8] reads
        // "iunreserved" there: %C3%A9 encodes é, which is iunreserved.
        ["/word/caf%C3%A9-é.xml", M1_8_RULE],
        ["/word/caf%c3%a9-é.xml", M1_8_RULE], // lowercase hex decodes the same
        ["/word/é%20%C3%A9.xml", M1_8_RULE], // an escape run still encodes é after a space
        ["/word/cafe%CC%81-é.xml", M1_8_RULE], // the combining acute is iunreserved too
        // An ASCII unreserved escape is iunreserved, so both forms refuse it.
        ["/word/n%6F-é.xml", M1_8_RULE],
        // A malformed, overlong or truncated escape is pct-encoded octets, not
        // an encoded character: the scan resynchronizes at the next octet, so
        // the é that follows it is still seen (#523).
        ["/ppt/x%E9%C3%A9-é.xml", M1_8_RULE], // a lone continuation byte
        ["/ppt/x%C0%AF%C3%A9-é.xml", M1_8_RULE], // an overlong lead
        ["/ppt/x%C3%C3%A9-é.xml", M1_8_RULE], // a truncated sequence
        // A literal non-ASCII character that is not ucschar is in neither grammar.
        ["/word/x\u{0080}y.xml", PCHAR_RULE], // C1 control
        ["/word/x\u{009F}y.xml", PCHAR_RULE], // top of the C1 range, just below ucschar
        ["/word/x\u{E000}y.xml", PCHAR_RULE], // private use is not ucschar
        // RFC 3987 §2.2 ends ucschar with %xE1000-EFFFD, so the head of the
        // last block is excluded (#522).
        ["/word/x\u{E0000}y.xml", PCHAR_RULE], // U+E0000: before %xE1000-EFFFD
        ["/word/x\u{E0FFF}y.xml", PCHAR_RULE], // U+E0FFF: the excluded head of the last block
        ["/word/x\u{F8FF}y.xml", PCHAR_RULE], // top of the private-use area
        ["/word/x\u{FFFD}y.xml", PCHAR_RULE], // replacement character
        ["/word/x\u{FFFE}y.xml", PCHAR_RULE], // noncharacters are not ucschar
        ["/word/x\u{FFFF}y.xml", PCHAR_RULE],
        ["/word/é no.xml", PCHAR_RULE], // a space is in neither grammar, ucschar beside it or not
        ["/word/caf%C3%A9 no.xml", PCHAR_RULE], // the pchar rule outranks an escaped iunreserved beside a space
    ])("%s breaks both grammars: it %s", (name, rule) => {
        expect(partNameViolation(name)).toBe(rule);
    });

    it.each([
        ["/word/café.xml?x=1", "must not contain '?', '#' or '\\'"],
        ["/word/café%2F.xml", "must not percent-encode '/' or '\\' (OPC [M1.7])"],
        ["/word/é/../a.xml", "must not have a '.' or '..' segment"],
        ["/word/.../éa.xml", "must not have a dot-only segment (OPC [M1.10])"],
        ["/word/café.xml.", "must not end with '.'"],
        ["/word/café./a.xml", PART_NAME_FOLDER_SEGMENT_DOT],
    ])("keeps every other rule's message and order for %s", (name, rule) => {
        expect(partNameViolation(name)).toBe(rule);
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

describe("isRelationshipsPartName (#398)", () => {
    it.each([
        "_rels/.rels",
        "_RELS/.RELS",
        "word/_rels/document.xml.Rels",
        "word/_rels/document.xml.%72els",
        "word/_rels/document.xml.%52ELS",
    ])("recognises %s by OPC equivalence", (part) => {
        expect(isRelationshipsPartName(part)).toBe(true);
    });
    it.each(["word/rels", "word/x.relsx", "word/x.rels.xml"])("does not recognise %s", (part) => {
        expect(isRelationshipsPartName(part)).toBe(false);
    });
    it("decodes an escaped dot as an unreserved character", () => {
        expect(canonicalPartName("word/x.%2Erels")).toBe(canonicalPartName("word/x..rels"));
        expect(isRelationshipsPartName("word/x.%2Erels")).toBe(true);
    });
});

describe("content-types stream names (#396)", () => {
    it.each([
        ["[Content_Types].xml", true],
        ["[CONTENT_TYPES].XML", true],
        ["[content_types].xml", true],
        ["/[CONTENT_TYPES].XML", true],
        ["\\[content_types].xml", true],
        ["//[Content_Types].xml", false],
        ["word/[Content_Types].xml", false],
        ["%5BContent_Types%5D.xml", false],
        ["[Content_%54ypes].xml", false],
        ["[Content_Types].%78ml", false],
        ["[Content_Typeſ].xml", false],
        ["", false],
    ])("matches ZIP stream %j without decoding escapes", (name, expected) => {
        expect(isContentTypesStreamName(name)).toBe(expected);
    });

    it.each([
        [["word/document.xml", "[CONTENT_TYPES].XML", "[Content_Types].xml"], "[Content_Types].xml"],
        [["[Content_Types].xml", "[CONTENT_TYPES].XML"], "[Content_Types].xml"],
        [["word/document.xml", "[content_types].xml", "[CONTENT_TYPES].XML"], "[content_types].xml"],
        [["[CONTENT_TYPES].XML", "[content_types].xml"], "[CONTENT_TYPES].XML"],
        [["/[CONTENT_TYPES].XML"], "/[CONTENT_TYPES].XML"],
        [["%5BContent_Types%5D.xml"], undefined],
        [[], undefined],
    ])("resolves the stored stream from %j", (names, expected) => {
        // A one-shot iterable verifies that exact preference needs only one pass.
        expect(findContentTypesStream(new Set(names).values())).toBe(expected);
    });
});

it("does not publish the orphaned hasContentTypeParameter helper (#580)", () => {
    expect(publicApi).not.toHaveProperty("hasContentTypeParameter");
    expect(publicApi.contentTypeMediaType(" Application/XML; charset=utf-8 ")).toBe("application/xml");
});
