/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

/**
 * Word-blocking reachability for XML errors (#437).
 *
 * Word 16.115 does not read a part because it lies under `word/`: an
 * unreferenced truncated `word/xtra.xml` opens clean even with its own
 * Override, and so does one named only by a relationship of an unknown type.
 * Word reads the parts its relationships reach — eagerly (styles, settings,
 * fontTable, webSettings, numbering, customXml: malformed means repair) or
 * lazily, only when the main document uses them (footer, comments,
 * footnotes: related-but-unused and malformed opens clean). A `customXml`
 * item named by a document relationship sits outside `word/` and still gets
 * repair, so the folder rule also hid a false negative.
 *
 * Evidence: docs/plans/evidence/word-probes-2026-10-24.md sessions 3c, 4, 5,
 * 8b and 9a (specimen hashes in the tables); #388, #397 regressions included.
 */

import { mkdirSync, readFileSync, writeFileSync } from "node:fs";
import path from "node:path";

import { describe, expect, it } from "vitest";

import { canonicalPartName } from "../src/lib/content-types";
import { withTempDir } from "../src/lib/run-cli";
import { ValidationContext } from "../src/lib/validation-context";
import {
    glossaryTypedRelationshipTargets,
    wordExactTypeTargets,
    wordMainReachedParts,
    wordReadParts,
    wordReadStoryParts,
    wordReadTypedTargets,
    wordPossiblyReadParts,
    wordStoryRootParts,
    wordTypedRelationshipTargets,
    wordUnobservedRelationshipTargets,
    wordReferenceLoadedSources,
} from "../src/lib/word-read-parts";
import { validate } from "../src/scripts/office/validate";
import { buildDocx, MINIMAL_PARTS } from "./support/minimal-docx";

import { BYTE_REJECTED_MAIN_PARTS } from "./support/word-profile-regression-specimens";
import { STORY_FOOTER_XML, storyContentType, storyReferenceParts } from "./support/story-content-type-specimens";
import { DOCUMENT_RELS, fallback, PACKAGE_RELS, wrapRelationship } from "./support/relationships-markup-compatibility-specimens";
import { BARE_CHART_XML, referencedChartSpecimenParts } from "./support/chart-specimens";

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const RELS_NS = "http://schemas.openxmlformats.org/package/2006/relationships";
const R_STRICT = "http://purl.oclc.org/ooxml/officeDocument/relationships";

type Parts = Record<string, string | Uint8Array>;

const CT = {
    styles: "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml",
    settings: "application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml",
    fontTable: "application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml",
    webSettings: "application/vnd.openxmlformats-officedocument.wordprocessingml.webSettings+xml",
    numbering: "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml",
    footer: "application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml",
    comments: "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml",
    footnotes: "application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml",
    header: "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml",
    endnotes: "application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml",
    glossary: "application/vnd.openxmlformats-officedocument.wordprocessingml.document.glossary+xml",
    theme: "application/vnd.openxmlformats-officedocument.theme+xml",
    generic: "application/xml",
} as const;

/** "Malformed" throughout: the root element is left unclosed (session 9a). */
const unclosed = (rootOpen: string): string => `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n${rootOpen}`;

/** The session 15 spelling: the part is cut before its root end tag, so it does not parse. */
const withoutRootEndTag = (xml: string): string => xml.slice(0, xml.lastIndexOf("</"));

const MALFORMED: Readonly<Record<string, string>> = {
    document: unclosed(`<w:document xmlns:w="${W}"><w:body><w:p><w:r><w:t>x</w:t>`),
    styles: unclosed(`<w:styles xmlns:w="${W}"><w:docDefaults>`),
    settings: unclosed(`<w:settings xmlns:w="${W}">`),
    fontTable: unclosed(`<w:fonts xmlns:w="${W}">`),
    webSettings: unclosed(`<w:webSettings xmlns:w="${W}">`),
    numbering: unclosed(`<w:numbering xmlns:w="${W}">`),
    footer: unclosed(`<w:ftr xmlns:w="${W}"><w:p>`),
    header: unclosed(`<w:hdr xmlns:w="${W}"><w:p>`),
    comments: unclosed(`<w:comments xmlns:w="${W}">`),
    footnotes: unclosed(`<w:footnotes xmlns:w="${W}">`),
    endnotes: unclosed(`<w:endnotes xmlns:w="${W}">`),
    glossary: unclosed(`<w:glossaryDocument xmlns:w="${W}">`),
    theme: unclosed(`<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="t">`),
    customXml: unclosed(`<datastore xmlns="urn:x-custom">`),
    xtra: unclosed(`<w:settings xmlns:w="${W}">`),
};

/** Append a relationship to the main document's relationships part. */
const withDocRel = (parts: Parts, id: string, type: string, target: string): Parts => ({
    ...parts,
    "word/_rels/document.xml.rels": (parts["word/_rels/document.xml.rels"] as string).replace(
        "</Relationships>",
        `<Relationship Id="${id}" Type="${type}" Target="${target}"/></Relationships>`,
    ),
});

/** Append an Override for `/partName`. */
const withOverride = (parts: Parts, partName: string, contentType: string): Parts => ({
    ...parts,
    "[Content_Types].xml": (parts["[Content_Types].xml"] as string).replace(
        "</Types>",
        `<Override PartName="/${partName}" ContentType="${contentType}"/></Types>`,
    ),
});

/** The main document with the `r` namespace declared on its root. */
const withRNamespace = (document: string): string => document.replace('mc:Ignorable="w14">', `mc:Ignorable="w14" xmlns:r="${R}">`);

/** The main document using a footer (`w:footerReference r:id` in `w:sectPr`). */
const documentUsingFooter = (id: string): string =>
    withRNamespace(MINIMAL_PARTS["word/document.xml"] as string).replace(
        "<w:sectPr>",
        `<w:sectPr><w:footerReference w:type="default" r:id="${id}"/>`,
    );

/** The main document using a comment (`w:commentReference` in a run). */
const documentUsingComment = (): string =>
    (MINIMAL_PARTS["word/document.xml"] as string).replace(
        "<w:r><w:t>Hello</w:t></w:r>",
        '<w:r><w:t>Hello</w:t></w:r><w:r><w:commentReference w:id="1"/></w:r>',
    );

/** The main document using a footnote (`w:footnoteReference` in a run). */
const documentUsingFootnote = (): string =>
    (MINIMAL_PARTS["word/document.xml"] as string).replace("<w:r><w:t>Hello</w:t></w:r>", '<w:r><w:footnoteReference w:id="1"/></w:r>');

/** The main document using a header (`w:headerReference r:id` in `w:sectPr`). */
const documentUsingHeader = (id: string): string =>
    withRNamespace(MINIMAL_PARTS["word/document.xml"] as string).replace(
        "<w:sectPr>",
        `<w:sectPr><w:headerReference w:type="default" r:id="${id}"/>`,
    );

/** The main document using an endnote (`w:endnoteReference` in a run). */
const documentUsingEndnote = (): string =>
    (MINIMAL_PARTS["word/document.xml"] as string).replace(
        "<w:r><w:t>Hello</w:t></w:r>",
        '<w:r><w:t>Hello</w:t></w:r><w:r><w:endnoteReference w:id="1"/></w:r>',
    );

/** A well-formed header part, the shape of session 15's clean-open control. */
const WELLFORMED_HEADER = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:hdr xmlns:w="${W}"><w:p><w:r><w:t>Header</w:t></w:r></w:p></w:hdr>`;

/** A part stored outside word/, named by a main-document relationship with its own Override; `doc` replaces the main document when given. */
const outsideWordPart = (stored: string, type: string, contentType: string, xml: string, doc?: string): Parts =>
    withDocRel(
        withOverride({ ...MINIMAL_PARTS, ...(doc ? { "word/document.xml": doc } : {}), [stored]: xml }, stored, contentType),
        "rId9",
        type,
        path.posix.relative("word", stored),
    );

describe("wordReadParts (#437)", () => {
    it("reaches the main document wherever it is stored", async () => {
        expect(await wordReadParts(ValidationContext.fromEntries(movedOutOfWord()))).toEqual(
            new Set(["xdocs/document.xml", "xdocs/styles.xml", "xdocs/_rels/document.xml.rels"]),
        );
    });

    it.each(["styles", "settings", "fontTable", "webSettings", "numbering", "customXml"] as const)(
        "eagerly reaches a %s part the main document relates, used or not",
        async (segment) => {
            const parts = withDocRel(
                { ...MINIMAL_PARTS, [`word/${segment}.xml`]: MALFORMED.styles! },
                "rId9",
                `${R}/${segment}`,
                `${segment}.xml`,
            );
            expect(await wordReadParts(ValidationContext.fromEntries(parts))).toContain(canonicalPartName(`word/${segment}.xml`));
        },
    );

    it("reaches an eagerly parsed part through its Strict-conformance relationship type", async () => {
        const parts = withDocRel(
            { ...MINIMAL_PARTS, "word/settings.xml": MALFORMED.settings! },
            "rId9",
            `${R_STRICT}/settings`,
            "settings.xml",
        );
        expect(await wordReadParts(ValidationContext.fromEntries(parts))).toContain(canonicalPartName("word/settings.xml"));
    });

    it.each([
        ["footer", MALFORMED.footer!, documentUsingFooter("rId9"), "word/footer1.xml"],
        ["comments", MALFORMED.comments!, documentUsingComment(), "word/comments.xml"],
        ["footnotes", MALFORMED.footnotes!, documentUsingFootnote(), "word/footnotes.xml"],
    ] as const)("reaches a lazily parsed %s part only when the document uses it", async (segment, malformed, document, part) => {
        const related = withDocRel(
            { ...MINIMAL_PARTS, "word/document.xml": document, [part]: malformed },
            "rId9",
            `${R}/${segment}`,
            path.posix.basename(part),
        );
        expect(await wordReadParts(ValidationContext.fromEntries(related)), "used").toContain(canonicalPartName(part));

        const unused = withDocRel({ ...MINIMAL_PARTS, [part]: malformed }, "rId9", `${R}/${segment}`, path.posix.basename(part));
        expect(await wordReadParts(ValidationContext.fromEntries(unused)), "related but unused").not.toContain(canonicalPartName(part));
    });

    it("collects the relationship id a footerReference names, in the Strict relationships namespace too", async () => {
        const document = (MINIMAL_PARTS["word/document.xml"] as string)
            .replace('mc:Ignorable="w14">', `mc:Ignorable="w14" xmlns:purl="${R_STRICT}">`)
            .replace("<w:sectPr>", `<w:sectPr><w:footerReference w:type="default" purl:id="rId9"/>`);
        const parts = withDocRel(
            { ...MINIMAL_PARTS, "word/document.xml": document, "word/footer1.xml": MALFORMED.footer! },
            "rId9",
            `${R}/footer`,
            "footer1.xml",
        );
        expect(await wordReadParts(ValidationContext.fromEntries(parts))).toContain(canonicalPartName("word/footer1.xml"));
    });

    it.each([["chart", `${R}/chart`, "word/charts/chart1.xml", MALFORMED.styles!]] as const)(
        "still reaches an unobserved %s relationship target under word/",
        async (_label, type, part, malformed) => {
            const parts = withDocRel({ ...MINIMAL_PARTS, [part]: malformed }, "rId9", type, path.posix.relative("word", part));
            expect(await wordReadParts(ValidationContext.fromEntries(parts))).toContain(canonicalPartName(part));
        },
    );

    it.each([
        ["Transitional", R],
        ["Strict", R_STRICT],
    ] as const)("an Office-namespace %s type with an extra path segment is no styles type (#519)", async (_label, namespace) => {
        // `…/relationships/custom/styles` is not the styles relationship type:
        // a type has a segment only when it is exactly the namespace plus one
        // segment. Its target is an unobserved Office type's target — unreached
        // outside word/, `word/`-folder behaviour inside it.
        const outside = withDocRel(
            { ...MINIMAL_PARTS, "xdocs/xtra.xml": MALFORMED.xtra! },
            "rId9",
            `${namespace}/custom/styles`,
            "../xdocs/xtra.xml",
        );
        expect(await wordReadParts(ValidationContext.fromEntries(outside)), "outside word/").not.toContain(
            canonicalPartName("xdocs/xtra.xml"),
        );

        const inside = withDocRel({ ...MINIMAL_PARTS, "word/xtra.xml": MALFORMED.xtra! }, "rId9", `${namespace}/custom/styles`, "xtra.xml");
        expect(await wordReadParts(ValidationContext.fromEntries(inside)), "under word/ keeps the folder rule").toContain(
            canonicalPartName("word/xtra.xml"),
        );
    });

    it("yields no wordTypedRelationshipTargets entry for a type with an extra path segment (#519)", async () => {
        const parts = withDocRel({ ...MINIMAL_PARTS, "word/extra.xml": MALFORMED.xtra! }, "rId9", `${R}/custom/styles`, "extra.xml");
        // The real `styles` relationship still yields word/styles.xml; the
        // `custom/styles` one yields nothing.
        expect(await wordTypedRelationshipTargets(ValidationContext.fromEntries(parts), new Set(["styles"]))).toEqual([
            { part: "word/styles.xml", segment: "styles" },
        ]);
    });

    it("reaches a Microsoft-namespace relationship target from any part's relationships", async () => {
        const settings = unclosed(`<w:settings xmlns:w="${W}" xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml">`);
        const parts: Parts = {
            ...withDocRel({ ...MINIMAL_PARTS, "word/settings.xml": settings }, "rId2", `${R}/settings`, "settings.xml"),
            "word/_rels/settings.xml.rels":
                `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n` +
                `<Relationships xmlns="${RELS_NS}">` +
                `<Relationship Id="rId3" Type="http://schemas.microsoft.com/office/2011/relationships/people" Target="people1.xml"/>` +
                `</Relationships>`,
            "word/people1.xml": MALFORMED.customXml!,
        };
        expect(await wordReadParts(ValidationContext.fromEntries(parts))).toContain(canonicalPartName("word/people1.xml"));
    });

    it("does not reach a lazily parsed type named from another part's relationships", async () => {
        // The observed types keep their special rules (main-document only);
        // rule 5 excludes them, so this footer stays unreached.
        const parts: Parts = {
            ...MINIMAL_PARTS,
            "word/_rels/styles.xml.rels":
                `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n` +
                `<Relationships xmlns="${RELS_NS}">` +
                `<Relationship Id="rId9" Type="${R}/footer" Target="footer1.xml"/>` +
                `</Relationships>`,
            "word/footer1.xml": MALFORMED.footer!,
        };
        expect(await wordReadParts(ValidationContext.fromEntries(parts))).not.toContain(canonicalPartName("word/footer1.xml"));
    });

    it.each([
        ["named by no relationship at all", false],
        ["named only by a relationship of an unknown type", true],
        ["named only by a relationship in a foreign namespace", true],
    ] as const)("does not reach a word/ part %s", async (_label, withRel) => {
        let parts: Parts = { ...MINIMAL_PARTS, "word/xtra.xml": MALFORMED.xtra! };
        if (withRel) parts = withDocRel(parts, "rId9", "urn:x:unknown", "xtra.xml");
        expect(await wordReadParts(ValidationContext.fromEntries(parts))).not.toContain(canonicalPartName("word/xtra.xml"));
    });

    it("does not reach a part outside word/ named by an unknown relationship type", async () => {
        const parts = withDocRel({ ...MINIMAL_PARTS, "xdocs/xtra.xml": MALFORMED.xtra! }, "rId9", "urn:x:unknown", "../xdocs/xtra.xml");
        expect(await wordReadParts(ValidationContext.fromEntries(parts))).not.toContain(canonicalPartName("xdocs/xtra.xml"));
    });

    it("skips external targets", async () => {
        const parts = withDocRel({ ...MINIMAL_PARTS }, "rId9", `${R}/settings`, "https://example.com/settings.xml");
        const external = {
            ...parts,
            "word/_rels/document.xml.rels": (parts["word/_rels/document.xml.rels"] as string).replace(
                'Target="https://example.com/settings.xml"',
                'Target="https://example.com/settings.xml" TargetMode="External"',
            ),
        };
        const reached = await wordReadParts(ValidationContext.fromEntries(external));
        expect(reached.has(canonicalPartName("https:/example.com/settings.xml"))).toBe(false);
        for (const part of reached) expect(part).not.toContain("example.com");
    });

    it("resolves a target relative to the main document's folder", async () => {
        const parts = withDocRel(
            { ...MINIMAL_PARTS, "customXml/item1.xml": MALFORMED.customXml! },
            "rId9",
            `${R}/customXml`,
            "../customXml/item1.xml",
        );
        expect(await wordReadParts(ValidationContext.fromEntries(parts))).toContain(canonicalPartName("customXml/item1.xml"));
    });

    it("matches targets by OPC part-name equivalence", async () => {
        const parts = withDocRel({ ...MINIMAL_PARTS, "WORD/SETTINGS.XML": MALFORMED.settings! }, "rId9", `${R}/settings`, "settings.xml");
        expect(await wordReadParts(ValidationContext.fromEntries(parts))).toContain("word/settings.xml");
    });

    it("keeps a lazily parsed part reached when the main document does not parse", async () => {
        // Whether an unparsable main document uses its footer cannot be
        // decided; keeping it reached opens no false negative (#437).
        const parts = withDocRel(
            { ...MINIMAL_PARTS, "word/document.xml": MALFORMED.document!, "word/footer1.xml": MALFORMED.footer! },
            "rId9",
            `${R}/footer`,
            "footer1.xml",
        );
        expect(await wordReadParts(ValidationContext.fromEntries(parts))).toContain(canonicalPartName("word/footer1.xml"));
    });

    it("returns nothing without a package relationships part", async () => {
        const { "_rels/.rels": _packageRels, ...rest } = MINIMAL_PARTS;
        expect(await wordReadParts(ValidationContext.fromEntries(rest))).toEqual(new Set());
    });

    it("returns nothing when the package relationships part does not parse", async () => {
        const parts: Parts = { ...MINIMAL_PARTS, "_rels/.rels": "<Relationships" };
        expect(await wordReadParts(ValidationContext.fromEntries(parts))).toEqual(new Set());
    });
});

describe("wordStoryRootParts: a package-absolute officeDocument target with dot segments (#947)", () => {
    it.each(["/word/../word/main.xml", "/word/./main.xml"])("resolves %s to the stored main document", async (target) => {
        const { "word/document.xml": _stored, "word/_rels/document.xml.rels": _storedRels, ...rest } = MINIMAL_PARTS;
        const parts: Parts = {
            ...rest,
            "_rels/.rels": relsPart(`<Relationship Id="rId1" Type="${R}/officeDocument" Target="${target}"/>`),
            "word/main.xml": MINIMAL_PARTS["word/document.xml"]!,
            "word/_rels/main.xml.rels": MINIMAL_PARTS["word/_rels/document.xml.rels"]!,
            "[Content_Types].xml": (MINIMAL_PARTS["[Content_Types].xml"] as string).replace(
                'PartName="/word/document.xml"',
                'PartName="/word/main.xml"',
            ),
        };
        expect(await wordStoryRootParts(ValidationContext.fromEntries(parts))).toEqual(new Set([canonicalPartName("word/main.xml")]));
    });
});

describe("wordReadParts: the main document's own relationships part (#514)", () => {
    it("reaches the main document's relationships part itself", async () => {
        expect(await wordReadParts(ValidationContext.fromEntries(MINIMAL_PARTS))).toContain(
            canonicalPartName("word/_rels/document.xml.rels"),
        );
    });

    it("reaches it wherever the main document is stored", async () => {
        expect(await wordReadParts(ValidationContext.fromEntries(movedOutOfWord()))).toContain(
            canonicalPartName("xdocs/_rels/document.xml.rels"),
        );
    });

    it("finds the sidecar by OPC part-name equivalence", async () => {
        const { "word/_rels/document.xml.rels": _stored, ...rest } = MINIMAL_PARTS;
        const parts: Parts = { ...rest, "WORD/_RELS/DOCUMENT.XML.RELS": MINIMAL_PARTS["word/_rels/document.xml.rels"]! };
        expect(await wordReadParts(ValidationContext.fromEntries(parts))).toContain(canonicalPartName("word/_rels/document.xml.rels"));
    });

    it("keeps the sidecar reached when it does not parse", async () => {
        const parts: Parts = { ...MINIMAL_PARTS, "word/_rels/document.xml.rels": "" };
        expect(await wordReadParts(ValidationContext.fromEntries(parts))).toContain(canonicalPartName("word/_rels/document.xml.rels"));
    });

    it("does not reach another part's relationships part", async () => {
        const parts: Parts = {
            ...MINIMAL_PARTS,
            "word/_rels/styles.xml.rels": `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="${RELS_NS}"/>`,
        };
        expect(await wordReadParts(ValidationContext.fromEntries(parts))).not.toContain(canonicalPartName("word/_rels/styles.xml.rels"));
    });
});

describe("wordMainReachedParts (#742)", () => {
    it("holds the main document itself, wherever it is stored, and every part its relationships reach", async () => {
        const parts = withDocRel(
            { ...MINIMAL_PARTS, "word/header1.xml": MALFORMED.header!, "word/extra.xml": MALFORMED.xtra! },
            "rId9",
            `${R}/customXml`,
            "extra.xml",
        );
        const second = withDocRel(parts, "rIdA", `${R}/header`, "header1.xml");
        expect(await wordMainReachedParts(ValidationContext.fromEntries(second))).toEqual(
            new Set(["word/document.xml", "word/styles.xml", "word/extra.xml", "word/header1.xml"].map(canonicalPartName)),
        );
    });

    it("holds the main document itself, wherever it is stored", async () => {
        expect(await wordMainReachedParts(ValidationContext.fromEntries(movedOutOfWord()))).toContain(
            canonicalPartName("xdocs/document.xml"),
        );
    });

    it("keeps a lazily parsed part the main document relates but does not use, ungate by usage", async () => {
        const parts = withDocRel({ ...MINIMAL_PARTS, "word/comments.xml": MALFORMED.comments! }, "rId9", `${R}/comments`, "comments.xml");
        expect(await wordMainReachedParts(ValidationContext.fromEntries(parts))).toContain(canonicalPartName("word/comments.xml"));
    });

    it("names no External target, and nothing without a package relationships part", async () => {
        const external = withDocRel({ ...MINIMAL_PARTS }, "rId9", `${R}/comments`, "https://example.com/comments.xml");
        expect(await wordMainReachedParts(ValidationContext.fromEntries(external))).toEqual(
            new Set(["word/document.xml", "word/styles.xml"].map(canonicalPartName)),
        );
        const { "_rels/.rels": _packageRels, ...rest } = MINIMAL_PARTS;
        expect(await wordMainReachedParts(ValidationContext.fromEntries(rest))).toEqual(new Set());
    });
});

describe("wordReadParts: header, endnotes, theme and glossaryDocument (#516)", () => {
    it.each([
        ["header", MALFORMED.header!, documentUsingHeader("rId9"), "word/header1.xml", `${R}/header`],
        ["endnotes", MALFORMED.endnotes!, documentUsingEndnote(), "word/endnotes.xml", `${R}/endnotes`],
    ] as const)("reaches a lazily parsed %s part only when the document uses it", async (_segment, malformed, document, part, type) => {
        const related = withDocRel(
            { ...MINIMAL_PARTS, "word/document.xml": document, [part]: malformed },
            "rId9",
            type,
            path.posix.basename(part),
        );
        expect(await wordReadParts(ValidationContext.fromEntries(related)), "used").toContain(canonicalPartName(part));

        const unused = withDocRel({ ...MINIMAL_PARTS, [part]: malformed }, "rId9", type, path.posix.basename(part));
        expect(await wordReadParts(ValidationContext.fromEntries(unused)), "related but unused").not.toContain(canonicalPartName(part));
    });

    it.each([
        ["theme", MALFORMED.theme!, "word/theme1.xml", `${R}/theme`],
        ["glossaryDocument", MALFORMED.glossary!, "word/glossary.xml", `${R}/glossaryDocument`],
    ] as const)("eagerly reaches a %s part the main document relates, used or not", async (_segment, malformed, part, type) => {
        const parts = withDocRel({ ...MINIMAL_PARTS, [part]: malformed }, "rId9", type, path.posix.relative("word", part));
        expect(await wordReadParts(ValidationContext.fromEntries(parts))).toContain(canonicalPartName(part));
    });

    it("reaches the four types wherever they are stored", async () => {
        const document = withRNamespace(documentUsingEndnote()).replace(
            "<w:sectPr>",
            `<w:sectPr><w:headerReference w:type="default" r:id="rId9"/>`,
        );
        const added = [
            `<Relationship Id="rId9" Type="${R}/header" Target="../hdr/header1.xml"/>`,
            `<Relationship Id="rId10" Type="${R}/endnotes" Target="../en/endnotes.xml"/>`,
            `<Relationship Id="rId11" Type="${R}/theme" Target="../th/theme1.xml"/>`,
            `<Relationship Id="rId12" Type="${R}/glossaryDocument" Target="../g/glossary.xml"/>`,
        ].join("");
        const parts: Parts = {
            ...MINIMAL_PARTS,
            "word/document.xml": document,
            "word/_rels/document.xml.rels": (MINIMAL_PARTS["word/_rels/document.xml.rels"] as string).replace(
                "</Relationships>",
                `${added}</Relationships>`,
            ),
            "hdr/header1.xml": MALFORMED.header!,
            "en/endnotes.xml": MALFORMED.endnotes!,
            "th/theme1.xml": MALFORMED.theme!,
            "g/glossary.xml": MALFORMED.glossary!,
        };
        const reached = await wordReadParts(ValidationContext.fromEntries(parts));
        expect([...reached]).toEqual(
            expect.arrayContaining(["hdr/header1.xml", "en/endnotes.xml", "th/theme1.xml", "g/glossary.xml"].map(canonicalPartName)),
        );
    });

    it("keeps a lazily parsed header reached when the main document does not parse", async () => {
        const parts = withDocRel(
            { ...MINIMAL_PARTS, "word/document.xml": MALFORMED.document!, "word/header1.xml": MALFORMED.header! },
            "rId9",
            `${R}/header`,
            "header1.xml",
        );
        expect(await wordReadParts(ValidationContext.fromEntries(parts))).toContain(canonicalPartName("word/header1.xml"));
    });

    it("does not reach a header named only from another part's relationships", async () => {
        // header is an observed type now, so the word/-folder rule no longer
        // covers it; like footer, it is reached only from the main document.
        const parts: Parts = {
            ...MINIMAL_PARTS,
            "word/_rels/styles.xml.rels":
                `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n` +
                `<Relationships xmlns="${RELS_NS}">` +
                `<Relationship Id="rId9" Type="${R}/header" Target="header2.xml"/>` +
                `</Relationships>`,
            "word/header2.xml": MALFORMED.header!,
        };
        expect(await wordReadParts(ValidationContext.fromEntries(parts))).not.toContain(canonicalPartName("word/header2.xml"));
    });
});

/** MINIMAL_PARTS with every `word/` part stored under `xdocs/`, relationships and all. */
const movedOutOfWord = (): Parts => {
    const moved: Parts = {};
    for (const [name, text] of Object.entries(MINIMAL_PARTS)) {
        moved[name.replaceAll("word/", "xdocs/")] = text
            .replaceAll('PartName="/word/', 'PartName="/xdocs/')
            .replaceAll('Target="word/', 'Target="xdocs/');
    }
    return moved;
};

/** A relationships part's XML. */
const relsPart = (body: string): string =>
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="${RELS_NS}">${body}</Relationships>`;

/**
 * The session-17 glossary shape (s17builder.mts): a `w:docPartBody` whose
 * paragraph carries `run` and whose `w:sectPr` carries `sectPr`.
 */
const glossaryWith = (run: string, sectPr: string): string =>
    `<w:glossaryDocument xmlns:w="${W}" xmlns:r="${R}"><w:docParts><w:docPart>` +
    `<w:docPartPr><w:name w:val="Part"/><w:category><w:name w:val="General"/><w:gallery w:val="placeholder"/></w:category>` +
    `<w:guid w:val="{6A9E2C40-3F1B-4C55-9D2E-7B1A0C9D4E11}"/></w:docPartPr>` +
    `<w:docPartBody><w:p><w:r><w:t>Part</w:t></w:r>${run}</w:p><w:sectPr>${sectPr}</w:sectPr></w:docPartBody>` +
    `</w:docPart></w:docParts></w:glossaryDocument>`;

/** The session-16 glossary: a `w:docPartBody` `w:sectPr` carrying `w:<localName> r:id="rIdH"` (s16builder.mts). */
const glossaryReferencing = (localName: "headerReference" | "footerReference"): string =>
    glossaryWith("", `<w:${localName} w:type="default" r:id="rIdH"/>`);

/** A glossary that references nothing. */
const GLOSSARY_PLAIN = `<w:glossaryDocument xmlns:w="${W}"><w:docParts/></w:glossaryDocument>`;

/** The lazy story segments a glossary's relationships part can name. */
const STORY_SEGMENTS = ["header", "footer", "comments", "footnotes", "endnotes"] as const;
type StorySegment = (typeof STORY_SEGMENTS)[number];

/** The file each session-17 specimen stores the story under (s17builder.mts). */
const STORY_FILE: Readonly<Record<StorySegment, string>> = {
    header: "header9.xml",
    footer: "footer9.xml",
    comments: "comments.xml",
    footnotes: "footnotes.xml",
    endnotes: "endnotes.xml",
};

/**
 * The glossary that uses each lazy segment, the session-17 way: a
 * `w:headerReference` / `w:footerReference` `r:id` in the `w:docPartBody`
 * `w:sectPr`, or a `w:commentReference` / `w:footnoteReference` /
 * `w:endnoteReference` in a run.
 */
const STORY_GLOSSARY: Readonly<Record<StorySegment, string>> = {
    header: glossaryReferencing("headerReference"),
    footer: glossaryReferencing("footerReference"),
    comments: glossaryWith(`<w:r><w:commentReference w:id="0"/></w:r>`, ""),
    footnotes: glossaryWith(`<w:r><w:footnoteReference w:id="1"/></w:r>`, ""),
    endnotes: glossaryWith(`<w:r><w:endnoteReference w:id="1"/></w:r>`, ""),
};

/** A well-formed story part of each kind, as session 17's controls carry them. */
const WELLFORMED_STORY: Readonly<Record<StorySegment, string>> = {
    header: WELLFORMED_HEADER,
    footer: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:ftr xmlns:w="${W}"><w:p><w:r><w:t>Footer</w:t></w:r></w:p></w:ftr>`,
    comments: `<w:comments xmlns:w="${W}"><w:comment w:id="0" w:author="A"><w:p><w:r><w:t>Comment</w:t></w:r></w:p></w:comment></w:comments>`,
    footnotes:
        `<w:footnotes xmlns:w="${W}"><w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote>` +
        `<w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>` +
        `<w:footnote w:id="1"><w:p><w:r><w:t>Footnote</w:t></w:r></w:p></w:footnote></w:footnotes>`,
    endnotes:
        `<w:endnotes xmlns:w="${W}"><w:endnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:endnote>` +
        `<w:endnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:endnote>` +
        `<w:endnote w:id="1"><w:p><w:r><w:t>Endnote</w:t></w:r></w:p></w:endnote></w:endnotes>`,
};

/**
 * The session-16/17 glossary-reach shape: `word/glossary/document.xml`,
 * reached through a main-document `glossaryDocument` relationship, whose own
 * relationships part names `word/glossary/<file>` as `rIdH`. The glossary
 * defaults to the one that uses `segment`. Pass `glossary` to replace it (a
 * glossary that references nothing, or one that does not parse) and
 * `glossaryRels` to replace its relationships part.
 */
const storyViaGlossary = (segment: StorySegment, xml: string, glossary?: string, glossaryRels?: string): Parts => ({
    ...MINIMAL_PARTS,
    "word/glossary/document.xml": glossary ?? STORY_GLOSSARY[segment],
    "word/glossary/_rels/document.xml.rels":
        glossaryRels ?? relsPart(`<Relationship Id="rIdH" Type="${R}/${segment}" Target="${STORY_FILE[segment]}"/>`),
    [`word/glossary/${STORY_FILE[segment]}`]: xml,
    "word/_rels/document.xml.rels": (MINIMAL_PARTS["word/_rels/document.xml.rels"] as string).replace(
        "</Relationships>",
        `<Relationship Id="rIdG" Type="${R}/glossaryDocument" Target="glossary/document.xml"/></Relationships>`,
    ),
    "[Content_Types].xml": (MINIMAL_PARTS["[Content_Types].xml"] as string).replace(
        "</Types>",
        `<Override PartName="/word/glossary/document.xml" ContentType="${CT.glossary}"/>` +
            `<Override PartName="/word/glossary/${STORY_FILE[segment]}" ContentType="${CT[segment]}"/></Types>`,
    ),
});

/** Session 16's `ar-endnote-ref-in-footnote` shape: the only `w:endnoteReference` lives inside a footnote. */
const FOOTNOTES_WITH_ENDNOTE_REF =
    `<w:footnotes xmlns:w="${W}">` +
    `<w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote>` +
    `<w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>` +
    `<w:footnote w:id="1"><w:p><w:r><w:t>Footnote</w:t></w:r><w:r><w:endnoteReference w:id="1"/></w:r></w:p></w:footnote>` +
    `</w:footnotes>`;
const endnoteViaFootnote = (endnotes: string): Parts =>
    withDocRel(
        withDocRel(
            withOverride(
                withOverride(
                    {
                        ...MINIMAL_PARTS,
                        "word/document.xml": documentUsingFootnote(),
                        "word/footnotes.xml": FOOTNOTES_WITH_ENDNOTE_REF,
                        "word/endnotes.xml": endnotes,
                    },
                    "word/footnotes.xml",
                    CT.footnotes,
                ),
                "word/endnotes.xml",
                CT.endnotes,
            ),
            "rIdF",
            `${R}/footnotes`,
            "footnotes.xml",
        ),
        "rIdE",
        `${R}/endnotes`,
        "endnotes.xml",
    );

/** The eagerly parsed types sessions 19 and 20 probed from a reached glossary (#562, #569). */
const GLOSSARY_EAGER_SEGMENTS = ["styles", "settings", "fontTable", "webSettings", "numbering", "theme"] as const;
type GlossaryEagerSegment = (typeof GLOSSARY_EAGER_SEGMENTS)[number];

/** The file each session-19/20 specimen stores the part under (s19builder.mts, s20builder.mts). */
const GLOSSARY_EAGER_FILE: Readonly<Record<GlossaryEagerSegment, string>> = {
    styles: "styles.xml",
    settings: "settings.xml",
    fontTable: "fontTable.xml",
    webSettings: "webSettings.xml",
    numbering: "numbering.xml",
    theme: "theme/theme1.xml",
};

/**
 * The root element each session-19/20 eager part carries, as an open/close
 * pair: `fontTable` is `w:fonts`, and a theme is an `a:theme` — the
 * DrawingML namespace the s20 Word theme lives in.
 */
const GLOSSARY_EAGER_ROOT: Readonly<Record<GlossaryEagerSegment, { readonly open: string; readonly close: string }>> = {
    styles: { open: `<w:styles xmlns:w="${W}">`, close: "</w:styles>" },
    settings: { open: `<w:settings xmlns:w="${W}">`, close: "</w:settings>" },
    fontTable: { open: `<w:fonts xmlns:w="${W}">`, close: "</w:fonts>" },
    webSettings: { open: `<w:webSettings xmlns:w="${W}">`, close: "</w:webSettings>" },
    numbering: { open: `<w:numbering xmlns:w="${W}">`, close: "</w:numbering>" },
    theme: { open: `<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="t">`, close: "</a:theme>" },
};

/**
 * The session-19/20 shape (s19builder.mts, s20builder.mts):
 * `word/glossary/document.xml`, reached through a main-document
 * `glossaryDocument` relationship, references no story; its own
 * relationships part relates `word/glossary/<file>` as `rIdE` with an
 * eager-type relationship and a matching Override.
 */
const eagerViaGlossary = (segment: GlossaryEagerSegment, xml: string): Parts => ({
    ...MINIMAL_PARTS,
    "word/glossary/document.xml": glossaryWith("", ""),
    "word/glossary/_rels/document.xml.rels": relsPart(
        `<Relationship Id="rIdE" Type="${R}/${segment}" Target="${GLOSSARY_EAGER_FILE[segment]}"/>`,
    ),
    [`word/glossary/${GLOSSARY_EAGER_FILE[segment]}`]: xml,
    "word/_rels/document.xml.rels": (MINIMAL_PARTS["word/_rels/document.xml.rels"] as string).replace(
        "</Relationships>",
        `<Relationship Id="rIdG" Type="${R}/glossaryDocument" Target="glossary/document.xml"/></Relationships>`,
    ),
    "[Content_Types].xml": (MINIMAL_PARTS["[Content_Types].xml"] as string).replace(
        "</Types>",
        `<Override PartName="/word/glossary/document.xml" ContentType="${CT.glossary}"/>` +
            `<Override PartName="/word/glossary/${GLOSSARY_EAGER_FILE[segment]}" ContentType="${CT[segment]}"/></Types>`,
    ),
});

/** The session-20 theme root, the shape of `glossary-theme-malformed`'s Word theme (s20builder.mts). */
const THEME_ROOT_OPEN = `<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="t">`;

/**
 * The session-20 shapes (s20builder.mts): a reached glossary that references
 * no story, whose own relationships part relates `target` (`stored`) as
 * `rIdX` with `override` (an Override element) when the specimen carries one.
 */
const relatedViaGlossary = (segment: string, target: string, stored: string, xml: string, override = ""): Parts => ({
    ...MINIMAL_PARTS,
    "word/glossary/document.xml": glossaryWith("", ""),
    "word/glossary/_rels/document.xml.rels": relsPart(`<Relationship Id="rIdX" Type="${R}/${segment}" Target="${target}"/>`),
    [stored]: xml,
    "word/_rels/document.xml.rels": (MINIMAL_PARTS["word/_rels/document.xml.rels"] as string).replace(
        "</Relationships>",
        `<Relationship Id="rIdG" Type="${R}/glossaryDocument" Target="glossary/document.xml"/></Relationships>`,
    ),
    "[Content_Types].xml": (MINIMAL_PARTS["[Content_Types].xml"] as string).replace(
        "</Types>",
        `<Override PartName="/word/glossary/document.xml" ContentType="${CT.glossary}"/>${override}</Types>`,
    ),
});

/**
 * The session-20 observations that stay unfollowed (#569): a `customXml`
 * item and a nested `glossaryDocument` a reached glossary relates are
 * observed, not read — malformed, both open clean in Word 16.115.
 */
const OBSERVED_NOT_READ_VIA_GLOSSARY: Readonly<Record<"customXml" | "glossaryDocument", () => Parts>> = {
    customXml: () => relatedViaGlossary("customXml", "../../customXml/item1.xml", "customXml/item1.xml", MALFORMED.customXml!),
    glossaryDocument: () =>
        relatedViaGlossary(
            "glossaryDocument",
            "glossary2.xml",
            "word/glossary/glossary2.xml",
            MALFORMED.glossary!,
            `<Override PartName="/word/glossary/glossary2.xml" ContentType="${CT.glossary}"/>`,
        ),
};

describe("wordReadParts: a reached glossary is a second story root (#539, #549, #562)", () => {
    // Session 16: Word 16.115 follows a glossaryDocument relationship to the
    // glossary's own relationships part and reads the header its
    // `w:docPartBody` `w:sectPr` names by r:id (`ar-header-ref-in-glossary-malformed`,
    // repair-warning; the well-formed control opens clean). Session 17
    // answers its open questions: the glossary's relationships part is read
    // even when the glossary references nothing (`glossary-rels-unclosed`,
    // `glossary-rels-zero-byte`, `glossary-rels-unclosed-unreferenced` — all
    // repair-warning), and the comments, footnotes and endnotes it references
    // are read (`*-ref-in-glossary-malformed` — repair-warning; the
    // well-formed controls open clean). A header it relates but does not
    // reference is not read (`header-related-from-glossary-unused-malformed`
    // opens clean), and references from inside a footnote stay unfollowed
    // (session 16's `ar-endnote-ref-in-footnote`).
    it.each(STORY_SEGMENTS)("reaches a glossary's %s when the glossary references it", async (segment) => {
        const parts = storyViaGlossary(segment, MALFORMED[segment]!);
        expect(await wordReadParts(ValidationContext.fromEntries(parts))).toContain(
            canonicalPartName(`word/glossary/${STORY_FILE[segment]}`),
        );
    });

    it.each(STORY_SEGMENTS)("does not reach a glossary's %s the glossary relates but does not reference", async (segment) => {
        const parts = storyViaGlossary(segment, MALFORMED[segment]!, GLOSSARY_PLAIN);
        const reached = await wordReadParts(ValidationContext.fromEntries(parts));
        expect(reached).toContain(canonicalPartName("word/glossary/document.xml"));
        expect(reached).not.toContain(canonicalPartName(`word/glossary/${STORY_FILE[segment]}`));
    });

    it.each(STORY_SEGMENTS)("keeps a glossary's %s reached when the glossary does not parse", async (segment) => {
        // An unparseable glossary leaves usage unknown — the glossary's bytes
        // are checked the way #536 checks the main's — so the target stays
        // reached and no false negative opens.
        const parts = storyViaGlossary(segment, MALFORMED[segment]!, MALFORMED.glossary!);
        expect(await wordReadParts(ValidationContext.fromEntries(parts))).toContain(
            canonicalPartName(`word/glossary/${STORY_FILE[segment]}`),
        );
    });

    it("reaches the glossary's own relationships part, used or not (#549)", async () => {
        // Session 17: as the main document's sidecar is read (#514), so is a
        // reached glossary's — `glossary-rels-unclosed-unreferenced` gets a
        // repair warning while its header goes unreferenced.
        for (const glossary of [STORY_GLOSSARY.header, glossaryWith("", "")]) {
            const parts = storyViaGlossary("header", MALFORMED.header!, glossary);
            expect(await wordReadParts(ValidationContext.fromEntries(parts))).toContain(
                canonicalPartName("word/glossary/_rels/document.xml.rels"),
            );
        }
    });

    it.each(GLOSSARY_EAGER_SEGMENTS)(
        "reaches a glossary's %s even when the glossary references nothing (#562, session 19; #569, session 20)",
        async (segment) => {
            // Session 19: a reached glossary parses its styles, settings,
            // fontTable, webSettings and numbering relationships eagerly —
            // `glossary-<segment>-malformed` gets a repair warning while the
            // well-formed control opens clean. Session 20: so is its theme —
            // `glossary-theme-malformed` gets a repair warning while the
            // well-formed control opens clean.
            const parts = eagerViaGlossary(segment, unclosed(GLOSSARY_EAGER_ROOT[segment].open));
            expect(await wordReadParts(ValidationContext.fromEntries(parts))).toContain(
                canonicalPartName(`word/glossary/${GLOSSARY_EAGER_FILE[segment]}`),
            );
        },
    );

    it.each(["customXml", "glossaryDocument"] as const)(
        "does not reach a glossary's %s — observed, not read (session 20, #569)",
        async (segment) => {
            // Session 20: `glossary-customxml-malformed` and
            // `glossary-nested-glossary-malformed` open clean in Word 16.115
            // while malformed — observed, not read — so their targets stay
            // unfollowed, as #549/#562 left them.
            const parts = OBSERVED_NOT_READ_VIA_GLOSSARY[segment]();
            const reached = await wordReadParts(ValidationContext.fromEntries(parts));
            expect(reached).toContain(canonicalPartName("word/glossary/document.xml"));
            expect(reached).toContain(canonicalPartName("word/glossary/_rels/document.xml.rels"));
            expect(reached).not.toContain(
                canonicalPartName(segment === "customXml" ? "customXml/item1.xml" : "word/glossary/glossary2.xml"),
            );
        },
    );
});

/** One observed specimen: the malformed part, and the word-valid verdict Word 16.115 gave it. */
interface Specimen {
    readonly label: string;
    readonly parts: () => Parts;
    /** Where the XML error lands. */
    readonly part: string;
    /** Blocking under word-valid: Word offered recovery or refused the file. */
    readonly blocking: boolean;
    /** The XML diagnostic code the defect reports. */
    readonly code?: "xml-syntax" | "xml-encoding" | "xml-undeclared-prefix";
}

const SPECIMENS: readonly Specimen[] = [
    {
        label: "unreferenced word/xtra.xml with its own Override (clean-open)",
        parts: () => withOverride({ ...MINIMAL_PARTS, "word/xtra.xml": MALFORMED.xtra! }, "word/xtra.xml", CT.generic),
        part: "word/xtra.xml",
        blocking: false,
    },
    {
        label: "word/xtra.xml named by an unknown relationship type (clean-open)",
        parts: () =>
            withDocRel(
                withOverride({ ...MINIMAL_PARTS, "word/xtra.xml": MALFORMED.xtra! }, "word/xtra.xml", CT.generic),
                "rId9",
                "urn:x:unknown",
                "xtra.xml",
            ),
        part: "word/xtra.xml",
        blocking: false,
    },
    {
        label: "footer related but unused (clean-open)",
        parts: () =>
            withDocRel(
                withOverride({ ...MINIMAL_PARTS, "word/footer1.xml": MALFORMED.footer! }, "word/footer1.xml", CT.footer),
                "rId9",
                `${R}/footer`,
                "footer1.xml",
            ),
        part: "word/footer1.xml",
        blocking: false,
    },
    {
        label: "comments related but unused (clean-open)",
        parts: () =>
            withDocRel(
                withOverride({ ...MINIMAL_PARTS, "word/comments.xml": MALFORMED.comments! }, "word/comments.xml", CT.comments),
                "rId9",
                `${R}/comments`,
                "comments.xml",
            ),
        part: "word/comments.xml",
        blocking: false,
    },
    {
        label: "footnotes related but unused (clean-open)",
        parts: () =>
            withDocRel(
                withOverride({ ...MINIMAL_PARTS, "word/footnotes.xml": MALFORMED.footnotes! }, "word/footnotes.xml", CT.footnotes),
                "rId9",
                `${R}/footnotes`,
                "footnotes.xml",
            ),
        part: "word/footnotes.xml",
        blocking: false,
    },
    {
        label: "settings malformed (repair-warning)",
        parts: () =>
            withDocRel(
                withOverride({ ...MINIMAL_PARTS, "word/settings.xml": MALFORMED.settings! }, "word/settings.xml", CT.settings),
                "rId9",
                `${R}/settings`,
                "settings.xml",
            ),
        part: "word/settings.xml",
        blocking: true,
    },
    {
        label: "fontTable malformed (repair-warning)",
        parts: () =>
            withDocRel(
                withOverride({ ...MINIMAL_PARTS, "word/fontTable.xml": MALFORMED.fontTable! }, "word/fontTable.xml", CT.fontTable),
                "rId9",
                `${R}/fontTable`,
                "fontTable.xml",
            ),
        part: "word/fontTable.xml",
        blocking: true,
    },
    {
        label: "webSettings malformed (repair-warning)",
        parts: () =>
            withDocRel(
                withOverride({ ...MINIMAL_PARTS, "word/webSettings.xml": MALFORMED.webSettings! }, "word/webSettings.xml", CT.webSettings),
                "rId9",
                `${R}/webSettings`,
                "webSettings.xml",
            ),
        part: "word/webSettings.xml",
        blocking: true,
    },
    {
        label: "numbering related but unused, malformed (repair-warning)",
        parts: () =>
            withDocRel(
                withOverride({ ...MINIMAL_PARTS, "word/numbering.xml": MALFORMED.numbering! }, "word/numbering.xml", CT.numbering),
                "rId9",
                `${R}/numbering`,
                "numbering.xml",
            ),
        part: "word/numbering.xml",
        blocking: true,
    },
    {
        label: "customXml item named by a document relationship (repair-warning)",
        parts: () =>
            withDocRel(
                withOverride({ ...MINIMAL_PARTS, "customXml/item1.xml": MALFORMED.customXml! }, "customXml/item1.xml", CT.generic),
                "rId9",
                `${R}/customXml`,
                "../customXml/item1.xml",
            ),
        part: "customXml/item1.xml",
        blocking: true,
    },
    {
        label: "footer referenced from w:sectPr (open-error)",
        parts: () => ({
            ...withDocRel(
                withOverride({ ...MINIMAL_PARTS, "word/document.xml": documentUsingFooter("rId9") }, "word/footer1.xml", CT.footer),
                "rId9",
                `${R}/footer`,
                "footer1.xml",
            ),
            "word/footer1.xml": MALFORMED.footer!,
        }),
        part: "word/footer1.xml",
        blocking: true,
    },
    {
        label: "comments referenced by w:commentReference (repair-warning)",
        parts: () => ({
            ...withDocRel(
                withOverride({ ...MINIMAL_PARTS, "word/document.xml": documentUsingComment() }, "word/comments.xml", CT.comments),
                "rId9",
                `${R}/comments`,
                "comments.xml",
            ),
            "word/comments.xml": MALFORMED.comments!,
        }),
        part: "word/comments.xml",
        blocking: true,
    },
    {
        label: "footnotes referenced by w:footnoteReference",
        parts: () => ({
            ...withDocRel(
                withOverride({ ...MINIMAL_PARTS, "word/document.xml": documentUsingFootnote() }, "word/footnotes.xml", CT.footnotes),
                "rId9",
                `${R}/footnotes`,
                "footnotes.xml",
            ),
            "word/footnotes.xml": MALFORMED.footnotes!,
        }),
        part: "word/footnotes.xml",
        blocking: true,
    },
    {
        label: "the main document malformed (#388 regression)",
        parts: () => ({ ...MINIMAL_PARTS, "word/document.xml": MALFORMED.document! }),
        part: "word/document.xml",
        blocking: true,
    },
    {
        label: "styles malformed (eagerly parsed)",
        parts: () => ({ ...MINIMAL_PARTS, "word/styles.xml": MALFORMED.styles! }),
        part: "word/styles.xml",
        blocking: true,
    },
    {
        label: "the main document declaring an unknown encoding (#397 regression)",
        parts: () => ({
            ...MINIMAL_PARTS,
            "word/document.xml": (MINIMAL_PARTS["word/document.xml"] as string).replace(
                'encoding="UTF-8"',
                'encoding="x-no-such-encoding"',
            ),
        }),
        part: "word/document.xml",
        blocking: true,
        code: "xml-encoding",
    },
    {
        label: "theme malformed (eagerly parsed since #516)",
        parts: () =>
            withDocRel(
                withOverride({ ...MINIMAL_PARTS, "word/theme1.xml": MALFORMED.theme! }, "word/theme1.xml", CT.theme),
                "rId9",
                `${R}/theme`,
                "theme1.xml",
            ),
        part: "word/theme1.xml",
        blocking: true,
    },
    {
        label: "a main document stored outside word/, malformed",
        parts: () => ({ ...movedOutOfWord(), "xdocs/document.xml": MALFORMED.document! }),
        part: "xdocs/document.xml",
        blocking: true,
    },
    {
        label: "the main document's relationships part missing its root end tag (#514, repair-warning)",
        parts: () => ({
            ...MINIMAL_PARTS,
            "word/_rels/document.xml.rels": withoutRootEndTag(MINIMAL_PARTS["word/_rels/document.xml.rels"] as string),
        }),
        part: "word/_rels/document.xml.rels",
        blocking: true,
    },
    {
        label: "the main document's relationships part zero bytes (#514, repair-warning)",
        parts: () => ({ ...MINIMAL_PARTS, "word/_rels/document.xml.rels": "" }),
        part: "word/_rels/document.xml.rels",
        blocking: true,
    },
    {
        label: "a used header stored outside word/ (#516, open-error)",
        parts: () => outsideWordPart("hdr/header1.xml", `${R}/header`, CT.header, MALFORMED.header!, documentUsingHeader("rId9")),
        part: "hdr/header1.xml",
        blocking: true,
    },
    {
        label: "an endnotes part outside word/ the document uses (#516, repair-warning)",
        parts: () => outsideWordPart("en/endnotes.xml", `${R}/endnotes`, CT.endnotes, MALFORMED.endnotes!, documentUsingEndnote()),
        part: "en/endnotes.xml",
        blocking: true,
    },
    {
        label: "a theme part outside word/ named by a document relationship (#516, repair-warning)",
        parts: () => outsideWordPart("th/theme1.xml", `${R}/theme`, CT.theme, MALFORMED.theme!),
        part: "th/theme1.xml",
        blocking: true,
    },
    {
        label: "a glossaryDocument part outside word/ named by a document relationship (#516, repair-warning)",
        parts: () => outsideWordPart("g/glossary.xml", `${R}/glossaryDocument`, CT.glossary, MALFORMED.glossary!),
        part: "g/glossary.xml",
        blocking: true,
    },
    {
        label: "an unused header under word/, related but not referenced (#516, clean-open)",
        parts: () =>
            withDocRel(
                withOverride({ ...MINIMAL_PARTS, "word/header1.xml": MALFORMED.header! }, "word/header1.xml", CT.header),
                "rId9",
                `${R}/header`,
                "header1.xml",
            ),
        part: "word/header1.xml",
        blocking: false,
    },
    {
        label: "an unused endnotes part under word/, related but not referenced (#516, clean-open)",
        parts: () =>
            withDocRel(
                withOverride({ ...MINIMAL_PARTS, "word/endnotes.xml": MALFORMED.endnotes! }, "word/endnotes.xml", CT.endnotes),
                "rId9",
                `${R}/endnotes`,
                "endnotes.xml",
            ),
        part: "word/endnotes.xml",
        blocking: false,
    },
    {
        label: "a malformed header a glossary's sectPr references (#539, repair-warning)",
        parts: () => storyViaGlossary("header", MALFORMED.header!),
        part: "word/glossary/header9.xml",
        blocking: true,
    },
    {
        label: "a malformed footer a glossary's sectPr references (#549, repair-warning)",
        parts: () => storyViaGlossary("footer", MALFORMED.footer!),
        part: "word/glossary/footer9.xml",
        blocking: true,
    },
    {
        label: "endnotes referenced only from inside a footnote (session 16, clean-open)",
        parts: () => endnoteViaFootnote(MALFORMED.endnotes!),
        part: "word/endnotes.xml",
        blocking: false,
    },
    {
        label: "a malformed comments part a glossary references (#549, repair-warning)",
        parts: () => storyViaGlossary("comments", MALFORMED.comments!),
        part: "word/glossary/comments.xml",
        blocking: true,
    },
    {
        label: "a malformed footnotes part a glossary references (#549, repair-warning)",
        parts: () => storyViaGlossary("footnotes", MALFORMED.footnotes!),
        part: "word/glossary/footnotes.xml",
        blocking: true,
    },
    {
        label: "a malformed endnotes part a glossary references (#549, repair-warning)",
        parts: () => storyViaGlossary("endnotes", MALFORMED.endnotes!),
        part: "word/glossary/endnotes.xml",
        blocking: true,
    },
    {
        label: "a malformed header a glossary relates but does not reference (#549, clean-open)",
        parts: () => storyViaGlossary("header", MALFORMED.header!, GLOSSARY_PLAIN),
        part: "word/glossary/header9.xml",
        blocking: false,
    },
    {
        label: "the glossary's relationships part missing its root end tag (#549, repair-warning)",
        parts: () =>
            storyViaGlossary(
                "header",
                MALFORMED.header!,
                undefined,
                withoutRootEndTag(relsPart(`<Relationship Id="rIdH" Type="${R}/header" Target="header9.xml"/>`)),
            ),
        part: "word/glossary/_rels/document.xml.rels",
        blocking: true,
    },
    {
        label: "the glossary's relationships part zero bytes (#549, repair-warning)",
        parts: () => storyViaGlossary("header", MALFORMED.header!, undefined, ""),
        part: "word/glossary/_rels/document.xml.rels",
        blocking: true,
    },
    {
        label: "the glossary's relationships part malformed with nothing referenced (#549, repair-warning)",
        parts: () =>
            storyViaGlossary(
                "header",
                MALFORMED.header!,
                glossaryWith("", ""),
                withoutRootEndTag(relsPart(`<Relationship Id="rIdH" Type="${R}/header" Target="header9.xml"/>`)),
            ),
        part: "word/glossary/_rels/document.xml.rels",
        blocking: true,
    },
    {
        label: "a malformed styles part a reached glossary relates (#562, repair-warning)",
        parts: () => eagerViaGlossary("styles", MALFORMED.styles!),
        part: "word/glossary/styles.xml",
        blocking: true,
    },
    {
        label: "a malformed settings part a reached glossary relates (#562, repair-warning)",
        parts: () => eagerViaGlossary("settings", MALFORMED.settings!),
        part: "word/glossary/settings.xml",
        blocking: true,
    },
    {
        label: "a malformed fontTable part a reached glossary relates (#562, repair-warning)",
        parts: () => eagerViaGlossary("fontTable", MALFORMED.fontTable!),
        part: "word/glossary/fontTable.xml",
        blocking: true,
    },
    {
        label: "a malformed webSettings part a reached glossary relates (#562, repair-warning)",
        parts: () => eagerViaGlossary("webSettings", MALFORMED.webSettings!),
        part: "word/glossary/webSettings.xml",
        blocking: true,
    },
    {
        label: "a malformed numbering part a reached glossary relates (#562, repair-warning)",
        parts: () => eagerViaGlossary("numbering", MALFORMED.numbering!),
        part: "word/glossary/numbering.xml",
        blocking: true,
    },
    {
        label: "a malformed theme part a reached glossary relates (#569, repair-warning)",
        parts: () => eagerViaGlossary("theme", MALFORMED.theme!),
        part: "word/glossary/theme/theme1.xml",
        blocking: true,
    },
    {
        label: "a customXml item a reached glossary relates, malformed (observed, not read — session 20, clean-open)",
        parts: OBSERVED_NOT_READ_VIA_GLOSSARY.customXml,
        part: "customXml/item1.xml",
        blocking: false,
    },
    {
        label: "a nested glossaryDocument a reached glossary relates, malformed (observed, not read — session 20, clean-open)",
        parts: OBSERVED_NOT_READ_VIA_GLOSSARY.glossaryDocument,
        part: "word/glossary/glossary2.xml",
        blocking: false,
    },
    {
        label: "a glossary styles part declaring an unknown encoding (#562, repair-warning)",
        parts: () =>
            eagerViaGlossary(
                "styles",
                `<?xml version="1.0" encoding="x-no-such-encoding" standalone="yes"?>\n<w:styles xmlns:w="${W}"></w:styles>`,
            ),
        part: "word/glossary/styles.xml",
        blocking: true,
        code: "xml-encoding",
    },
    {
        label: "a glossary settings part using an undeclared prefix (#562, repair-warning)",
        parts: () => eagerViaGlossary("settings", `<w:settings xmlns:w="${W}"><xx:bogus/></w:settings>`),
        part: "word/glossary/settings.xml",
        blocking: true,
        code: "xml-undeclared-prefix",
    },
    {
        label: "a glossary theme part declaring an unknown encoding (#569, repair-warning)",
        parts: () =>
            eagerViaGlossary("theme", `<?xml version="1.0" encoding="x-no-such-encoding" standalone="yes"?>\n${THEME_ROOT_OPEN}</a:theme>`),
        part: "word/glossary/theme/theme1.xml",
        blocking: true,
        code: "xml-encoding",
    },
    {
        label: "a glossary theme part using an undeclared prefix (#569, repair-warning)",
        parts: () => eagerViaGlossary("theme", `${THEME_ROOT_OPEN}<xx:bogus/></a:theme>`),
        part: "word/glossary/theme/theme1.xml",
        blocking: true,
        code: "xml-undeclared-prefix",
    },
];

const FOOTER_UNUSED = SPECIMENS[2]!;
const CUSTOMXML = SPECIMENS[9]!;
const specimenByLabel = (label: string): Specimen => SPECIMENS.find((specimen) => specimen.label === label)!;
const USED_HEADER_OUTSIDE = specimenByLabel("a used header stored outside word/ (#516, open-error)");
const HEADER_UNUSED = specimenByLabel("an unused header under word/, related but not referenced (#516, clean-open)");
const GLOSSARY_HEADER_USED = specimenByLabel("a malformed header a glossary's sectPr references (#539, repair-warning)");
const GLOSSARY_COMMENTS_USED = specimenByLabel("a malformed comments part a glossary references (#549, repair-warning)");
const GLOSSARY_RELS_BROKEN = specimenByLabel("the glossary's relationships part missing its root end tag (#549, repair-warning)");
const GLOSSARY_HEADER_UNUSED = specimenByLabel("a malformed header a glossary relates but does not reference (#549, clean-open)");
const GLOSSARY_STYLES_EAGER = specimenByLabel("a malformed styles part a reached glossary relates (#562, repair-warning)");
const GLOSSARY_THEME_EAGER = specimenByLabel("a malformed theme part a reached glossary relates (#569, repair-warning)");

describe("xml errors are Word-blocking by reachability (#437)", () => {
    it.each(SPECIMENS)("$label", { timeout: 120_000 }, async (specimen) => {
        await withTempDir(async (dir) => {
            const file = path.join(dir, "specimen.docx");
            writeFileSync(file, await buildDocx(specimen.parts()));

            const wordValid = await validate(file, { profile: "word-valid" });
            const severity = wordValid.issues.find(
                (issue) => issue.code === (specimen.code ?? "xml-syntax") && issue.path === specimen.part,
            )?.severity;
            expect(severity, "word-valid severity").toBe(specimen.blocking ? "error" : "warning");
            expect(wordValid.valid, "word-valid validity").toBe(!specimen.blocking);

            // Strict and lenient never demote: the part is malformed there.
            for (const profile of ["strict", "lenient"] as const) {
                const result = await validate(file, { profile });
                expect(
                    result.issues.find((issue) => issue.code === (specimen.code ?? "xml-syntax") && issue.path === specimen.part)?.severity,
                    profile,
                ).toBe("error");
                expect(result.valid, profile).toBe(false);
            }
        });
    });

    it.each([
        FOOTER_UNUSED,
        CUSTOMXML,
        USED_HEADER_OUTSIDE,
        HEADER_UNUSED,
        GLOSSARY_HEADER_USED,
        GLOSSARY_COMMENTS_USED,
        GLOSSARY_RELS_BROKEN,
        GLOSSARY_HEADER_UNUSED,
        GLOSSARY_STYLES_EAGER,
        GLOSSARY_THEME_EAGER,
    ])(
        "agrees between full, quick and named runs, packed and directory, .docx and .docm: $label",
        { timeout: 240_000 },
        async (specimen) => {
            await withTempDir(async (dir) => {
                const bytes = await buildDocx(specimen.parts());
                const packedDocx = path.join(dir, "packed.docx");
                const packedDocm = path.join(dir, "packed.docm");
                writeFileSync(packedDocx, bytes);
                writeFileSync(packedDocm, bytes);
                const targets = [packedDocx, packedDocm];
                for (const name of ["unpacked.docx", "unpacked.docm"]) {
                    const directory = path.join(dir, name);
                    for (const [part, content] of Object.entries(specimen.parts())) {
                        mkdirSync(path.dirname(path.join(directory, part)), { recursive: true });
                        writeFileSync(path.join(directory, part), content as string);
                    }
                    targets.push(directory);
                }

                for (const file of targets) {
                    for (const profile of ["word-valid", "lenient"] as const) {
                        const full = await validate(file, { profile });
                        const errors = (result: Awaited<ReturnType<typeof validate>>) =>
                            new Set(
                                result.issues
                                    .filter((issue) => issue.severity === "error")
                                    .map((issue) => `${issue.code}@${issue.path ?? ""}`),
                            );
                        for (const selection of [{ quick: true }, { checks: ["xml-syntax"] }]) {
                            const partial = await validate(file, { profile, ...selection });
                            const tag = `${path.basename(file)} ${profile} ${JSON.stringify(selection)}`;
                            expect(
                                partial.issues.find((issue) => issue.code === "xml-syntax" && issue.path === specimen.part)?.severity,
                                tag,
                            ).toBe(specimen.blocking || profile === "lenient" ? "error" : "warning");
                            // #381: a valid full run implies a valid partial run,
                            // and the partial errors are a subset of the full ones.
                            if (full.valid) expect(partial.valid, tag).toBe(true);
                            for (const error of errors(partial)) expect(errors(full), tag).toContain(error);
                        }
                    }
                }
            });
        },
    );

    it(
        "warns under word-valid on a malformed part outside word/ named by an Office type with an extra segment (#519)",
        { timeout: 120_000 },
        async () => {
            // `…/relationships/custom/styles` names no styles part — an unknown
            // type inside the Office namespace is an unobserved type, so a
            // target outside word/ is unreached and its XML error warns. Full,
            // quick and named runs agree (#381).
            await withTempDir(async (dir) => {
                const file = path.join(dir, "specimen.docx");
                writeFileSync(file, await buildDocx(outsideWordPart("xdocs/xtra.xml", `${R}/custom/styles`, CT.settings, MALFORMED.xtra!)));
                for (const selection of [{}, { quick: true }, { checks: ["xml-syntax"] }] as const) {
                    const wordValid = await validate(file, { profile: "word-valid", ...selection });
                    const tag = JSON.stringify(selection);
                    expect(
                        wordValid.issues.find((issue) => issue.code === "xml-syntax" && issue.path === "xdocs/xtra.xml")?.severity,
                        tag,
                    ).toBe("warning");
                    expect(wordValid.valid, tag).toBe(true);
                }
            });
        },
    );

    it("still blocks a malformed word/ part named by an Office type with an extra segment (#519)", { timeout: 60_000 }, async () => {
        await withTempDir(async (dir) => {
            const file = path.join(dir, "specimen.docx");
            writeFileSync(
                file,
                await buildDocx(
                    withDocRel(
                        withOverride({ ...MINIMAL_PARTS, "word/xtra.xml": MALFORMED.xtra! }, "word/xtra.xml", CT.settings),
                        "rId9",
                        `${R}/custom/styles`,
                        "xtra.xml",
                    ),
                ),
            );
            const wordValid = await validate(file, { profile: "word-valid" });
            expect(wordValid.issues.find((issue) => issue.code === "xml-syntax" && issue.path === "word/xtra.xml")?.severity).toBe("error");
            expect(wordValid.valid).toBe(false);
        });
    });

    it("keeps a used, well-formed header outside word/ word-valid (#516 control)", { timeout: 60_000 }, async () => {
        await withTempDir(async (dir) => {
            const file = path.join(dir, "specimen.docx");
            writeFileSync(
                file,
                await buildDocx(
                    outsideWordPart("hdr/header1.xml", `${R}/header`, CT.header, WELLFORMED_HEADER, documentUsingHeader("rId9")),
                ),
            );
            const result = await validate(file, { profile: "word-valid" });
            expect(result.valid).toBe(true);
        });
    });

    it("keeps a well-formed glossary-referenced header word-valid (#539 control)", { timeout: 60_000 }, async () => {
        // `ar-header-ref-in-glossary-wellformed-control` opens clean — the
        // shape itself is fine; only the malformed header makes Word repair.
        await withTempDir(async (dir) => {
            const file = path.join(dir, "specimen.docx");
            writeFileSync(file, await buildDocx(storyViaGlossary("header", WELLFORMED_HEADER)));
            const result = await validate(file, { profile: "word-valid" });
            expect(result.valid).toBe(true);
        });
    });

    it.each(STORY_SEGMENTS)(
        "keeps a well-formed story a glossary references word-valid (#539/#549 controls)",
        { timeout: 120_000 },
        async (segment) => {
            // Session 16/17: every `*-ref-in-glossary-wellformed-control`
            // opens clean in Word.
            await withTempDir(async (dir) => {
                const file = path.join(dir, "specimen.docx");
                writeFileSync(file, await buildDocx(storyViaGlossary(segment, WELLFORMED_STORY[segment])));
                const result = await validate(file, { profile: "word-valid" });
                expect(result.valid).toBe(true);
            });
        },
    );

    it.each(GLOSSARY_EAGER_SEGMENTS)(
        "keeps a well-formed %s part a reached glossary relates word-valid (#562/#569 controls)",
        { timeout: 60_000 },
        async (segment) => {
            // Session 19: every `glossary-<segment>-wellformed-control` opens
            // clean in Word — the malformed part is what makes Word repair.
            // Session 20: so does `glossary-theme-wellformed-control`.
            await withTempDir(async (dir) => {
                const file = path.join(dir, "specimen.docx");
                const root = GLOSSARY_EAGER_ROOT[segment];
                writeFileSync(file, await buildDocx(eagerViaGlossary(segment, `${root.open}${root.close}`)));
                const result = await validate(file, { profile: "word-valid" });
                expect(result.valid).toBe(true);
            });
        },
    );

    it("keeps the packed input bytes unchanged after validation", { timeout: 120_000 }, async () => {
        await withTempDir(async (dir) => {
            const file = path.join(dir, "specimen.docx");
            const bytes = await buildDocx(FOOTER_UNUSED.parts());
            writeFileSync(file, bytes);
            await validate(file, { profile: "word-valid" });
            await validate(file, { profile: "word-valid", checks: ["xml-syntax"] });
            expect(readFileSync(file)).toEqual(bytes);
        });
    });

    it("rejects a partial strict run through ModeConflictError", { timeout: 120_000 }, async () => {
        await withTempDir(async (dir) => {
            const file = path.join(dir, "specimen.docx");
            writeFileSync(file, await buildDocx(FOOTER_UNUSED.parts()));
            for (const selection of [{ quick: true }, { checks: ["xml-syntax"] }]) {
                await expect(validate(file, { profile: "strict", ...selection }), JSON.stringify(selection)).rejects.toThrow(/strict/i);
            }
        });
    });
});

describe("wordReadParts: byte-rejected main usage (#536)", () => {
    it.each(BYTE_REJECTED_MAIN_PARTS)("keeps $segment reached when the main DOM parses but its bytes fail", async ({ part, parts }) => {
        // In-memory integration with the byte parser: the recovering DOM alone
        // cannot establish usage for an unsupported declared encoding.
        const context = ValidationContext.fromEntries(parts);
        expect(context.effectiveXml("word/document.xml")?.ok).toBe(true);
        expect(await wordReadParts(context), "usage must remain unknown after byte-parser rejection").toContain(canonicalPartName(part));
    });

    it.each(BYTE_REJECTED_MAIN_PARTS.slice(0, 2))(
        "retains the malformed $segment error in a full run",
        async ({ part, parts }) => {
            await withTempDir(async (dir) => {
                const file = path.join(dir, "specimen.docx");
                writeFileSync(file, await buildDocx(parts));
                const result = await validate(file, { profile: "word-valid" });
                expect(result.issues).toContainEqual(
                    expect.objectContaining({ code: "xml-encoding", path: "word/document.xml", severity: "error" }),
                );
                expect(
                    result.issues.find((issue) => issue.path === part && issue.code === "xml-syntax")?.severity,
                    "usage must remain unknown after byte-parser rejection",
                ).toBe("error");
            });
        },
        60_000,
    );
});

/**
 * Word reads the relationships markup compatibility selects (#634, session
 * 34 s34a/s34b): a relationship inside an `mc:Fallback` — or an understood
 * `mc:Choice` — reaches its part exactly as a direct child does, while the
 * stored view of the same relationships part shows no `Relationship` child
 * there. Under word-valid and strict the reachability helpers read each
 * relationships part through `context.effectiveXml`; under lenient they keep
 * the stored view, so a wrapped relationship is followed nowhere it was not
 * followed before.
 */
describe("wordReadParts: relationships markup compatibility (#634)", () => {
    const PACKAGE_READ = new Set(["word/document.xml", "word/styles.xml", "word/_rels/document.xml.rels"]);
    const FOOTER_PARTS = storyReferenceParts("footerReference", "footer", "footer1.xml", STORY_FOOTER_XML, storyContentType("footer"));

    it("follows the package officeDocument relationship an mc:Fallback selects", async () => {
        const parts = wrapRelationship({ ...MINIMAL_PARTS }, PACKAGE_RELS, "rId1", fallback);
        const context = ValidationContext.fromEntries(parts);
        expect(await wordReadParts(context)).toEqual(PACKAGE_READ);
        expect(await wordTypedRelationshipTargets(context, new Set(["officeDocument"]))).toEqual([
            { part: "word/document.xml", segment: "officeDocument" },
        ]);
    });

    it("follows a footer relationship an understood w14 mc:Choice selects", async () => {
        const parts = wrapRelationship(
            FOOTER_PARTS,
            DOCUMENT_RELS,
            "rIdX",
            (rel) => `<mc:AlternateContent><mc:Choice Requires="w14">${rel}</mc:Choice><mc:Fallback/></mc:AlternateContent>`,
        );
        const read = await wordReadParts(ValidationContext.fromEntries(parts));
        expect(read).toContain("word/footer1.xml");
    });

    it("does not follow a relationship in an unselected mc:Choice", async () => {
        const parts = wrapRelationship(
            FOOTER_PARTS,
            DOCUMENT_RELS,
            "rIdX",
            (rel) => `<mc:AlternateContent><mc:Choice Requires="foo">${rel}</mc:Choice><mc:Fallback/></mc:AlternateContent>`,
        );
        const context = ValidationContext.fromEntries(parts);
        expect(await wordReadParts(context)).not.toContain("word/footer1.xml");
        expect(await wordReadParts(context, "lenient")).not.toContain("word/footer1.xml");
    });

    it("keeps the stored view under lenient: a Fallback relationship is not followed", async () => {
        const parts = wrapRelationship({ ...MINIMAL_PARTS }, PACKAGE_RELS, "rId1", fallback);
        const context = ValidationContext.fromEntries(parts);
        expect(await wordReadParts(context, "lenient")).toEqual(new Set());
        expect(await wordTypedRelationshipTargets(context, new Set(["officeDocument"]), "lenient")).toEqual([]);
        expect(await wordReadParts(context)).toEqual(PACKAGE_READ);
    });

    it("keeps a foreign wrapper's nested relationship unfollowed in the processed view (#434)", async () => {
        const parts = wrapRelationship({ ...MINIMAL_PARTS }, PACKAGE_RELS, "rId1", (rel) => fallback(`<x:Wrap>${rel}</x:Wrap>`));
        // The selected Fallback holds a foreign wrapper, not a Relationship
        // child: the officeDocument relationship inside it stays unfollowed,
        // so the package has no readable main document.
        const context = ValidationContext.fromEntries(parts);
        expect(await wordReadParts(context)).toEqual(new Set());
    });
});

/**
 * A part related by two relationship types keeps every segment (#593):
 * `wordTypedRelationshipTargets` and `glossaryTypedRelationshipTargets`
 * yield one entry per (part, segment) pair — a part reached twice no longer
 * keeps only the segment that reached it first, so the order of
 * relationships in the relationships part cannot decide which type judges
 * the part. Session 20 (s20b) observed Word 16.115 offer recovery for the
 * footer-first specimen the first-segment rule let pass.
 */
describe("a part related by two relationship types keeps every segment (#593)", () => {
    const NUMBERING_XML =
        `<w:numbering xmlns:w="${W}"><w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/>` +
        `<w:lvlText w:val="%1."/></w:lvl></w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num></w:numbering>`;

    /** The minimal package plus `word/numbering.xml`, related from the main document as `first` and then `second`. */
    const dualRelated = (first: string, second: string): Parts => ({
        ...MINIMAL_PARTS,
        "word/numbering.xml": NUMBERING_XML,
        "word/_rels/document.xml.rels": MINIMAL_PARTS["word/_rels/document.xml.rels"]!.replace(
            "</Relationships>",
            `<Relationship Id="rIdA" Type="${R}/${first}" Target="numbering.xml"/>` +
                `<Relationship Id="rIdB" Type="${R}/${second}" Target="numbering.xml"/></Relationships>`,
        ),
        "[Content_Types].xml": MINIMAL_PARTS["[Content_Types].xml"]!.replace(
            "</Types>",
            `<Override PartName="/word/numbering.xml" ContentType="${CT.numbering}"/></Types>`,
        ),
    });

    /** A reached glossary (referencing nothing) whose own relationships part relates `word/glossary/styles2.xml` as `first` and then `second`. */
    const dualGlossary = (first: string, second: string): Parts => ({
        ...MINIMAL_PARTS,
        "word/glossary/document.xml": glossaryWith("", ""),
        "word/glossary/_rels/document.xml.rels": relsPart(
            `<Relationship Id="rIdA" Type="${R}/${first}" Target="styles2.xml"/>` +
                `<Relationship Id="rIdB" Type="${R}/${second}" Target="styles2.xml"/>`,
        ),
        "word/glossary/styles2.xml": `<w:styles xmlns:w="${W}"></w:styles>`,
        "word/_rels/document.xml.rels": MINIMAL_PARTS["word/_rels/document.xml.rels"]!.replace(
            "</Relationships>",
            `<Relationship Id="rIdG" Type="${R}/glossaryDocument" Target="glossary/document.xml"/></Relationships>`,
        ),
        "[Content_Types].xml": MINIMAL_PARTS["[Content_Types].xml"]!.replace(
            "</Types>",
            `<Override PartName="/word/glossary/document.xml" ContentType="${CT.glossary}"/>` +
                `<Override PartName="/word/glossary/styles2.xml" ContentType="${CT.styles}"/></Types>`,
        ),
    });

    it.each([
        ["footer", "numbering"],
        ["numbering", "footer"],
    ] as const)("wordTypedRelationshipTargets keeps both segments, %s first", async (first, second) => {
        expect(
            await wordTypedRelationshipTargets(ValidationContext.fromEntries(dualRelated(first, second)), new Set(["footer", "numbering"])),
        ).toEqual([
            { part: "word/numbering.xml", segment: first },
            { part: "word/numbering.xml", segment: second },
        ]);
    });

    it("wordTypedRelationshipTargets yields one entry when two relationships of the same type name one part", async () => {
        expect(
            await wordTypedRelationshipTargets(
                ValidationContext.fromEntries(dualRelated("numbering", "numbering")),
                new Set(["numbering"]),
            ),
        ).toEqual([{ part: "word/numbering.xml", segment: "numbering" }]);
    });

    it.each([
        ["header", "styles"],
        ["styles", "header"],
    ] as const)("glossaryTypedRelationshipTargets keeps both segments, %s first", async (first, second) => {
        expect(
            await glossaryTypedRelationshipTargets(
                ValidationContext.fromEntries(dualGlossary(first, second)),
                new Set(["header", "styles"]),
            ),
        ).toEqual([
            { part: "word/glossary/styles2.xml", segment: first, via: "glossary", root: "word/glossary/document.xml" },
            { part: "word/glossary/styles2.xml", segment: second, via: "glossary", root: "word/glossary/document.xml" },
        ]);
    });

    it.each([
        ["footer", "numbering"],
        ["numbering", "footer"],
    ] as const)("wordReadTypedTargets reads the part through its eager segment alone, %s first", async (first, second) => {
        // The footer relationship is unreferenced — lazy and unread — while
        // the numbering one is eagerly parsed: exactly one typed read.
        const typed = await wordReadTypedTargets(ValidationContext.fromEntries(dualRelated(first, second)));
        expect(typed.filter((target) => target.part === "word/numbering.xml")).toEqual([
            { part: "word/numbering.xml", segment: "numbering" },
        ]);
    });

    it.each([
        ["header", "styles"],
        ["styles", "header"],
    ] as const)("wordReadTypedTargets reads the glossary's part through styles alone, %s first", async (first, second) => {
        const typed = await wordReadTypedTargets(ValidationContext.fromEntries(dualGlossary(first, second)));
        expect(typed.filter((target) => target.part === "word/glossary/styles2.xml")).toEqual([
            { part: "word/glossary/styles2.xml", segment: "styles", via: "glossary", root: "word/glossary/document.xml" },
        ]);
    });

    it.each([
        ["footer", "numbering"],
        ["numbering", "footer"],
    ] as const)("wordReadTypedTargets keeps the lazy segment when the document uses it, %s first", async (first, second) => {
        // A footerReference to the lazy relationship's r:id makes that typed
        // read real: both segments are read, in either order.
        const footerId = first === "footer" ? "rIdA" : "rIdB";
        const parts: Parts = { ...dualRelated(first, second), "word/document.xml": documentUsingFooter(footerId) };
        const typed = await wordReadTypedTargets(ValidationContext.fromEntries(parts));
        expect(typed.filter((target) => target.part === "word/numbering.xml")).toEqual([
            { part: "word/numbering.xml", segment: first },
            { part: "word/numbering.xml", segment: second },
        ]);
    });
});

describe("wordReferenceLoadedSources (#645, #647)", () => {
    // Session 36 (s36b/s36d): the `chart` relationship type is observed for
    // `rels-missing-sidecar` — Word loads a chart only when a Word-read
    // part's XML references the relationship's Id. The sources this returns
    // are where such a reference could live; every other shape yields
    // `undefined`, keeping the reach-based verdict.
    it.each([
        ["under word/", "word/charts/chart1.xml", "charts/chart1.xml"],
        ["outside word/", "charts/chart1.xml", "../charts/chart1.xml"],
    ] as const)("names the source relationships of a chart reached %s", async (_label, part, target) => {
        const parts = withDocRel({ ...MINIMAL_PARTS, [part]: BARE_CHART_XML }, "rId9", `${R}/chart`, target);
        expect(await wordReferenceLoadedSources(ValidationContext.fromEntries(parts), part)).toEqual([
            { source: "word/document.xml", id: "rId9", reference: "declared" },
        ]);
    });

    it.each([
        ["an unobserved relationship", "audio", `${R}/audio`],
        ["an observed relationship", "styles", `${R}/styles`],
        ["an unknown-type relationship", "urn:x", "urn:x:unknown"],
    ] as const)("yields undefined when %s also reaches the part", async (_label, _id, type) => {
        const parts = withDocRel(
            withDocRel({ ...MINIMAL_PARTS, "charts/chart1.xml": BARE_CHART_XML }, "rId9", `${R}/chart`, "../charts/chart1.xml"),
            "rId8",
            type,
            "../charts/chart1.xml",
        );
        expect(await wordReferenceLoadedSources(ValidationContext.fromEntries(parts), "charts/chart1.xml")).toBe(undefined);
    });

    it("yields undefined for a chart named by nothing or by an external relationship", async () => {
        const unrelated = { ...MINIMAL_PARTS, "charts/chart1.xml": BARE_CHART_XML };
        expect(await wordReferenceLoadedSources(ValidationContext.fromEntries(unrelated), "charts/chart1.xml")).toBe(undefined);
        const parts = withDocRel({ ...MINIMAL_PARTS, "charts/chart1.xml": BARE_CHART_XML }, "rId9", `${R}/chart`, "../charts/chart1.xml");
        parts["word/_rels/document.xml.rels"] = (parts["word/_rels/document.xml.rels"] as string).replace(
            '<Relationship Id="rId9"',
            '<Relationship Id="rId9" TargetMode="External"',
        );
        expect(await wordReferenceLoadedSources(ValidationContext.fromEntries(parts), "charts/chart1.xml")).toBe(undefined);
    });

    it("names each chart relationship, with the package part sourceless", async () => {
        const parts: Parts = {
            ...MINIMAL_PARTS,
            "charts/chart1.xml": BARE_CHART_XML,
            "word/_rels/styles.xml.rels":
                `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n` +
                `<Relationships xmlns="${RELS_NS}">` +
                `<Relationship Id="rId3" Type="${R}/chart" Target="../charts/chart1.xml"/>` +
                `</Relationships>`,
            "_rels/.rels": (MINIMAL_PARTS["_rels/.rels"] as string).replace(
                "</Relationships>",
                `<Relationship Id="rId2" Type="${R}/chart" Target="charts/chart1.xml"/></Relationships>`,
            ),
        };
        parts["word/_rels/document.xml.rels"] = (parts["word/_rels/document.xml.rels"] as string).replace(
            "</Relationships>",
            `<Relationship Id="rId9" Type="${R}/chart" Target="../charts/chart1.xml"/></Relationships>`,
        );
        expect(await wordReferenceLoadedSources(ValidationContext.fromEntries(parts), "charts/chart1.xml")).toEqual([
            { source: undefined, id: "rId2", reference: "declared" },
            { source: "word/document.xml", id: "rId9", reference: "declared" },
            { source: "word/styles.xml", id: "rId3", reference: "declared" },
        ]);
    });

    it("compares target names as OPC part names, and yields nothing without a package relationships part", async () => {
        const parts = withDocRel({ ...MINIMAL_PARTS, "charts/chart1.xml": BARE_CHART_XML }, "rId9", `${R}/chart`, "../CHARTS/chart1.xml");
        expect(await wordReferenceLoadedSources(ValidationContext.fromEntries(parts), "CHARTS/chart1.xml")).toEqual([
            { source: "word/document.xml", id: "rId9", reference: "declared" },
        ]);
        const { "_rels/.rels": _dropped, ...noMain } = parts;
        expect(await wordReferenceLoadedSources(ValidationContext.fromEntries(noMain), "charts/chart1.xml")).toBe(undefined);
    });

    // Session 44 (s44a–s44c, #647): an image relationship never loads its
    // target's relationships, so it is reached but never a source; an
    // unreferenced diagramData, oleObject or hdphoto reach opened clean, and
    // their referenced shapes are matched on any attribute but a picture's
    // a:blip r:embed, which session 45 opened clean (#702 S1).
    it("never names an image relationship, Transitional or Strict (#706)", async () => {
        for (const type of [`${R}/image`, "http://purl.oclc.org/ooxml/officeDocument/relationships/image"]) {
            const parts = withDocRel({ ...MINIMAL_PARTS, "charts/chart1.xml": BARE_CHART_XML }, "rId8", type, "../charts/chart1.xml");
            expect(await wordReferenceLoadedSources(ValidationContext.fromEntries(parts), "charts/chart1.xml")).toEqual([]);
        }
    });

    it("drops an image relationship beside a chart one, so the chart rule alone answers", async () => {
        const parts = withDocRel(
            withDocRel({ ...MINIMAL_PARTS, "charts/chart1.xml": BARE_CHART_XML }, "rId9", `${R}/chart`, "../charts/chart1.xml"),
            "rId8",
            `${R}/image`,
            "../charts/chart1.xml",
        );
        expect(await wordReferenceLoadedSources(ValidationContext.fromEntries(parts), "charts/chart1.xml")).toEqual([
            { source: "word/document.xml", id: "rId9", reference: "declared" },
        ]);
    });

    it.each([`${R}/diagramData`, `${R}/oleObject`, "http://schemas.microsoft.com/office/2007/relationships/hdphoto"])(
        "matches any attribute but a picture's a:blip r:embed for a %s reach",
        async (type) => {
            const parts = withDocRel({ ...MINIMAL_PARTS, "word/charts/chart1.xml": BARE_CHART_XML }, "rId8", type, "charts/chart1.xml");
            expect(await wordReferenceLoadedSources(ValidationContext.fromEntries(parts), "word/charts/chart1.xml")).toEqual([
                { source: "word/document.xml", id: "rId8", reference: "non-picture" },
            ]);
        },
    );

    it("yields undefined for an image reach beside an unobserved type", async () => {
        const parts = withDocRel(
            withDocRel({ ...MINIMAL_PARTS, "charts/chart1.xml": BARE_CHART_XML }, "rId8", `${R}/image`, "../charts/chart1.xml"),
            "rId7",
            `${R}/audio`,
            "../charts/chart1.xml",
        );
        expect(await wordReferenceLoadedSources(ValidationContext.fromEntries(parts), "charts/chart1.xml")).toBe(undefined);
    });

    // #781: the sidecar's `_rels` folder segment compares canonically — the
    // same equivalence the sidecar lookup finds the part by — so a sidecar
    // stored under an escaped spelling still names the source part whose
    // reference loads the reach.
    it.each(["_rels", "_RELS", "_%72els", "%5Frels", "_r%65ls"] as const)(
        "names the source part of a chart a sidecar under %s reaches",
        async (folder) => {
            const { "word/_rels/document.xml.rels": sidecar, ...rest } = withDocRel(
                { ...MINIMAL_PARTS, "charts/chart1.xml": BARE_CHART_XML },
                "rId9",
                `${R}/chart`,
                "../charts/chart1.xml",
            );
            const parts: Parts = { ...rest, [`word/${folder}/document.xml.rels`]: sidecar! };
            expect(await wordReferenceLoadedSources(ValidationContext.fromEntries(parts), "charts/chart1.xml")).toEqual([
                { source: "word/document.xml", id: "rId9", reference: "declared" },
            ]);
        },
    );

    // #781: a `source` of `undefined` cannot credit the Word-read part's
    // reference, so the reach is recorded "unreferenced" and a full
    // word-valid run demotes its `rels-missing-sidecar` to a warning the
    // `_rels` and `_RELS` spellings keep blocking.
    it.each(["_rels", "_RELS", "_%72els", "%5Frels", "_r%65ls"] as const)(
        "keeps a referenced no-sidecar chart a sidecar under %s reaches word-blocking",
        { timeout: 60_000 },
        async (folder) => {
            const { "word/_rels/document.xml.rels": sidecar, ...rest } = referencedChartSpecimenParts(
                "charts/chart1.xml",
                "../charts/chart1.xml",
                true,
            );
            const parts: Parts = { ...rest, [`word/${folder}/document.xml.rels`]: sidecar! };
            await withTempDir(async (dir) => {
                const file = path.join(dir, "specimen.docx");
                writeFileSync(file, await buildDocx(parts));
                const result = await validate(file, { profile: "word-valid" });
                expect(
                    result.issues.find((issue) => issue.code === "rels-missing-sidecar" && issue.path === "charts/chart1.xml")?.severity,
                ).toBe("error");
            });
        },
    );
});

describe("wordExactTypeTargets (#690)", () => {
    // The Microsoft comment-thread types carry no Office relationship-type
    // segment, so the typed walks never reach their targets. This walk reads
    // the main documents' relationships parts by exact Type, so a reached
    // glossary's comment-thread parts never stand in for the main document's.
    const COMMENTS_IDS = "http://schemas.microsoft.com/office/2016/09/relationships/commentsIds";
    const IDS_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w16cid:commentsIds xmlns:w16cid="http://schemas.microsoft.com/office/word/2016/wordml/cid"/>`;
    const withGlossaryIds = (parts: Parts): Parts => ({
        ...parts,
        "word/glossary/document.xml": glossaryWith("", ""),
        "word/glossary/_rels/document.xml.rels": relsPart(`<Relationship Id="rIdI" Type="${COMMENTS_IDS}" Target="commentsIds.xml"/>`),
        "word/glossary/commentsIds.xml": IDS_XML,
        "word/_rels/document.xml.rels": (parts["word/_rels/document.xml.rels"] as string).replace(
            "</Relationships>",
            `<Relationship Id="rIdG" Type="${R}/glossaryDocument" Target="glossary/document.xml"/></Relationships>`,
        ),
    });

    it("names the main document's target and never a reached glossary's", async () => {
        const glossaryOnly = withGlossaryIds({ ...MINIMAL_PARTS });
        expect(await wordExactTypeTargets(ValidationContext.fromEntries(glossaryOnly), COMMENTS_IDS)).toEqual([]);
        const both = withDocRel(
            withGlossaryIds({ ...MINIMAL_PARTS, "word/commentsIds.xml": IDS_XML }),
            "rId9",
            COMMENTS_IDS,
            "commentsIds.xml",
        );
        expect(await wordExactTypeTargets(ValidationContext.fromEntries(both), COMMENTS_IDS)).toEqual(["word/commentsIds.xml"]);
    });

    it("resolves against the main document's folder and takes the OPC-equivalent stored entry, once", async () => {
        const parts = withDocRel(
            withDocRel({ ...MINIMAL_PARTS, "word/commentsIds.xml": IDS_XML }, "rId9", COMMENTS_IDS, "CommentsIds.xml"),
            "rId8",
            COMMENTS_IDS,
            "/word/commentsIds.xml",
        );
        expect(await wordExactTypeTargets(ValidationContext.fromEntries(parts), COMMENTS_IDS)).toEqual(["word/commentsIds.xml"]);
    });

    it.each([
        ["an external target", (rel: string) => rel.replace('<Relationship Id="rId9"', '<Relationship Id="rId9" TargetMode="External"')],
        ["a Type with a trailing slash", (rel: string) => rel.replace(`Type="${COMMENTS_IDS}"`, `Type="${COMMENTS_IDS}/"`)],
        ["a Type in another case", (rel: string) => rel.replace(`Type="${COMMENTS_IDS}"`, `Type="${COMMENTS_IDS.toUpperCase()}"`)],
        ["a target that names no stored part", (rel: string) => rel.replace('Target="commentsIds.xml"', 'Target="missing.xml"')],
    ] as const)("skips %s", async (_label, edit) => {
        const parts = withDocRel({ ...MINIMAL_PARTS, "word/commentsIds.xml": IDS_XML }, "rId9", COMMENTS_IDS, "commentsIds.xml");
        parts["word/_rels/document.xml.rels"] = edit(parts["word/_rels/document.xml.rels"] as string);
        expect(await wordExactTypeTargets(ValidationContext.fromEntries(parts), COMMENTS_IDS)).toEqual([]);
    });

    it("returns nothing without a package relationships part", async () => {
        const parts: Parts = withDocRel({ ...MINIMAL_PARTS, "word/commentsIds.xml": IDS_XML }, "rId9", COMMENTS_IDS, "commentsIds.xml");
        delete parts["_rels/.rels"];
        expect(await wordExactTypeTargets(ValidationContext.fromEntries(parts), COMMENTS_IDS)).toEqual([]);
    });
});

describe("wordUnobservedRelationshipTargets (#638)", () => {
    // rels-missing-sidecar blocked in every part before #638. Word was never
    // shown a part that an unobserved Office relationship type names outside
    // word/, so those keep blocking for that finding; wordReadParts keeps the
    // word/-folder rule for XML syntax (#437). The package's officeDocument
    // relationship is no eager or lazy type either, so the main document is
    // named too; it is a story root, which blocks anyway.
    const MAIN_NAMED = canonicalPartName("word/document.xml");
    it.each([
        ["outside word/", "../charts/chart1.xml", "charts/chart1.xml"],
        ["under word/", "charts/chart1.xml", "word/charts/chart1.xml"],
    ] as const)("names an unobserved Office-type target %s", async (_label, target, part) => {
        const parts = withDocRel({ ...MINIMAL_PARTS, [part]: MALFORMED.xtra! }, "rId9", `${R}/chart`, target);
        expect(await wordUnobservedRelationshipTargets(ValidationContext.fromEntries(parts))).toEqual(
            new Set([MAIN_NAMED, canonicalPartName(part)]),
        );
    });

    it("names one from any part's relationships, in a Microsoft namespace too", async () => {
        const parts: Parts = {
            ...MINIMAL_PARTS,
            "word/_rels/styles.xml.rels":
                `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n` +
                `<Relationships xmlns="${RELS_NS}">` +
                `<Relationship Id="rId3" Type="http://schemas.microsoft.com/office/2011/relationships/people" Target="../xdocs/people1.xml"/>` +
                `</Relationships>`,
            "xdocs/people1.xml": MALFORMED.customXml!,
        };
        expect(await wordUnobservedRelationshipTargets(ValidationContext.fromEntries(parts))).toEqual(
            new Set([MAIN_NAMED, canonicalPartName("xdocs/people1.xml")]),
        );
    });

    it.each(["_%72els", "%5Frels", "_r%65ls"])(
        "resolves targets against the source part's folder from a word/%s sidecar (#772)",
        async (folder) => {
            const parts: Parts = {
                ...MINIMAL_PARTS,
                [`word/${folder}/styles.xml.rels`]:
                    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n` +
                    `<Relationships xmlns="${RELS_NS}">` +
                    `<Relationship Id="rId3" Type="http://schemas.microsoft.com/office/2011/relationships/people" Target="../xdocs/people1.xml"/>` +
                    `</Relationships>`,
                "xdocs/people1.xml": MALFORMED.customXml!,
            };
            expect(await wordUnobservedRelationshipTargets(ValidationContext.fromEntries(parts))).toEqual(
                new Set([MAIN_NAMED, canonicalPartName("xdocs/people1.xml")]),
            );
        },
    );

    it.each([
        ["an unknown type", "urn:x:unknown", "../xdocs/xtra.xml"],
        ["an observed type", `${R}/styles`, "../xdocs/xtra.xml"],
        ["an Office type with an extra path segment", `${R}/custom/styles`, "../xdocs/xtra.xml"],
    ] as const)("does not name the target of %s", async (_label, type, target) => {
        const parts = withDocRel({ ...MINIMAL_PARTS, "xdocs/xtra.xml": MALFORMED.xtra! }, "rId9", type, target);
        const named = await wordUnobservedRelationshipTargets(ValidationContext.fromEntries(parts));
        expect(named.has(canonicalPartName("xdocs/xtra.xml"))).toBe(type === `${R}/custom/styles`);
    });

    it("skips external targets and yields nothing without a main document", async () => {
        const parts = withDocRel({ ...MINIMAL_PARTS }, "rId9", `${R}/chart`, "https://example.com/chart1.xml");
        const external = {
            ...parts,
            "word/_rels/document.xml.rels": (parts["word/_rels/document.xml.rels"] as string).replace(
                'Target="https://example.com/chart1.xml"',
                'Target="https://example.com/chart1.xml" TargetMode="External"',
            ),
        };
        expect(await wordUnobservedRelationshipTargets(ValidationContext.fromEntries(external))).toEqual(new Set([MAIN_NAMED]));
        const { "_rels/.rels": _dropped, ...noMain } = withDocRel(
            { ...MINIMAL_PARTS, "charts/chart1.xml": MALFORMED.xtra! },
            "rId9",
            `${R}/chart`,
            "../charts/chart1.xml",
        );
        expect(await wordUnobservedRelationshipTargets(ValidationContext.fromEntries(noMain))).toEqual(new Set());
    });
});

describe("wordReadStoryParts (#1106, #1113)", () => {
    const HEADER_TYPE = "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml";
    const HEADER = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:hdr xmlns:w="${W}"><w:p/></w:hdr>`;

    /**
     * The minimal package plus word/header1.xml, which the main document
     * relates as rIdH and its w:sectPr references when `used`, and
     * word/header9.xml, which no relationship names (sessions 82 and 83).
     */
    const withHeaders = (used: boolean): Parts => ({
        ...MINIMAL_PARTS,
        "[Content_Types].xml": MINIMAL_PARTS["[Content_Types].xml"]!.replace(
            "</Types>",
            `<Override PartName="/word/header1.xml" ContentType="${HEADER_TYPE}"/><Override PartName="/word/header9.xml" ContentType="${HEADER_TYPE}"/></Types>`,
        ),
        "word/_rels/document.xml.rels": MINIMAL_PARTS["word/_rels/document.xml.rels"]!.replace(
            "</Relationships>",
            `<Relationship Id="rIdH" Type="${R}/header" Target="header1.xml"/></Relationships>`,
        ),
        "word/document.xml": used
            ? MINIMAL_PARTS["word/document.xml"]!.replace(
                  "<w:sectPr>",
                  `<w:sectPr><w:headerReference xmlns:r="${R}" w:type="default" r:id="rIdH"/>`,
              )
            : MINIMAL_PARTS["word/document.xml"]!,
        "word/header1.xml": HEADER,
        "word/header9.xml": HEADER,
    });

    /** `parts` with the main document stored as word/main.xml, every name pointing at it following. */
    const renamed = (parts: Parts): Parts => {
        const { "word/document.xml": main, "word/_rels/document.xml.rels": sidecar, ...rest } = parts;
        return {
            ...rest,
            "[Content_Types].xml": (rest["[Content_Types].xml"] as string).replace(
                'PartName="/word/document.xml"',
                'PartName="/word/main.xml"',
            ),
            "_rels/.rels": (rest["_rels/.rels"] as string).replace('Target="word/document.xml"', 'Target="word/main.xml"'),
            "word/main.xml": main!,
            "word/_rels/main.xml.rels": sidecar!,
        };
    };

    it("reads a main document stored as word/main.xml and no part no relationship names (sessions 82 and 83)", async () => {
        expect(await wordReadStoryParts(ValidationContext.fromEntries(renamed(withHeaders(false))))).toEqual(["word/main.xml"]);
    });

    it.each([
        [true, ["word/document.xml", "word/header1.xml"]],
        [false, ["word/document.xml"]],
    ] as const)("reads the related header only when a section references it (used: %s)", async (used, expected) => {
        expect(await wordReadStoryParts(ValidationContext.fromEntries(withHeaders(used)))).toEqual(expected);
    });

    // #1118 (review 1115 F1): the read model lists a part once per segment
    // that reaches it, so a header the body uses as its header and its footer
    // arrives twice; the helper lists it once.
    it("lists a part the body uses as header and footer once (#1118)", async () => {
        const parts = withHeaders(true);
        parts["word/_rels/document.xml.rels"] = (parts["word/_rels/document.xml.rels"] as string).replace(
            "</Relationships>",
            `<Relationship Id="rIdF" Type="${R}/footer" Target="header1.xml"/></Relationships>`,
        );
        parts["word/document.xml"] = (parts["word/document.xml"] as string).replace(
            'r:id="rIdH"/>',
            `r:id="rIdH"/><w:footerReference xmlns:r="${R}" w:type="default" r:id="rIdF"/>`,
        );
        const ctx = ValidationContext.fromEntries(parts);
        const segments = (await wordReadTypedTargets(ctx))
            .filter((target) => target.part === "word/header1.xml")
            .map((target) => target.segment);
        expect(segments).toEqual(["header", "footer"]);
        expect(await wordReadStoryParts(ctx)).toEqual(["word/document.xml", "word/header1.xml"]);
    });

    it("lists the parts in the context's part-name order, as the by-name list does, not in read order", async () => {
        expect(await wordReadStoryParts(ValidationContext.fromEntries(renamed(withHeaders(true))))).toEqual([
            "word/header1.xml",
            "word/main.xml",
        ]);
    });

    it("falls back to the parts named like a story when the officeDocument relationship names no stored part (#816)", async () => {
        const parts = withHeaders(true);
        parts["_rels/.rels"] = (parts["_rels/.rels"] as string).replace('Target="word/document.xml"', 'Target="word/missing.xml"');
        const ctx = ValidationContext.fromEntries(parts);
        expect(await wordReadStoryParts(ctx)).toEqual(ctx.storyParts());
        expect(ctx.storyParts()).toEqual(["word/document.xml", "word/header1.xml", "word/header9.xml"]);
    });

    it("computes the list once per context and profile", async () => {
        const ctx = ValidationContext.fromEntries(withHeaders(true));
        expect(wordReadStoryParts(ctx, "word-valid")).toBe(wordReadStoryParts(ctx, "word-valid"));
        expect(wordReadStoryParts(ctx, "strict")).not.toBe(wordReadStoryParts(ctx, "word-valid"));
    });

    // #1128: the scheduler hands each check gated on xml-syntax a new
    // withUnparseable view, so a per-view cache recomputed the list for every
    // story rule.
    it("computes the list once for every view with the same unparseable parts (#1128)", async () => {
        const ctx = ValidationContext.fromEntries(withHeaders(true));
        const gate = (): ValidationContext => ctx.withUnparseable(new Map([["word/header9.xml", "malformed"]]));
        expect(wordReadStoryParts(gate(), "word-valid")).toBe(wordReadStoryParts(gate(), "word-valid"));
        expect(wordReadStoryParts(ctx.withUnparseable(new Map()), "word-valid")).toBe(wordReadStoryParts(ctx, "word-valid"));
    });

    // A main document the view cannot parse leaves lazy usage unknown (#536),
    // so that view keeps the related header the base drops.
    it("computes its own list for a view with other unparseable parts (#1128)", async () => {
        const ctx = ValidationContext.fromEntries(withHeaders(false));
        expect(await wordReadStoryParts(ctx, "word-valid")).toEqual(["word/document.xml"]);
        const view = ctx.withUnparseable(new Map([["word/document.xml", "malformed"]]));
        expect(await wordReadStoryParts(view, "word-valid")).toEqual(["word/document.xml", "word/header1.xml"]);
    });
});

describe("wordPossiblyReadParts (#1116, #1125)", () => {
    const HEADER_TYPE = "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml";
    const HEADER = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:hdr xmlns:w="${W}"><w:p/></w:hdr>`;
    const MC = "http://schemas.openxmlformats.org/markup-compatibility/2006";
    const STRAY = "word/_rels/stray.xml.rels";
    const relationships = (inner: string): string =>
        `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="${RELS_NS}">${inner}</Relationships>`;
    const header = (id: string, target: string, mode = ""): string =>
        `<Relationship Id="${id}" Type="${R}/header" Target="${target}"${mode}/>`;

    /**
     * The minimal package plus word/header1.xml, which the main document
     * relates, and word/header9.xml, which no relationship names (session 83's
     * word/header83.xml), with `extra` parts set over it.
     */
    const withHeaders = (extra: Parts = {}): Parts => ({
        ...MINIMAL_PARTS,
        "[Content_Types].xml": MINIMAL_PARTS["[Content_Types].xml"]!.replace(
            "</Types>",
            `<Override PartName="/word/header1.xml" ContentType="${HEADER_TYPE}"/><Override PartName="/word/header9.xml" ContentType="${HEADER_TYPE}"/></Types>`,
        ),
        "word/_rels/document.xml.rels": MINIMAL_PARTS["word/_rels/document.xml.rels"]!.replace(
            "</Relationships>",
            `${header("rIdH", "header1.xml")}</Relationships>`,
        ),
        "word/header1.xml": HEADER,
        "word/header9.xml": HEADER,
        ...extra,
    });

    const named = async (parts: Parts): Promise<readonly string[] | undefined> => {
        const set = await wordPossiblyReadParts(ValidationContext.fromEntries(parts), "word-valid");
        return set === undefined ? undefined : [...set].sort();
    };
    const HEADER1 = canonicalPartName("word/header1.xml");
    const HEADER9 = canonicalPartName("word/header9.xml");
    const READ = ["word/_rels/document.xml.rels", "word/document.xml", "word/styles.xml"];

    it("lists every part Word reads or a relationship names, canonically, and no part nothing names", async () => {
        expect(await named(withHeaders({ "word/document.xml": documentUsingHeader("rIdH") }))).toEqual(
            [...READ, "word/header1.xml"].map(canonicalPartName).sort(),
        );
    });

    // #1125 (session 84 s84a mu-unused-header): Word does not read a header
    // the main document relates but no section references, and opened an
    // undeclared mc:MustUnderstand there clean.
    it("does not list a header the main document relates but does not use (#1125)", async () => {
        expect(await named(withHeaders())).toEqual(READ.map(canonicalPartName).sort());
    });

    const USING: Readonly<Record<StorySegment, (id: string) => string>> = {
        header: documentUsingHeader,
        footer: documentUsingFooter,
        comments: documentUsingComment,
        footnotes: documentUsingFootnote,
        endnotes: documentUsingEndnote,
    };

    it.each(STORY_SEGMENTS.flatMap((segment) => [true, false].map((used) => [segment, used] as const)))(
        "lists a %s part the main document relates only when it uses it (used: %s) (#1125)",
        async (segment, used) => {
            const story = canonicalPartName(`word/${STORY_FILE[segment]}`);
            const parts = withDocRel(
                { ...MINIMAL_PARTS, [`word/${STORY_FILE[segment]}`]: WELLFORMED_STORY[segment] },
                "rIdS",
                `${R}/${segment}`,
                STORY_FILE[segment],
            );
            if (used) {
                parts["word/document.xml"] = USING[segment]("rIdS");
            }
            const list = await named(parts);
            if (used) {
                expect(list).toContain(story);
            } else {
                expect(list).not.toContain(story);
            }
        },
    );

    // #1145: the read model decides a root's comments and notes use from the
    // root alone, so a read header or footer holding a reference of that kind,
    // or failing the byte-level XML checks, keeps the related part listed.
    const NOTE_SEGMENTS = ["comments", "footnotes", "endnotes"] as const;
    const NOTE_REFERENCE: Readonly<Record<(typeof NOTE_SEGMENTS)[number], string>> = {
        comments: '<w:commentReference w:id="0"/>',
        footnotes: '<w:footnoteReference w:id="1"/>',
        endnotes: '<w:endnoteReference w:id="1"/>',
    };
    const HOLDER_ROOT = { header: "hdr", footer: "ftr" } as const;
    /** The main document relating a `segment` part it does not use, and a `holder` part with `content`, used or not. */
    const withStoryHolding = (
        segment: (typeof NOTE_SEGMENTS)[number],
        holder: "header" | "footer",
        content: string,
        used: boolean,
    ): Parts => {
        const parts = withDocRel(
            withDocRel(
                { ...MINIMAL_PARTS, [`word/${STORY_FILE[segment]}`]: WELLFORMED_STORY[segment], [`word/${STORY_FILE[holder]}`]: content },
                "rIdS",
                `${R}/${segment}`,
                STORY_FILE[segment],
            ),
            "rIdH",
            `${R}/${holder}`,
            STORY_FILE[holder],
        );
        if (used) {
            parts["word/document.xml"] = USING[holder]("rIdH");
        }
        return parts;
    };
    const holding = (holder: "header" | "footer", inner: string): string =>
        `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:${HOLDER_ROOT[holder]} xmlns:w="${W}"><w:p><w:r>${inner}</w:r></w:p></w:${HOLDER_ROOT[holder]}>`;

    it.each(NOTE_SEGMENTS.flatMap((segment) => (["header", "footer"] as const).map((holder) => [segment, holder] as const)))(
        "lists a related %s part the main document does not use when a used %s references it (#1145)",
        async (segment, holder) => {
            const story = canonicalPartName(`word/${STORY_FILE[segment]}`);
            expect(await named(withStoryHolding(segment, holder, holding(holder, NOTE_REFERENCE[segment]), true))).toContain(story);
            expect(await named(withStoryHolding(segment, holder, holding(holder, NOTE_REFERENCE[segment]), false))).not.toContain(story);
            expect(await named(withStoryHolding(segment, holder, holding(holder, "<w:t>No reference</w:t>"), true))).not.toContain(story);
        },
    );

    it.each(NOTE_SEGMENTS)(
        "lists a related %s part the main document does not use when a used header fails the byte-level XML checks (#1145)",
        async (segment) => {
            const story = canonicalPartName(`word/${STORY_FILE[segment]}`);
            const unclosed = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:hdr xmlns:w="${W}"><w:p>`;
            expect(await named(withStoryHolding(segment, "header", unclosed, true))).toContain(story);
        },
    );

    it("does not list a related endnotes part when a used header holds only a comment reference (#1145)", async () => {
        const parts = withStoryHolding("endnotes", "header", holding("header", NOTE_REFERENCE.comments), true);
        const story = canonicalPartName(`word/${STORY_FILE.endnotes}`);
        expect(await named(parts)).not.toContain(story);
    });

    it("lists an unused related header another relationships part also names (#1125)", async () => {
        expect(await named(withHeaders({ [STRAY]: relationships(header("rId1", "header1.xml")) }))).toContain(HEADER1);
    });

    it("lists an unused related header whose relationship is not a direct child of the sidecar (#1125)", async () => {
        const parts = withHeaders();
        parts["word/_rels/document.xml.rels"] = MINIMAL_PARTS["word/_rels/document.xml.rels"]!.replace(
            "</Relationships>",
            `<mc:AlternateContent xmlns:mc="${MC}"><mc:Choice xmlns:foo="urn:foo" Requires="foo">${header("rIdH", "header1.xml")}</mc:Choice></mc:AlternateContent></Relationships>`,
        );
        expect(await named(parts)).toContain(HEADER1);
    });

    it.each([
        ["reads eagerly", "styles"],
        ["does not read at open", "image"],
    ])("lists a header the main document relates through a type Word %s (#1125)", async (_label, type) => {
        const parts = withDocRel(withHeaders(), "rIdI", `${R}/${type}`, "header9.xml");
        expect(await named(parts)).toContain(HEADER9);
    });

    it.each(BYTE_REJECTED_MAIN_PARTS)(
        "lists an unused related $segment part when the main fails the byte-level XML checks (#1125)",
        async ({ part, parts }) => {
            expect(await named(parts)).toContain(canonicalPartName(part));
        },
    );

    it("does not list a header a reached glossary relates but does not use, and lists it when the glossary is unreached (#1125)", async () => {
        const story = canonicalPartName("word/glossary/header9.xml");
        const reached = storyViaGlossary("header", WELLFORMED_STORY.header, glossaryWith("", ""));
        expect(await named(reached)).toContain(canonicalPartName("word/glossary/document.xml"));
        expect(await named(reached)).not.toContain(story);
        expect(await named({ ...reached, "word/_rels/document.xml.rels": MINIMAL_PARTS["word/_rels/document.xml.rels"]! })).toContain(
            story,
        );
        expect(await named(storyViaGlossary("header", WELLFORMED_STORY.header))).toContain(story);
    });

    // #1147 (review 1138 F7): the exception covers the lazy story types only.
    // A reached glossary's customXml relationship, which the read model does
    // not follow (session 20), still names its part.
    it("lists the part a reached glossary relates as customXml", async () => {
        const parts = storyViaGlossary(
            "header",
            WELLFORMED_STORY.header,
            glossaryWith("", ""),
            relsPart(`<Relationship Id="rIdC" Type="${R}/customXml" Target="item1.xml"/>`),
        );
        parts["word/glossary/item1.xml"] = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<root/>`;
        expect(await named(parts)).toContain(canonicalPartName("word/glossary/item1.xml"));
    });

    // The read model reaches the main document through the Fallback Word
    // selects in the package relationships part (#634), so the main
    // document's sidecar is a root's.
    it("does not list an unused related header when the officeDocument relationship sits in a selected Fallback (#634, #1125)", async () => {
        const parts = wrapRelationship(withHeaders() as Record<string, string>, PACKAGE_RELS, "rId1", fallback);
        expect(await named(parts)).toEqual(READ.map(canonicalPartName).sort());
    });

    it("lists an unused related header a relationships part names that the main document reads as another type (#1125)", async () => {
        const parts = withDocRel(
            withHeaders({ [STRAY]: relationships(header("rId1", "header1.xml")) }),
            "rIdS",
            `${R}/styles`,
            "_rels/stray.xml.rels",
        );
        expect(await named(parts)).toContain(HEADER1);
    });

    // Word-valid's file-reference walk resolves every internal scheme but
    // http:, https: and mailto: as a part (#988, #1008); the read model
    // follows none, so a story root's own relationship to one still counts.
    it.each(["ftp:header9.xml", "httpfoo:header9.xml", "FTP:header9.xml"])(
        "lists the part a scheme-like target %s names",
        async (target) => {
            const stored = `word/${target}`;
            const list = await named(withHeaders({ [stored]: HEADER, [STRAY]: relationships(header("rId9", target)) }));
            expect(list).toContain(canonicalPartName(stored));
        },
    );

    it.each(["http:header9.xml", "HTTPS:header9.xml", "mailto:header9.xml"])(
        "does not list the part a target %s spells",
        async (target) => {
            const stored = `word/${target}`;
            const list = await named(withHeaders({ [stored]: HEADER, [STRAY]: relationships(header("rId9", target)) }));
            expect(list).not.toContain(canonicalPartName(stored));
        },
    );

    it("lists an unused header the main document relates through a scheme-like target the read model does not follow", async () => {
        const parts = withHeaders({ "word/ftp:header1.xml": HEADER });
        parts["word/_rels/document.xml.rels"] = (parts["word/_rels/document.xml.rels"] as string).replace(
            'Target="header1.xml"',
            'Target="ftp:header1.xml"',
        );
        expect(await named(parts)).toContain(canonicalPartName("word/ftp:header1.xml"));
    });

    it("lists a part only a relationships part with no stored source names", async () => {
        expect(await named(withHeaders({ [STRAY]: relationships(header("rId9", "header9.xml")) }))).toContain(HEADER9);
    });

    // #1135: the read model reads only the package-relationships namespace's
    // Relationship children of a Relationships root in that namespace. Any
    // other Relationship element names its part, used or not, so a reading
    // the model cannot make never narrows the list (review 1126 F7).
    it.each([
        ["in no namespace", (rels: string) => rels.replace('Target="header1.xml"/>', 'Target="header1.xml" xmlns=""/>')],
        [
            "under a root outside the package-relationships namespace",
            (rels: string) => rels.replace(`<Relationships xmlns="${RELS_NS}">`, '<Relationships xmlns="urn:not-pr">'),
        ],
        [
            "in the package-relationships namespace under a root outside it",
            (rels: string) =>
                rels
                    .replace(`<Relationships xmlns="${RELS_NS}">`, '<Relationships xmlns="urn:not-pr">')
                    .replace('<Relationship Id="rIdH"', `<Relationship xmlns="${RELS_NS}" Id="rIdH"`),
        ],
        [
            "in a foreign namespace under the package-relationships root",
            (rels: string) => rels.replace('<Relationship Id="rIdH"', '<x:Relationship xmlns:x="urn:x" Id="rIdH"'),
        ],
    ])("lists a header the main document's relationships part names %s, used or not (#1135)", async (_label, spell) => {
        for (const document of [documentUsingHeader("rIdH"), MINIMAL_PARTS["word/document.xml"]!]) {
            const parts = withHeaders({ "word/document.xml": document });
            parts["word/_rels/document.xml.rels"] = spell(parts["word/_rels/document.xml.rels"] as string);
            expect(parts["word/_rels/document.xml.rels"]).not.toBe(withHeaders()["word/_rels/document.xml.rels"]);
            expect(await named(parts)).toContain(HEADER1);
        }
    });

    it.each([
        ["an escaped spelling", STRAY, "header%39.xml"],
        ["a climb past the package root", STRAY, "../../word/header9.xml"],
        ["a package-absolute target", STRAY, "/word/header9.xml"],
        // #1134: a target resolves with its query or fragment dropped (#361).
        ["a target with a fragment", STRAY, "header9.xml#frag"],
        ["a target with a query", STRAY, "header9.xml?v=1"],
    ])("lists a part %s names", async (_label, part, target) => {
        expect(await named(withHeaders({ [part]: relationships(header("rId9", target)) }))).toContain(HEADER9);
    });

    // Conservative: a relationship Word's markup-compatibility view drops
    // still names its part, so no Choice can turn a finding into a warning.
    it("lists a part only an unselected Choice of a relationships part names", async () => {
        const choice =
            `<mc:AlternateContent xmlns:mc="${MC}"><mc:Choice xmlns:foo="urn:foo" Requires="foo">` +
            `${header("rId9", "header9.xml")}</mc:Choice></mc:AlternateContent>`;
        expect(await named(withHeaders({ [STRAY]: relationships(choice) }))).toContain(HEADER9);
    });

    it("does not list a part only an External relationship names", async () => {
        const external = relationships(header("rId9", "header9.xml", ' TargetMode="External"'));
        expect(await named(withHeaders({ [STRAY]: external }))).not.toContain(HEADER9);
    });

    // #1134: only the exact spelling External is skipped. Whether Word reads
    // another spelling as External is unprobed, so the relationship stays
    // internal and the list errs wide.
    it.each([' TargetMode="external"', ' TargetMode=" External "'])("lists a part a relationship with%s names", async (mode) => {
        expect(await named(withHeaders({ [STRAY]: relationships(header("rId9", "header9.xml", mode)) }))).toContain(HEADER9);
    });

    // #1134: a relationship of any type names its part, an Office type or not,
    // in a stray relationships part and among the main document's own.
    it("lists a part only a relationship of a custom type names", async () => {
        const custom = '<Relationship Id="rId9" Type="urn:custom:hdr" Target="header9.xml"/>';
        expect(await named(withHeaders({ [STRAY]: relationships(custom) }))).toContain(HEADER9);
        const parts = withHeaders();
        parts["word/_rels/document.xml.rels"] = (parts["word/_rels/document.xml.rels"] as string).replace(
            "</Relationships>",
            `${custom}</Relationships>`,
        );
        expect(await named(parts)).toContain(HEADER9);
    });

    // The DOM parser decodes invalid UTF-8 with replacement characters and
    // keeps the relationship; the byte-level checks reject the part.
    it.each([
        [
            "an unclosed root",
            `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="${RELS_NS}">${header("rId9", "x.xml")}`,
        ],
        [
            "invalid UTF-8 the DOM parser decodes",
            new Uint8Array([
                ...new TextEncoder().encode(relationships(header("rId9", "x.xml")).replace("/></Relationships>", ' Note="')),
                0xff,
                0xfe,
                ...new TextEncoder().encode('"/></Relationships>'),
            ]),
        ],
    ] as const)("knows nothing when a relationships part fails the byte-level XML checks: %s", async (_label, stray) => {
        expect(await named(withHeaders({ [STRAY]: stray }))).toBeUndefined();
    });

    it("lists the canonical name of a part stored under another spelling", async () => {
        const stored = "word/Header7.xml";
        expect(canonicalPartName(stored)).not.toBe(stored);
        const list = await named(withHeaders({ [stored]: HEADER, [STRAY]: relationships(header("rId7", "Header7.xml")) }));
        expect(list).toContain(canonicalPartName(stored));
    });

    it("knows nothing without a package relationships part", async () => {
        const { "_rels/.rels": _packageRels, ...rest } = withHeaders();
        expect(await named(rest)).toBeUndefined();
    });

    it("knows nothing when the officeDocument relationship names no stored part (#816)", async () => {
        const parts = withHeaders();
        parts["_rels/.rels"] = (parts["_rels/.rels"] as string).replace('Target="word/document.xml"', 'Target="word/missing.xml"');
        expect(await named(parts)).toBeUndefined();
    });
});
