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
import { wordReadParts, wordTypedRelationshipTargets } from "../src/lib/word-read-parts";
import { validate } from "../src/scripts/office/validate";
import { buildDocx, MINIMAL_PARTS } from "./support/minimal-docx";

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
        expect([...reached].some((part) => part.includes("example.com"))).toBe(false);
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

/** One observed specimen: the malformed part, and the word-valid verdict Word 16.115 gave it. */
interface Specimen {
    readonly label: string;
    readonly parts: () => Parts;
    /** Where the XML error lands. */
    readonly part: string;
    /** Blocking under word-valid: Word offered recovery or refused the file. */
    readonly blocking: boolean;
    /** The XML diagnostic code the defect reports. */
    readonly code?: "xml-syntax" | "xml-encoding";
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
];

const FOOTER_UNUSED = SPECIMENS[2]!;
const CUSTOMXML = SPECIMENS[9]!;
const specimenByLabel = (label: string): Specimen => SPECIMENS.find((specimen) => specimen.label === label)!;
const USED_HEADER_OUTSIDE = specimenByLabel("a used header stored outside word/ (#516, open-error)");
const HEADER_UNUSED = specimenByLabel("an unused header under word/, related but not referenced (#516, clean-open)");

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

    it.each([FOOTER_UNUSED, CUSTOMXML, USED_HEADER_OUTSIDE, HEADER_UNUSED])(
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
