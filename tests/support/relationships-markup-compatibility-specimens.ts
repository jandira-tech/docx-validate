/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

import { NS } from "../../src/lib/types";
import { MINIMAL_PARTS } from "./minimal-docx";
import { STORY_FOOTER_XML, glossaryStoryReferenceParts, storyContentType, storyReferenceParts } from "./story-content-type-specimens";

export const RELATIONSHIPS_MCE_CODE = "opc-relationships-markup-compatibility";
export const PACKAGE_RELS = "_rels/.rels";
export const DOCUMENT_RELS = "word/_rels/document.xml.rels";
export const MCE_DECLARATIONS = `xmlns:mc="${NS.MC}" xmlns:foo="urn:foo" xmlns:x="urn:x" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"`;

export const fallback = (content: string): string =>
    `<mc:AlternateContent><mc:Choice Requires="foo"/><mc:Fallback>${content}</mc:Fallback></mc:AlternateContent>`;

/** Session 33 s33d/s33e: replace just the named relationship, retaining the other package bytes. */
export const wrapRelationship = (
    parts: Readonly<Record<string, string>>,
    part: string,
    id: string,
    wrap: (relationship: string) => string,
    declarations = MCE_DECLARATIONS,
): Record<string, string> => ({
    ...parts,
    [part]: parts[part]!.replace("<Relationships ", `<Relationships ${declarations} `).replace(
        new RegExp(`<Relationship Id="${id}"[^>]*/>`),
        wrap,
    ),
});

const footer = storyReferenceParts("footerReference", "footer", "footer1.xml", STORY_FOOTER_XML, storyContentType("footer"));

const MUST_UNDERSTAND_DECLARATIONS = `xmlns:mc="${NS.MC}" xmlns:foo="urn:foo" mc:MustUnderstand="foo"`;

const withRootDeclarations = (parts: Readonly<Record<string, string>>, part: string, declarations: string): Record<string, string> => ({
    ...parts,
    [part]: parts[part]!.replace("<Relationships ", `<Relationships ${declarations} `),
});

/**
 * Session 34 s34c (#636): `mc:MustUnderstand` of a namespace Word does not
 * understand on the `_rels/.rels` root. Word 16.115 offered recovery and,
 * declined, reported an open error; the control opened clean.
 */
export const packageRelsRootMustUnderstand = (): Record<string, string> =>
    withRootDeclarations(MINIMAL_PARTS, PACKAGE_RELS, MUST_UNDERSTAND_DECLARATIONS);

/**
 * Session 34 s34d (#636): session 26's footer relationship inside an
 * `mc:Choice` whose `Requires` prefix is undeclared, with an empty Fallback —
 * an `mce-malformed` AlternateContent. Word 16.115 offered recovery.
 */
export const documentRelsChoiceUndeclaredPrefix = (): Record<string, string> =>
    wrapRelationship(
        footer,
        DOCUMENT_RELS,
        "rIdX",
        (rel) => `<mc:AlternateContent><mc:Choice Requires="nope">${rel}</mc:Choice><mc:Fallback/></mc:AlternateContent>`,
    );

/**
 * Session 36 s36b (#644): Word 16.115 offered recovery for this
 * `mc:MustUnderstand` on the main document's own sidecar root.
 */
export const documentRelsRootMustUnderstand = (): Record<string, string> =>
    withRootDeclarations(footer, DOCUMENT_RELS, MUST_UNDERSTAND_DECLARATIONS);

/** The five relationship shapes observed in Word 16.115, without saved fixture files. */
export const RELATIONSHIPS_MCE_CASES = [
    {
        label: "mce-rels-fallback-footer",
        part: DOCUMENT_RELS,
        clean: true,
        parts: wrapRelationship(footer, DOCUMENT_RELS, "rIdX", fallback),
    },
    {
        label: "mce-rels-choice-w14-footer",
        part: DOCUMENT_RELS,
        clean: true,
        parts: wrapRelationship(
            footer,
            DOCUMENT_RELS,
            "rIdX",
            (rel) => `<mc:AlternateContent><mc:Choice Requires="w14">${rel}</mc:Choice><mc:Fallback/></mc:AlternateContent>`,
        ),
    },
    {
        label: "mce-rels-ignorable-foreign",
        part: DOCUMENT_RELS,
        clean: true,
        parts: wrapRelationship(footer, DOCUMENT_RELS, "rIdX", (rel) => `${rel}<x:Wrap/>`, `${MCE_DECLARATIONS} mc:Ignorable="x"`),
    },
    {
        label: "mce-package-rels-fallback-main",
        part: PACKAGE_RELS,
        clean: true,
        parts: wrapRelationship(MINIMAL_PARTS, PACKAGE_RELS, "rId1", fallback),
    },
    {
        label: "mce-rels-fallback-foreign",
        part: DOCUMENT_RELS,
        clean: false,
        parts: wrapRelationship(footer, DOCUMENT_RELS, "rIdX", (rel) => fallback(`${rel}<x:Wrap/>`)),
    },
] as const;

/**
 * Root-wrapped parts from #640. Word 16.115 offered recovery for the
 * `fallback` shape in `_rels/.rels` (session 36, s36c); the rest are unobserved.
 */
export const ROOT_MCE_SHAPES = ["fallback", "choice", "process-content"] as const;
export type RootMceShape = (typeof ROOT_MCE_SHAPES)[number];
export const rootMceXml = (shape: RootMceShape, content: string): string => {
    if (shape === "process-content") {
        return `<x:Wrap ${MCE_DECLARATIONS} mc:Ignorable="x" mc:ProcessContent="x:Wrap">${content}</x:Wrap>`;
    }
    const branches =
        shape === "choice"
            ? `<mc:Choice Requires="w14">${content}</mc:Choice><mc:Fallback/>`
            : `<mc:Choice Requires="foo"/><mc:Fallback>${content}</mc:Fallback>`;
    return `<mc:AlternateContent ${MCE_DECLARATIONS}>${branches}</mc:AlternateContent>`;
};

export const rootMceRelationshipsParts = (part: string, shape: RootMceShape): Record<string, string> => ({
    ...MINIMAL_PARTS,
    [part]: rootMceXml(
        shape,
        `<Relationship xmlns="${NS.PR}" Id="rId1" Type="${NS.R}/officeDocument" Target="${part === PACKAGE_RELS ? "word/document.xml" : "document.xml"}"/>` +
            `<Relationship xmlns="${NS.PR}" Id="rId2" Type="${NS.R}/styles" Target="${part === PACKAGE_RELS ? "word/styles.xml" : "styles.xml"}"/>`,
    ),
});

/** Session 36's malformed sibling selects nothing and removes no relationship. */
export const appendMalformedSibling = (rels: string): string =>
    rels
        .replace("<Relationships ", `<Relationships xmlns:mc="${NS.MC}" `)
        .replace(
            "</Relationships>",
            '<mc:AlternateContent><mc:Choice Requires="nope"/><mc:Fallback/></mc:AlternateContent></Relationships>',
        );

export const rootMustUnderstand = (rels: string): string =>
    rels.replace("<Relationships ", `<Relationships ${MUST_UNDERSTAND_DECLARATIONS} `);

/** Session 36 s36e: move the main document, its Override and its sidecar together. */
export const renamedMainParts = (
    sidecar: (rels: string) => string = (rels) => rels,
    stray?: (rels: string) => string,
): Record<string, string> => {
    const parts: Record<string, string> = {
        ...MINIMAL_PARTS,
        [PACKAGE_RELS]: MINIMAL_PARTS[PACKAGE_RELS]!.replace('Target="word/document.xml"', 'Target="word/doc2.xml"'),
        "[Content_Types].xml": MINIMAL_PARTS["[Content_Types].xml"]!.replace('PartName="/word/document.xml"', 'PartName="/word/doc2.xml"'),
        "word/doc2.xml": MINIMAL_PARTS["word/document.xml"]!,
        "word/_rels/doc2.xml.rels": sidecar(MINIMAL_PARTS[DOCUMENT_RELS]!),
    };
    delete parts["word/document.xml"];
    delete parts[DOCUMENT_RELS];
    if (stray) {
        parts["word/document.xml"] = MINIMAL_PARTS["word/document.xml"]!;
        parts[DOCUMENT_RELS] = stray(MINIMAL_PARTS[DOCUMENT_RELS]!);
    }
    return parts;
};

const glossary = (): Record<string, string> =>
    glossaryStoryReferenceParts("footerReference", "footer", "footer9.xml", STORY_FOOTER_XML, storyContentType("footer"));
const GLOSSARY_RELS = "word/glossary/_rels/document.xml.rels";
const emptyRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="${NS.PR}"></Relationships>`;
const atPart = (parts: Readonly<Record<string, string>>, part: string, transform: (rels: string) => string): Record<string, string> => ({
    ...parts,
    [part]: transform(parts[part] ?? emptyRels),
});

/** Session 36 #644 rows, plus session 34's two calibrated shapes. */
export const STORY_RELS_MCE_CASES = [
    {
        label: "docrels-mc-malformed-sibling",
        part: DOCUMENT_RELS,
        code: "mce-malformed",
        blocks: true,
        parts: () => atPart(MINIMAL_PARTS, DOCUMENT_RELS, appendMalformedSibling),
    },
    {
        label: "pkgrels-mc-malformed-sibling",
        part: PACKAGE_RELS,
        code: "mce-malformed",
        blocks: true,
        parts: () => atPart(MINIMAL_PARTS, PACKAGE_RELS, appendMalformedSibling),
    },
    {
        label: "glossary-rels-mc-malformed-sibling",
        part: GLOSSARY_RELS,
        code: "mce-malformed",
        blocks: true,
        parts: () => atPart(glossary(), GLOSSARY_RELS, appendMalformedSibling),
    },
    // s36a retained the unreached original main document; only its Override moved.
    {
        label: "stray-docrels-mc-malformed-sibling",
        part: DOCUMENT_RELS,
        code: "mce-malformed",
        blocks: false,
        parts: () => renamedMainParts(undefined, appendMalformedSibling),
    },
    {
        label: "docrels-root-mustunderstand",
        part: DOCUMENT_RELS,
        code: "mce-must-understand",
        blocks: true,
        parts: () => atPart(MINIMAL_PARTS, DOCUMENT_RELS, rootMustUnderstand),
    },
    {
        label: "styles-rels-mc-malformed-sibling",
        part: "word/_rels/styles.xml.rels",
        code: "mce-malformed",
        blocks: false,
        parts: () => atPart(MINIMAL_PARTS, "word/_rels/styles.xml.rels", appendMalformedSibling),
    },
    {
        label: "footer-rels-mc-malformed-sibling",
        part: "word/_rels/footer1.xml.rels",
        code: "mce-malformed",
        blocks: false,
        parts: () => atPart(footer, "word/_rels/footer1.xml.rels", appendMalformedSibling),
    },
    { label: "renamed-main-control", part: "word/_rels/doc2.xml.rels", code: undefined, blocks: false, parts: () => renamedMainParts() },
    {
        label: "renamed-main-rels-mc-malformed-sibling",
        part: "word/_rels/doc2.xml.rels",
        code: "mce-malformed",
        blocks: true,
        parts: () => renamedMainParts(appendMalformedSibling),
    },
    {
        label: "renamed-main-rels-root-mustunderstand",
        part: "word/_rels/doc2.xml.rels",
        code: "mce-must-understand",
        blocks: true,
        parts: () => renamedMainParts(rootMustUnderstand),
    },
    {
        label: "stray-docrels-root-mustunderstand",
        part: DOCUMENT_RELS,
        code: "mce-must-understand",
        blocks: false,
        parts: () => renamedMainParts(undefined, rootMustUnderstand),
    },
    {
        label: "glossary-rels-root-mustunderstand",
        part: GLOSSARY_RELS,
        code: "mce-must-understand",
        blocks: true,
        parts: () => atPart(glossary(), GLOSSARY_RELS, rootMustUnderstand),
    },
    {
        label: "packageRelsRootMustUnderstand",
        part: PACKAGE_RELS,
        code: "mce-must-understand",
        blocks: true,
        parts: packageRelsRootMustUnderstand,
    },
    {
        label: "documentRelsChoiceUndeclaredPrefix",
        part: DOCUMENT_RELS,
        code: "mce-malformed",
        blocks: true,
        parts: documentRelsChoiceUndeclaredPrefix,
    },
    {
        label: "pkgrels-relationship-mustunderstand",
        part: PACKAGE_RELS,
        code: "mce-must-understand",
        blocks: true,
        parts: () =>
            wrapRelationship(MINIMAL_PARTS, PACKAGE_RELS, "rId1", (rel) =>
                rel.replace("<Relationship ", '<Relationship mc:MustUnderstand="foo" '),
            ),
    },
    {
        label: "docrels-mc-fallback-only",
        part: DOCUMENT_RELS,
        code: "mce-malformed",
        blocks: true,
        parts: () =>
            wrapRelationship(
                MINIMAL_PARTS,
                DOCUMENT_RELS,
                "rId1",
                (rel) => `<mc:AlternateContent><mc:Fallback>${rel}</mc:Fallback></mc:AlternateContent>`,
            ),
    },
    {
        label: "docrels-mc-foreign-branch",
        part: DOCUMENT_RELS,
        code: "mce-malformed",
        blocks: true,
        parts: () =>
            wrapRelationship(
                MINIMAL_PARTS,
                DOCUMENT_RELS,
                "rId1",
                (rel) => `<mc:AlternateContent><foo:bar/><mc:Fallback>${rel}</mc:Fallback></mc:AlternateContent>`,
            ),
    },
] as const;
