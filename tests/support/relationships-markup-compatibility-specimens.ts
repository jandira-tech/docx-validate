/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

import { NS } from "../../src/lib/types";
import { MINIMAL_PARTS } from "./minimal-docx";
import { STORY_FOOTER_XML, storyContentType, storyReferenceParts } from "./story-content-type-specimens";

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
 * Unobserved in Word: the same `mc:MustUnderstand` on the main document's
 * own sidecar root. Session 34 probed `mc:MustUnderstand` only on
 * `_rels/.rels` (s34c), so this shape keeps its word-valid warning.
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
