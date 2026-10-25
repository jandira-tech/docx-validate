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
