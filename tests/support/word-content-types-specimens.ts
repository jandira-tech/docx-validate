/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

import { MINIMAL_PARTS } from "./minimal-docx";

const STREAM = "[Content_Types].xml";
const TYPES = MINIMAL_PARTS[STREAM]!;
const CORE_TYPE = "application/vnd.openxmlformats-package.core-properties+xml";
const CORE = '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"/>';

interface Specimen {
    name: string;
    parts: Readonly<Record<string, string>>;
    rejected: boolean;
    /**
     * The specimen declares an `Override` whose `PartName` is not a valid
     * OPC part name (#414), so the package is invalid on that account too,
     * whatever the content-type scope.
     */
    malformedPartName?: boolean;
}

const append = (declarations: string): string => TYPES.replace("</Types>", `${declarations}</Types>`);
const override = (partName: string, contentType: string): string => `<Override PartName="${partName}" ContentType="${contentType}"/>`;

/** Single-defect packages shared by direct-check and full/partial regressions. */
export const contentTypeSpecimens: readonly Specimen[] = [
    ...[STREAM, "[CONTENT_TYPES].XML"].map((stream) => {
        const { [STREAM]: _stream, ...parts } = MINIMAL_PARTS;
        return {
            name: `missing Override target in ${stream}`,
            parts: { ...parts, [stream]: append(override("/word/x.xml", "invalid-type")) },
            rejected: false,
        };
    }),
    {
        name: "unused Default extension",
        parts: { ...MINIMAL_PARTS, [STREAM]: append('<Default Extension="foo" ContentType="invalid-type"/>') },
        rejected: false,
    },
    ...["/word/styles.xml", "/WORD/STYLES.XML", "/word/%73tyles.xml", "/word/document.xml"].flatMap((partName) => {
        const target = partName.includes("document") ? "document" : "styles";
        const contentType = `application/vnd.openxmlformats-officedocument.wordprocessingml.${target === "document" ? "document.main" : "styles"}+xml`;
        return [false, true].map((rejected) => ({
            name: `${rejected ? "invalid" : "valid"} Override for ${partName}`,
            parts: {
                ...MINIMAL_PARTS,
                [STREAM]: TYPES.replace(
                    override(`/word/${target}.xml`, contentType),
                    override(partName, `${rejected ? "invalid-" : ""}${contentType}`),
                ),
            },
            rejected,
            // `%73` escapes the unreserved `s`, so it is not a valid part name (#414).
            malformedPartName: partName === "/word/%73tyles.xml",
        }));
    }),
    ...[false, true].map((rejected) => ({
        name: `${rejected ? "invalid" : "valid"} Override for /docProps/core.xml`,
        parts: {
            ...MINIMAL_PARTS,
            "docProps/core.xml": CORE,
            [STREAM]: append(override("/docProps/core.xml", `${rejected ? "invalid-" : ""}${CORE_TYPE}`)),
        },
        rejected,
    })),
    ...["png", "PNG"].flatMap((extension) =>
        [false, true].map((rejected) => ({
            name: `${rejected ? "invalid" : "valid"} Default covering a.${extension}`,
            parts: {
                ...MINIMAL_PARTS,
                [`word/media/a.${extension}`]: "unreferenced image payload",
                [STREAM]: append(`<Default Extension="png" ContentType="${rejected ? "invalid-" : ""}image/png"/>`),
            },
            rejected,
        })),
    ),
    {
        name: "Default shadowed by every matching part's Override",
        parts: {
            ...MINIMAL_PARTS,
            "word/media/a.PNG": "first image payload",
            "word/media/b.png": "second image payload",
            [STREAM]: append(
                '<Default Extension="PNG" ContentType="invalid-image/png"/>' +
                    override("/WORD/MEDIA/A.PNG", "image/png") +
                    override("/word/media/%62.png", "image/png"),
            ),
        },
        rejected: false,
        // The escaped unreserved character is not a valid part name (#414).
        malformedPartName: true,
    },
    {
        name: "Default still covers one part when another has an Override",
        parts: {
            ...MINIMAL_PARTS,
            "word/media/a.PNG": "first image payload",
            "word/media/b.png": "second image payload",
            [STREAM]: append('<Default Extension="PNG" ContentType="invalid-image/png"/>' + override("/word/media/a.PNG", "image/png")),
        },
        rejected: true,
    },
    {
        name: "reserved escape does not name a stored part",
        parts: { ...MINIMAL_PARTS, [STREAM]: append(override("/word%2Fstyles.xml", "invalid-type")) },
        rejected: false,
        // The escaped separator is not a valid part name (#414).
        malformedPartName: true,
    },
];
