/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

/**
 * Unit tests for the OPC content-type resolver (`src/lib/content-types.ts`).
 *
 * Pure-function level: [Content_Types].xml documents are inline XML strings
 * parsed through `parseXml` (never `DOMParser` directly). Package-level
 * behaviour lives in `tests/content-types.integration.test.ts`.
 */

import { describe, expect, it } from "vitest";

import { parseContentTypes, resolveContentType } from "../src/lib/content-types";
import { parseXml } from "../src/lib/xml-helpers";

const CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types";

const typesXml = (inner: string): string => `<?xml version="1.0"?><Types xmlns="${CT_NS}">${inner}</Types>`;

const DOC_CT = "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml";

describe("parseContentTypes + resolveContentType", () => {
    it("resolves a part covered by an exact Override with source 'override'", () => {
        const table = parseContentTypes(parseXml(typesXml(`<Override PartName="/word/document.xml" ContentType="${DOC_CT}"/>`)));
        expect(resolveContentType(table, "/word/document.xml")).toEqual({ contentType: DOC_CT, source: "override" });
        // Leading slash is optional on lookup — package-relative paths walk
        // in without one.
        expect(resolveContentType(table, "word/document.xml")).toEqual({ contentType: DOC_CT, source: "override" });
    });

    it("resolves a part covered only by a Default extension with source 'default'", () => {
        const table = parseContentTypes(parseXml(typesXml(`<Default Extension="png" ContentType="image/png"/>`)));
        expect(resolveContentType(table, "/word/media/image1.png")).toEqual({ contentType: "image/png", source: "default" });
    });

    it("prefers the Override over the Default for the same part", () => {
        // OPC §10.1.2.4: an exact Override wins over an extension Default.
        const table = parseContentTypes(
            parseXml(
                typesXml(
                    `<Default Extension="jpeg" ContentType="image/jpeg"/>` +
                        `<Override PartName="/word/media/dupe.jpeg" ContentType="image/tiff"/>`,
                ),
            ),
        );
        expect(resolveContentType(table, "/word/media/dupe.jpeg")).toEqual({ contentType: "image/tiff", source: "override" });
    });

    it("returns null when neither Override nor Default covers the part", () => {
        const table = parseContentTypes(parseXml(typesXml(`<Default Extension="xml" ContentType="application/xml"/>`)));
        expect(resolveContentType(table, "/word/media/image1.jpeg")).toBeNull();
    });
});
