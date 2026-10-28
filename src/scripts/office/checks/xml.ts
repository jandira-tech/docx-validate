/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

/**
 * XML syntax gate. Every DOM-reading check requires it, so a malformed part
 * stops those checks before they run instead of letting them crash or report
 * findings about a document the parser had to guess at.
 */

import type { CheckDefinition } from "../../../lib/check-registry";
import {
    canonicalPartName,
    contentTypeMediaType,
    type ContentTypeTable,
    findContentTypesStream,
    parseContentTypes,
    resolveContentType,
} from "../../../lib/content-types";
import { parseXmlBytes } from "../../../lib/xml-helpers";
import { checkXmlSyntax } from "../../../lib/xml-syntax";
import type { ValidationIssue } from "../../../lib/types";

const XML_PART = /\.(?:xml|rels)$/i;

/**
 * Whether a part is one the XML syntax gate parses: its extension is `.xml`
 * or `.rels` as OPC compares part names, ASCII case-insensitively and with
 * escapes of unreserved characters decoded, so `word/COMMENTS.XML` and
 * `word/comments.%78ml` are XML parts. The full run's own XML pass uses this
 * predicate too, so both judge the same parts (#390).
 */
export const isXmlPartName = (partName: string): boolean => XML_PART.test(canonicalPartName(partName));

/**
 * Whether a content type declares XML content: its media type is
 * `application/xml` or `text/xml` (RFC 7303), or an `application/` type with
 * the `+xml` suffix (RFC 6839), as every OOXML XML part type is. Image types
 * such as `image/svg+xml` are left out: Word renders them through its image
 * pipeline, and no probe has observed it refuse malformed markup there.
 */
export const isXmlContentType = (contentType: string): boolean => {
    const media = contentTypeMediaType(contentType);
    return media === "application/xml" || media === "text/xml" || (media.startsWith("application/") && media.endsWith("+xml"));
};

/** The content-types table of a package from its stream's bytes, or `undefined` when the stream is absent or does not parse. */
export const contentTypeTableOf = (bytes: Uint8Array | undefined): ContentTypeTable | undefined => {
    if (bytes === undefined) return undefined;
    try {
        return parseContentTypes(parseXmlBytes(bytes));
    } catch {
        return undefined;
    }
};

/**
 * The parts the XML passes judge: every part {@link isXmlPartName} accepts,
 * and every other part the content-types stream declares with an XML type
 * (#620). Word reads a part by its relationship and its content type, not
 * its name: a main document stored as `word/document.bin` under the main
 * document's Override opens as one, and Word refuses it truncated (session
 * 32). `bytesOf` reads the content-types stream; with none that parses, only
 * names decide.
 */
export const xmlPartNames = (partNames: readonly string[], bytesOf: (part: string) => Uint8Array | undefined): Set<string> => {
    const stream = findContentTypesStream(partNames);
    const table = stream === undefined ? undefined : contentTypeTableOf(bytesOf(stream));
    return new Set(
        partNames.filter((part) => {
            if (isXmlPartName(part)) return true;
            const resolved = table === undefined ? null : resolveContentType(table, part);
            return resolved !== null && isXmlContentType(resolved.contentType);
        }),
    );
};

/**
 * The id of the XML well-formedness check. The validators' own XML pass
 * does the same work on the same bytes, so they report it under this id
 * when they disclose the checks a run executed (#319).
 */
export const XML_SYNTAX_CHECK_ID = "xml-syntax";

export const xmlSyntaxCheck: CheckDefinition = {
    id: XML_SYNTAX_CHECK_ID,
    description: "Every XML part of the package is well-formed, parsed from its original bytes without recovery.",
    groups: ["xml"],
    formats: ["docx", "pptx", "xlsx"],
    // A malformed part hides only itself from the checks that require this
    // one; they still judge every part that parsed (#307).
    gatesByPart: true,
    run: async (ctx): Promise<ValidationIssue[]> => {
        const issues: ValidationIssue[] = [];
        const xmlParts = xmlPartNames(ctx.partNames, (part) => ctx.bytes(part));
        for (const part of ctx.partNames) {
            const bytes = xmlParts.has(part) ? ctx.bytes(part) : undefined;
            if (bytes !== undefined) {
                issues.push(...(await checkXmlSyntax(bytes, part)));
            }
        }
        return issues;
    },
};
