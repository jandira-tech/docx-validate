/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

/**
 * The `word-part-content-type-generic` check (#432): a part Word reads, typed
 * only by the generic Default.
 *
 * Word 16.115 offers to recover the package for every part probed whose
 * content type resolves only through
 * `<Default Extension="xml" ContentType="application/xml"/>`: the main
 * document and styles with their Override removed, and settings, fontTable
 * and webSettings added with a relationship but no Override. Every control
 * with the part's proper Override opens clean (session 8a in
 * docs/plans/evidence/word-probes-2026-10-24.md). Numbering and customXml
 * were not probed, so they stay silent until a probe observes them.
 *
 * Resolution follows ECMA-376 Part 2 §10.1.2.4 through the shared
 * `src/lib/content-types.ts` table: an exact Override wins, so a part with a
 * specific Override — right or wrong — is out of scope here. The finding is
 * an error, and Word-blocking, under word-valid for DOCX/DOCM, an error
 * under strict (Part 1 gives each of these parts its own content type), and
 * a warning under lenient. A full strict run does not report the main
 * document, which `ct-undeclared-part` already rejects.
 */

import type { CheckDefinition } from "../../../lib/check-registry";
import { asciiLowercase, findContentTypesStream, parseContentTypes, resolveContentType } from "../../../lib/content-types";
import type { ValidationIssue } from "../../../lib/types";
import type { ValidationContext } from "../../../lib/validation-context";
import { wordTypedRelationshipTargets } from "../../../lib/word-read-parts";

export const WORD_PART_CONTENT_TYPE_GENERIC = "word-part-content-type-generic";

/** The only content type this check judges: the generic Default for `.xml`. */
const GENERIC_CONTENT_TYPE = "application/xml";

/** The relationship segments Word was observed to reject on the generic Default, and the content type each part needs. */
const EXPECTED_CONTENT_TYPES: ReadonlyMap<string, string> = new Map([
    ["officeDocument", "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"],
    ["styles", "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"],
    ["settings", "application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"],
    ["fontTable", "application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml"],
    ["webSettings", "application/vnd.openxmlformats-officedocument.wordprocessingml.webSettings+xml"],
]);

export const wordPartContentTypeGenericCheck: CheckDefinition = {
    id: WORD_PART_CONTENT_TYPE_GENERIC,
    description: "Parts Word reads carry their own content type, not the generic application/xml Default.",
    groups: ["content-types"],
    formats: ["docx"],
    requires: ["xml-syntax"],
    run: async (ctx: ValidationContext): Promise<ValidationIssue[]> => {
        const issues: ValidationIssue[] = [];

        // A missing stream is `ct-missing`; one that does not parse is the
        // XML syntax gate's finding. Neither is judged here.
        const stream = findContentTypesStream(ctx.partNames);
        if (stream === undefined) {
            return issues;
        }
        // Resolution reads the stream after Markup Compatibility processing
        // (OPC [M1.19]), as every content-type judgment does (#459, #489).
        const parsed = ctx.effectiveXml(stream);
        if (!parsed?.ok) {
            return issues;
        }
        const table = parseContentTypes(parsed.document);

        const targets = await wordTypedRelationshipTargets(ctx, new Set(EXPECTED_CONTENT_TYPES.keys()));
        for (const { part, segment } of targets) {
            const resolved = resolveContentType(table, part);
            if (resolved?.source !== "default" || asciiLowercase(resolved.contentType) !== GENERIC_CONTENT_TYPE) {
                continue;
            }
            const by =
                segment === "officeDocument" ? "the package officeDocument relationship" : `the main document's ${segment} relationship`;
            issues.push({
                severity: "error",
                code: WORD_PART_CONTENT_TYPE_GENERIC,
                path: part,
                message:
                    `${part}, named by ${by}, resolves only through the Default for extension "xml" to "${GENERIC_CONTENT_TYPE}"; ` +
                    `it needs its own Override "${EXPECTED_CONTENT_TYPES.get(segment)!}". Word 16.115 offers to recover the package.`,
            });
        }

        return issues;
    },
};
