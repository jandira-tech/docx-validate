/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

/**
 * The `Override` `PartName` grammar (#414).
 */

import type { CheckDefinition } from "../../../lib/check-registry";
import { CONTENT_TYPES_NAMESPACE, findContentTypesStream, partNameViolation } from "../../../lib/content-types";
import type { ValidationIssue } from "../../../lib/types";
import type { ValidationContext } from "../../../lib/validation-context";
import { getElementsByTagNameNSAll } from "../../../lib/xml-helpers";

export const CT_OVERRIDE_PARTNAME_INVALID = "ct-override-partname-invalid";

/**
 * Report every `Override` whose `PartName` is not a valid OPC part name
 * (ECMA-376 Part 2, 4th edition (2012), §9.1.1.1). A part name shall be an
 * IRI and shall be encoded as either a Part URI (§9.1.1.1.2, RFC 3986
 * `pchar`) or a Part IRI (§9.1.1.1.1, RFC 3987 `ipchar`, which allows a
 * literal `ucschar`), so `/word/café.xml` and `/word/caf%C3%A9.xml` are both
 * valid while the mixed `/word/caf%C3%A9-é.xml` is neither — [M1.8] forbids
 * percent-encoded `iunreserved` characters in an IRI-only name. Strict
 * follows this 4th edition, the edition of the bundled
 * `schemas/ecma/fourth-edition/` OPC schemas. Such a PartName names no part,
 * so the declaration is malformed whatever its content type. Resolution is unchanged:
 * `parseContentTypes` still keys each PartName by its canonical spelling, so a
 * malformed one can still resolve a stored part (#457).
 * System.IO.Packaging is no reference here: `GetNormalizedPartUri` (.NET 10)
 * accepts an empty segment, `%2F`, a segment ending with `.`, `[` and `]`.
 *
 * Word 16.115 offers to recover a package that declares one (probe sessions
 * 3c, 4, 5, 7 and 12), even when nothing is stored under that name and the content
 * type is well formed, so the finding is Word-blocking for DOCX/DOCM; strict
 * reports it as an error, and lenient and the unobserved PPTX/XLSX warn. A
 * folder segment ending with `.` that holds another character is the
 * exception: Word opens it clean, so word-valid warns on it (#450). A dot-only
 * folder segment is refused, and reported under its own rule (#453). A
 * literal `ucschar` the grammar accepts is out of this check's scope: Word
 * still refuses the spelling, which `ct-override-partname-word-refused`
 * reports (#451, #455).
 *
 * The stream is judged after Markup Compatibility processing (OPC [M1.19],
 * as `word-content-type-invalid` already reads it), so an `Override` in an
 * ignored branch is not judged; Word refuses the markup itself, which
 * `ct-markup-compatibility` reports (#459).
 *
 * A content-types stream that does not parse is reported by the XML syntax
 * gate (#388); nothing is reported from it here.
 */
export const ctOverridePartnameInvalidCheck: CheckDefinition = {
    id: CT_OVERRIDE_PARTNAME_INVALID,
    description: "Override PartNames are valid OPC part names (ECMA-376 Part 2, 4th edition, §9.1.1.1: a Part URI or a Part IRI).",
    groups: ["content-types"],
    formats: ["docx", "pptx", "xlsx"],
    requires: ["xml-syntax"],
    run: (ctx: ValidationContext): ValidationIssue[] => {
        const issues: ValidationIssue[] = [];

        const stream = findContentTypesStream(ctx.partNames);
        if (stream === undefined) return issues;
        const parsed = ctx.effectiveXml(stream);
        if (!parsed?.ok) return issues;

        for (const elem of getElementsByTagNameNSAll(parsed.document, CONTENT_TYPES_NAMESPACE, "Override")) {
            // Word 16.115 offers recovery for a missing PartName as for an
            // empty one (session 13); only the wording tells them apart (#458).
            if (!elem.hasAttribute("PartName")) {
                issues.push({
                    severity: "error",
                    code: CT_OVERRIDE_PARTNAME_INVALID,
                    path: stream,
                    message: "Override has no PartName attribute, so it names no part.",
                });
                continue;
            }
            const partName = elem.getAttribute("PartName") ?? "";
            const violation = partNameViolation(partName);
            if (violation === undefined) continue;
            issues.push({
                severity: "error",
                code: CT_OVERRIDE_PARTNAME_INVALID,
                path: stream,
                message: `Override PartName ${JSON.stringify(partName)} is not a valid part name: it ${violation}.`,
            });
        }

        return issues;
    },
};
