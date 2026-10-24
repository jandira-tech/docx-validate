/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */
import type { CheckDefinition } from "../../../lib/check-registry";
import { CONTENT_TYPES_NAMESPACE, findContentTypesStream, partNameViolation } from "../../../lib/content-types";
import { getElementsByTagNameNSAll } from "../../../lib/xml-helpers";
import { wordRefusedNameReasons } from "./opc-part-name-word-refused";

export const CT_OVERRIDE_PARTNAME_WORD_REFUSED = "ct-override-partname-word-refused";

/**
 * Diagnose even Overrides naming no stored part (#451). Word 16.115 evidence:
 * docs/plans/evidence/word-probes-2026-10-24.md, sessions 6–12.
 * Literal non-ASCII already has an OPC grammar finding, so never repeat it.
 * The stream is judged after Markup Compatibility processing (OPC [M1.19]),
 * as `ct-override-partname-invalid` reads it: an `Override` in an ignored
 * branch is not judged, and `ct-markup-compatibility` reports the markup (#459).
 */
export const ctOverridePartnameWordRefusedCheck: CheckDefinition = {
    id: CT_OVERRIDE_PARTNAME_WORD_REFUSED,
    description: "Override PartNames contain no escaped marks or high octets that Word refuses.",
    groups: ["content-types"],
    formats: ["docx", "pptx", "xlsx"],
    requires: ["xml-syntax"],
    run: (ctx) => {
        const stream = findContentTypesStream(ctx.partNames);
        if (stream === undefined) return [];
        const parsed = ctx.effectiveXml(stream);
        if (!parsed?.ok) return [];
        return getElementsByTagNameNSAll(parsed.document, CONTENT_TYPES_NAMESPACE, "Override").flatMap((elem) => {
            const partName = elem.getAttribute("PartName") ?? "";
            const reasons = wordRefusedNameReasons(partName, partNameViolation(partName) === undefined);
            return reasons.length === 0
                ? []
                : [
                      {
                          code: CT_OVERRIDE_PARTNAME_WORD_REFUSED,
                          severity: "error" as const,
                          path: stream,
                          message: `Override PartName ${JSON.stringify(partName)} contains ${reasons.join("; ")}; Word 16.115 offers recovery.`,
                      },
                  ];
        });
    },
};
