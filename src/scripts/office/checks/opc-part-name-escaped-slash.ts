/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

import type { CheckDefinition } from "../../../lib/check-registry";
import { hasEscapedSeparatorCharacter } from "../../../lib/content-types";

export const OPC_PART_NAME_ESCAPED_SLASH = "opc-part-name-escaped-slash";

/**
 * Diagnose stored spellings before OPC equivalence or extraction can drop
 * them (#436). Word 16.115 offers recovery for an unreferenced
 * `word%2Fxtra.xml`, refuses `word%5Cxtra.xml`, and opens the well-formed
 * `word/xtra.xml` control clean (docs/plans/evidence/word-probes-2026-10-24.md).
 */
export const opcPartNameEscapedSlashCheck: CheckDefinition = {
    id: OPC_PART_NAME_ESCAPED_SLASH,
    description: "Stored part-name segments contain no percent-encoded '/' or '\\' separators (OPC [M1.7]).",
    groups: ["package"],
    formats: ["docx", "pptx", "xlsx"],
    storedNamesOnly: true,
    run: (ctx) =>
        // Each distinct stored spelling once: two entries may share a name (#411).
        [...new Set(ctx.storedNames)].filter(hasEscapedSeparatorCharacter).map((name) => ({
            code: OPC_PART_NAME_ESCAPED_SLASH,
            severity: "error",
            path: name,
            message: `OPC [M1.7] forbids percent-encoded '/' or '\\' in stored name: ${JSON.stringify(name)}.`,
        })),
};
