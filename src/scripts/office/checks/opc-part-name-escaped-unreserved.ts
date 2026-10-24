/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

import type { CheckDefinition } from "../../../lib/check-registry";
import { hasEscapedUnreservedCharacter } from "../../../lib/content-types";

export const OPC_PART_NAME_ESCAPED_UNRESERVED = "opc-part-name-escaped-unreserved";

/** Diagnose stored spellings before OPC equivalence or extraction can drop them (#408). */
export const opcPartNameEscapedUnreservedCheck: CheckDefinition = {
    id: OPC_PART_NAME_ESCAPED_UNRESERVED,
    description: "Stored part-name segments contain no percent-encoded unreserved characters (OPC [M1.8]).",
    groups: ["package"],
    formats: ["docx", "pptx", "xlsx"],
    run: (ctx) =>
        // Each distinct stored spelling once: two entries may share a name (#411).
        [...new Set(ctx.storedNames)].filter(hasEscapedUnreservedCharacter).map((name) => ({
            code: OPC_PART_NAME_ESCAPED_UNRESERVED,
            severity: "error",
            path: name,
            message: `OPC [M1.8] forbids percent-encoded unreserved characters in stored name: ${JSON.stringify(name)}.`,
        })),
};
