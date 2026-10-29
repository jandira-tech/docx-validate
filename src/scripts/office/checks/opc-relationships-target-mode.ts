/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

import type { CheckDefinition } from "../../../lib/check-registry";
import { canonicalPartName } from "../../../lib/content-types";
import type { Profile, ValidationIssue } from "../../../lib/types";
import type { ValidationContext } from "../../../lib/validation-context";
import { sidecarOf, wordReadTypedTargets } from "../../../lib/word-read-parts";
import { getRelationshipChildren } from "../../../lib/xml-helpers";

/** The relationships parts whose present TargetMode this check judges. */
export const targetModeRelationshipParts = async (ctx: ValidationContext, profile?: Profile): Promise<ReadonlySet<string>> => {
    const parts = new Set(ctx.partNames.filter((part) => canonicalPartName(part) === "_rels/.rels"));
    for (const target of await wordReadTypedTargets(ctx, profile)) {
        if (target.segment !== "officeDocument") continue;
        const sidecar = sidecarOf(ctx.partNames, target.part);
        if (sidecar !== undefined) parts.add(sidecar);
    }
    return parts;
};

/**
 * OPC's ST_TargetMode is exactly Internal | External. Word 16.115 offers
 * recovery for another spelling in the package or main document's
 * relationships part, whatever references the relationship (#753; session
 * 55, s55b–s55e): external on a chart or diagramData relationship, external,
 * External with a trailing space or empty on an unreferenced hyperlink in
 * the main sidecar, and external on that hyperlink in the package part.
 * Internal on the chart and diagramData and External on the hyperlinks
 * open clean. The same hyperlink with external in a referenced header's
 * sidecar opens clean (s55e), so that part keeps its verdict. Session 57
 * (s57f, #775) refused the same unreferenced hyperlink with external in a
 * reached glossary's sidecar; its External calibration opens clean. Main
 * documents and reached glossaries are found by relationship, wherever
 * stored. Other relationships parts — footnotes, comments and so on —
 * are unshown and keep their verdict.
 */
export const opcRelationshipsTargetModeCheck: CheckDefinition = {
    id: "opc-relationships-target-mode",
    description: "Package, main-document and reached-glossary relationships use exactly Internal or External when TargetMode is present.",
    groups: ["relationships"],
    formats: ["docx"],
    requires: ["xml-syntax"],
    run: async (ctx, profile): Promise<ValidationIssue[]> => {
        const parts = await targetModeRelationshipParts(ctx, profile);
        const issues: ValidationIssue[] = [];
        for (const part of parts) {
            const parsed = ctx.effectiveXml(part);
            if (!parsed?.ok) continue;
            for (const relationship of getRelationshipChildren(parsed.document)) {
                if (!relationship.hasAttribute("TargetMode")) continue;
                const mode = relationship.getAttribute("TargetMode");
                if (mode === "Internal" || mode === "External") continue;
                issues.push({
                    code: opcRelationshipsTargetModeCheck.id,
                    path: part,
                    severity: "error",
                    message: `Relationship ${JSON.stringify(relationship.getAttribute("Id") ?? "")} has TargetMode=${JSON.stringify(mode)}; expected exactly Internal or External. Microsoft Word offers to recover the file.`,
                });
            }
        }
        return issues;
    },
};
