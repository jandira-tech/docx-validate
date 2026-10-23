/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

/**
 * w:lvlPicBulletId whose value is not a w:numPicBullet id in the
 * same numbering part makes Word offer recovery (fixture 1541).
 * Removing that element opens clean. The p16 restart and durableId
 * attributes on the same part stayed, and Word still opened clean.
 */

import type { CheckDefinition } from "../../../lib/check-registry";
import type { ValidationContext } from "../../../lib/validation-context";
import type { ValidationIssue } from "../../../lib/types";
import { WORD_PARAGRAPH_NAMESPACES } from "../word-namespaces";
import path from "node:path";
import { getElementsByTagNameNSAll } from "../../../lib/xml-helpers";

export const wordNumberingPictureBulletCheck: CheckDefinition = {
    id: "word-numbering-picture-bullet",
    description: "Numbering levels reference picture bullets defined in the same part.",
    groups: ["story"],
    formats: ["docx"],
    requires: ["xml-syntax"],
    run: (ctx: ValidationContext): ValidationIssue[] => {
        const issues: ValidationIssue[] = [];

        for (const xmlFile of ctx.partNames) {
            if (path.posix.basename(xmlFile) !== "numbering.xml") {
                continue;
            }
            const parsed = ctx.effectiveXml(xmlFile);
            if (!parsed?.ok) {
                continue;
            }
            const dom = parsed.document;
            for (const ns of WORD_PARAGRAPH_NAMESPACES) {
                const defined = new Set<string>();
                const bullets = getElementsByTagNameNSAll(dom, ns, "numPicBullet");
                for (const entry of bullets) {
                    const id = entry?.getAttributeNS(ns, "numPicBulletId");
                    if (id) {
                        defined.add(id);
                    }
                }
                const refs = getElementsByTagNameNSAll(dom, ns, "lvlPicBulletId");
                for (const entry of refs) {
                    const val = entry?.getAttributeNS(ns, "val");
                    if (!val || defined.has(val)) {
                        continue;
                    }
                    issues.push({
                        severity: "error",
                        message:
                            `A numbering level references picture bullet ${val}, and this part has no w:numPicBullet with that id. ` +
                            "Microsoft Word offers to recover the file. " +
                            "Remove the w:lvlPicBulletId, or add the matching picture bullet.",
                        path: xmlFile,
                        code: "word-numbering-picture-bullet",
                    });
                }
            }
        }

        return issues;
    },
};
