/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

/**
 * Two theme relationships in one .rels part make Word offer recovery.
 * Either theme alone opens clean; rewriting an absolute target still warns.
 * Fixture: tests/fixtures/broken/package.word-duplicate-theme.docx.
 *
 * Two `fontTable` relationships in one `.rels` part make Word offer to
 * recover the file. A second relationship that targets the same
 * `fontTable.xml` as the first still warns. Rewriting an absolute
 * `/word/...` target to a relative path still warns. One font table
 * relationship opens clean. The extra part and its content type are
 * not required for the warning.
 *
 * Two `webSettings` relationships in one `.rels` part make Word offer to
 * recover the file. A second relationship that targets the same
 * `webSettings.xml` as the first still warns. One web settings
 * relationship opens clean. The extra part is not required for the warning.
 *
 * Two settings relationships in one .rels part make Word offer
 * recovery (fixture 0689). A second relationship that targets
 * the same settings.xml still warns. One settings relationship
 * opens clean.
 */

import type { CheckDefinition } from "../../../lib/check-registry";
import { isRelationshipsPartName } from "../../../lib/content-types";
import type { ValidationContext } from "../../../lib/validation-context";
import type { ValidationIssue } from "../../../lib/types";
import { getElementsByTagNameNSAll } from "../../../lib/xml-helpers";

const PACKAGE_RELATIONSHIPS_NAMESPACE = "http://schemas.openxmlformats.org/package/2006/relationships";

export const wordDuplicateThemeRelationshipCheck: CheckDefinition = {
    id: "word-duplicate-theme-relationship",
    description: "Relationship parts contain at most one theme relationship.",
    groups: ["relationships"],
    formats: ["docx"],
    requires: ["xml-syntax"],
    run: (ctx: ValidationContext): ValidationIssue[] => {
        const issues: ValidationIssue[] = [];

        for (const relsFile of ctx.partNames) {
            if (!isRelationshipsPartName(relsFile)) {
                continue;
            }
            const parsed = ctx.effectiveXml(relsFile);
            if (!parsed?.ok) {
                continue;
            }
            const dom = parsed.document;
            const rels = getElementsByTagNameNSAll(dom, PACKAGE_RELATIONSHIPS_NAMESPACE, "Relationship");
            const themeIds: string[] = [];
            for (const rel of rels) {
                const type = rel.getAttribute("Type") ?? "";
                if (!type.endsWith("/theme")) {
                    continue;
                }
                themeIds.push(rel.getAttribute("Id") ?? "");
            }
            if (themeIds.length < 2) {
                continue;
            }
            issues.push({
                severity: "error",
                message:
                    `${relsFile} contains ${themeIds.length} theme relationships ` +
                    `(${themeIds.join(", ")}). Microsoft Word offers to recover the file. Keep one theme relationship.`,
                path: relsFile,
                code: "word-duplicate-theme-relationship",
            });
        }

        return issues;
    },
};

export const wordDuplicateFontTableRelationshipCheck: CheckDefinition = {
    id: "word-duplicate-font-table-relationship",
    description: "Relationship parts contain at most one font table relationship.",
    groups: ["relationships"],
    formats: ["docx"],
    requires: ["xml-syntax"],
    run: (ctx: ValidationContext): ValidationIssue[] => {
        const issues: ValidationIssue[] = [];

        for (const relsFile of ctx.partNames) {
            if (!isRelationshipsPartName(relsFile)) {
                continue;
            }
            const parsed = ctx.effectiveXml(relsFile);
            if (!parsed?.ok) {
                continue;
            }
            const dom = parsed.document;
            const rels = getElementsByTagNameNSAll(dom, PACKAGE_RELATIONSHIPS_NAMESPACE, "Relationship");
            const fontIds: string[] = [];
            for (const rel of rels) {
                const type = rel.getAttribute("Type") ?? "";
                if (!type.endsWith("/fontTable")) {
                    continue;
                }
                fontIds.push(rel.getAttribute("Id") ?? "");
            }
            if (fontIds.length < 2) {
                continue;
            }
            issues.push({
                severity: "error",
                message:
                    `${relsFile} contains ${fontIds.length} font table relationships ` +
                    `(${fontIds.join(", ")}). Microsoft Word offers to recover the file. Keep one font table relationship.`,
                path: relsFile,
                code: "word-duplicate-font-table-relationship",
            });
        }

        return issues;
    },
};

export const wordDuplicateWebSettingsRelationshipCheck: CheckDefinition = {
    id: "word-duplicate-web-settings-relationship",
    description: "Relationship parts contain at most one web settings relationship.",
    groups: ["relationships"],
    formats: ["docx"],
    requires: ["xml-syntax"],
    run: (ctx: ValidationContext): ValidationIssue[] => {
        const issues: ValidationIssue[] = [];

        for (const relsFile of ctx.partNames) {
            if (!isRelationshipsPartName(relsFile)) {
                continue;
            }
            const parsed = ctx.effectiveXml(relsFile);
            if (!parsed?.ok) {
                continue;
            }
            const dom = parsed.document;
            const rels = getElementsByTagNameNSAll(dom, PACKAGE_RELATIONSHIPS_NAMESPACE, "Relationship");
            const webIds: string[] = [];
            for (const rel of rels) {
                const type = rel.getAttribute("Type") ?? "";
                if (!type.endsWith("/webSettings")) {
                    continue;
                }
                webIds.push(rel.getAttribute("Id") ?? "");
            }
            if (webIds.length < 2) {
                continue;
            }
            issues.push({
                severity: "error",
                message:
                    `${relsFile} contains ${webIds.length} web settings relationships ` +
                    `(${webIds.join(", ")}). Microsoft Word offers to recover the file. Keep one web settings relationship.`,
                path: relsFile,
                code: "word-duplicate-web-settings-relationship",
            });
        }

        return issues;
    },
};

export const wordDuplicateSettingsRelationshipCheck: CheckDefinition = {
    id: "word-duplicate-settings-relationship",
    description: "Relationship parts contain at most one settings relationship.",
    groups: ["relationships"],
    formats: ["docx"],
    requires: ["xml-syntax"],
    run: (ctx: ValidationContext): ValidationIssue[] => {
        const issues: ValidationIssue[] = [];

        for (const relsFile of ctx.partNames) {
            if (!isRelationshipsPartName(relsFile)) {
                continue;
            }
            const parsed = ctx.effectiveXml(relsFile);
            if (!parsed?.ok) {
                continue;
            }
            const dom = parsed.document;
            const rels = getElementsByTagNameNSAll(dom, PACKAGE_RELATIONSHIPS_NAMESPACE, "Relationship");
            const settingsIds: string[] = [];
            for (const rel of rels) {
                const type = rel.getAttribute("Type") ?? "";
                if (!type.endsWith("/settings")) {
                    continue;
                }
                settingsIds.push(rel.getAttribute("Id") ?? "");
            }
            if (settingsIds.length < 2) {
                continue;
            }
            issues.push({
                severity: "error",
                message:
                    `${relsFile} contains ${settingsIds.length} settings relationships ` +
                    `(${settingsIds.join(", ")}). Microsoft Word offers to recover the file. Keep one settings relationship.`,
                path: relsFile,
                code: "word-duplicate-settings-relationship",
            });
        }

        return issues;
    },
};
