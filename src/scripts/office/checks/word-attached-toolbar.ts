/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

/**
 * word/attachedToolbars.bin referenced by wne:toolbarData contains
 * the UTF-8 replacement sequence EF BF BD (fixture 2130). One
 * sequence is enough. The same bytes in a part toolbarData does
 * not reference are tolerated. A toolbar whose sequences have
 * been replaced, including button-flag bytes set to 0x10, opens
 * clean.
 */

import type { CheckDefinition } from "../../../lib/check-registry";
import { equivalentPartLookup } from "../../../lib/content-types";
import type { ValidationContext } from "../../../lib/validation-context";
import type { ValidationIssue } from "../../../lib/types";
import path from "node:path";
import { getElementsByTagNameNSAll, getRelationshipChildren } from "../../../lib/xml-helpers";

const OFFICE_RELATIONSHIPS_NAMESPACE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

const resolveToolbarTarget = (relsFile: string, target: string): string | null => {
    // `\` reads as `/`, as it does everywhere in the package view (#350). A
    // query, like a fragment, is not part of the path (#359).
    const [name] = target.replace(/\\/g, "/").split(/[?#]/, 1);
    if (!name) {
        return null;
    }
    const relsDir = path.posix.dirname(relsFile);
    const baseDir = path.posix.basename(relsDir) === "_rels" ? path.posix.dirname(relsDir) : relsDir;
    const resolved = path.posix.normalize(name.startsWith("/") ? name.replace(/^\/+/, "") : path.posix.join(baseDir, name));
    return resolved === ".." || resolved.startsWith("../") ? null : resolved;
};

export const wordAttachedToolbarUtf8Check: CheckDefinition = {
    id: "word-attached-toolbar-utf8",
    description: "Referenced attached toolbars contain no UTF-8 replacement byte sequences.",
    groups: ["settings"],
    formats: ["docx"],
    requires: ["xml-syntax"],
    run: (ctx: ValidationContext): ValidationIssue[] => {
        const issues: ValidationIssue[] = [];
        // One finding per referenced part, however many toolbarData elements
        // (in however many parts) name it (Jules on #323). A part is marked
        // seen before its bytes are read, so a clean part named again is not
        // scanned again (Jules on #345).
        const seen = new Set<string>();
        let lookup: ((partName: string) => string | undefined) | undefined;

        const wordmlNs = "http://schemas.microsoft.com/office/word/2006/wordml";
        const relationshipNamespaces = [OFFICE_RELATIONSHIPS_NAMESPACE, "http://purl.oclc.org/ooxml/officeDocument/relationships"];
        const needle = Buffer.from([0xef, 0xbf, 0xbd]);
        for (const xmlFile of ctx.partNames) {
            if (!xmlFile.endsWith(".xml")) {
                continue;
            }
            const parsed = ctx.effectiveXml(xmlFile);
            if (!parsed?.ok) {
                continue;
            }
            const dom = parsed.document;
            const toolbarData = getElementsByTagNameNSAll(dom, wordmlNs, "toolbarData");
            if (toolbarData.length === 0) {
                continue;
            }
            const relsFile = path.posix.join(path.posix.dirname(xmlFile), "_rels", `${path.posix.basename(xmlFile)}.rels`);
            const parsedRels = ctx.effectiveXml(relsFile);
            if (!parsedRels?.ok) {
                continue;
            }
            const relDom = parsedRels.document;
            const relationships = getRelationshipChildren(relDom);
            for (const elem of toolbarData) {
                let rid: string | null = null;
                for (const relNs of relationshipNamespaces) {
                    const value = elem.getAttributeNS(relNs, "id");
                    if (value) {
                        rid = value;
                        break;
                    }
                }
                if (!rid) {
                    rid = elem.getAttribute("r:id");
                }
                if (!rid) {
                    continue;
                }
                let target: string | null = null;
                for (const rel of relationships) {
                    if (!rel || rel.getAttribute("Id") !== rid) {
                        continue;
                    }
                    const type = rel.getAttribute("Type") ?? "";
                    if (!type.endsWith("/attachedToolbars")) {
                        continue;
                    }
                    // An External target is a URI outside the package: it
                    // names no part whose bytes this check reads (#359).
                    if (rel.getAttribute("TargetMode") !== "External") {
                        target = rel.getAttribute("Target");
                    }
                    break;
                }
                if (!target) {
                    continue;
                }
                const resolved = resolveToolbarTarget(relsFile, target);
                // The entry the target names is both the key and the reported
                // path, so two spellings of one part yield one finding under
                // its real name (#350).
                // One index per run, not a scan of every entry per reference (#358).
                lookup ??= equivalentPartLookup(ctx.partNames);
                const part = resolved && (lookup(resolved) ?? resolved);
                if (!part || seen.has(part)) {
                    continue;
                }
                seen.add(part);
                const bytes = ctx.bytes(part);
                if (!bytes) {
                    continue;
                }
                if (!Buffer.from(bytes).includes(needle)) {
                    continue;
                }
                issues.push({
                    severity: "error",
                    message:
                        `${part} is referenced by wne:toolbarData and contains the UTF-8 replacement sequence EF BF BD. ` +
                        "Microsoft Word offers to recover the file. One sequence is enough. " +
                        "Replace each sequence with the original byte.",
                    path: part,
                    code: "word-attached-toolbar-utf8",
                });
            }
        }

        return issues;
    },
};
