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
 *
 * The reference itself must name a relationship of exactly the
 * Microsoft `attachedToolbars` Type (#716; session 48, Word 16.115):
 * a `wne:toolbarData r:id` naming a relationship of another Type —
 * `…/relationships/custom/attachedToolbars` or
 * `http://example.com/attachedToolbars` — gets the recovery offer
 * whatever the target's bytes hold, clean or refused, as does one
 * naming no relationship. The UTF-8 byte scan follows only the exact
 * Type; the Type check is `word-attached-toolbar-relationship`.
 */

import type { CheckDefinition } from "../../../lib/check-registry";
import { equivalentPartLookup } from "../../../lib/content-types";
import type { ValidationContext } from "../../../lib/validation-context";
import type { ValidationIssue } from "../../../lib/types";
import path from "node:path";
import { registerStoryReferenceIdentity } from "../story-reference-identity";
import { getElementsByTagNameNSAll, getRelationshipChildren } from "../../../lib/xml-helpers";

const OFFICE_RELATIONSHIPS_NAMESPACE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

/** The only relationship Type Word accepts for `wne:toolbarData r:id` (#716, session 48). */
const ATTACHED_TOOLBARS_RELATIONSHIP_TYPE = "http://schemas.microsoft.com/office/2006/relationships/attachedToolbars";

/** One `wne:toolbarData r:id`, resolved against its part's own sidecar. */
type ToolbarReference = {
    readonly xmlFile: string;
    readonly relsFile: string;
    readonly elem: Element;
    readonly rid: string;
    /** The relationships namespace the `r:id` was read in; `""` for a literal `r:id`. */
    readonly ridNamespaceURI: string;
    /** The Relationship element `rid` names, or `undefined` when the sidecar holds none. */
    readonly relationship: Element | undefined;
};

/**
 * Each `wne:toolbarData r:id` in the package's XML parts with the
 * relationship its part's own sidecar gives it. A part with no sidecar
 * resolves every `r:id` to `undefined`; an unparseable sidecar yields
 * nothing, left for `xml-syntax` to report.
 */
const toolbarReferences = function* (ctx: ValidationContext): Generator<ToolbarReference> {
    const wordmlNs = "http://schemas.microsoft.com/office/word/2006/wordml";
    const relationshipNamespaces = [OFFICE_RELATIONSHIPS_NAMESPACE, "http://purl.oclc.org/ooxml/officeDocument/relationships"];
    for (const xmlFile of ctx.partNames) {
        if (!xmlFile.endsWith(".xml")) {
            continue;
        }
        const parsed = ctx.effectiveXml(xmlFile);
        if (!parsed?.ok) {
            continue;
        }
        const toolbarData = getElementsByTagNameNSAll(parsed.document, wordmlNs, "toolbarData");
        if (toolbarData.length === 0) {
            continue;
        }
        const relsFile = path.posix.join(path.posix.dirname(xmlFile), "_rels", `${path.posix.basename(xmlFile)}.rels`);
        const parsedRels = ctx.effectiveXml(relsFile);
        if (parsedRels !== undefined && !parsedRels.ok) {
            continue;
        }
        const relationships = parsedRels?.ok ? getRelationshipChildren(parsedRels.document) : [];
        for (const elem of toolbarData) {
            let rid: string | null = null;
            let ridNamespaceURI = "";
            for (const relNs of relationshipNamespaces) {
                const value = elem.getAttributeNS(relNs, "id");
                if (value) {
                    rid = value;
                    ridNamespaceURI = relNs;
                    break;
                }
            }
            if (!rid) {
                rid = elem.getAttribute("r:id");
            }
            if (!rid) {
                continue;
            }
            yield {
                xmlFile,
                relsFile,
                elem,
                rid,
                ridNamespaceURI,
                relationship: relationships.find((rel) => rel.getAttribute("Id") === rid),
            };
        }
    }
};

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

export const WORD_ATTACHED_TOOLBAR_RELATIONSHIP = "word-attached-toolbar-relationship";

export const wordAttachedToolbarRelationshipCheck: CheckDefinition = {
    id: WORD_ATTACHED_TOOLBAR_RELATIONSHIP,
    description: "wne:toolbarData references name an attachedToolbars relationship.",
    groups: ["settings"],
    formats: ["docx"],
    requires: ["xml-syntax"],
    run: (ctx: ValidationContext): ValidationIssue[] => {
        const issues: ValidationIssue[] = [];
        for (const { xmlFile, elem, rid, ridNamespaceURI, relationship } of toolbarReferences(ctx)) {
            if (relationship !== undefined && (relationship.getAttribute("Type") ?? "") === ATTACHED_TOOLBARS_RELATIONSHIP_TYPE) {
                continue;
            }
            const why =
                relationship === undefined
                    ? `no Relationship with Id '${rid}'`
                    : `Relationship '${rid}' has Type '${relationship.getAttribute("Type") ?? "(none)"}'`;
            const issue: ValidationIssue = {
                severity: "error",
                message:
                    `<wne:toolbarData> r:id="${rid}" does not reference an attached toolbar relationship (${why}). ` +
                    "Microsoft Word offers to recover the file. " +
                    `Use a Relationship of Type ${ATTACHED_TOOLBARS_RELATIONSHIP_TYPE}, or point the r:id at one.`,
                path: xmlFile,
                code: WORD_ATTACHED_TOOLBAR_RELATIONSHIP,
            };
            // #716: a dangling Transitional r:id here is also the legacy
            // walk's rels-id-missing; the full run's hand-off pairs the two
            // through this identity, never the prose.
            registerStoryReferenceIdentity(issue, {
                part: xmlFile,
                namespaceURI: elem.namespaceURI ?? "",
                localName: "toolbarData",
                rid,
                ridNamespaceURI,
            });
            issues.push(issue);
        }
        return issues;
    },
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

        const needle = Buffer.from([0xef, 0xbf, 0xbd]);
        for (const { relsFile, relationship } of toolbarReferences(ctx)) {
            // The byte scan follows only the exact Microsoft
            // attachedToolbars Type (#716); every other Type — and an `r:id`
            // naming no relationship — is the relationship check's finding.
            if (relationship === undefined || (relationship.getAttribute("Type") ?? "") !== ATTACHED_TOOLBARS_RELATIONSHIP_TYPE) {
                continue;
            }
            // An External target is a URI outside the package: it
            // names no part whose bytes this check reads (#359).
            if (relationship.getAttribute("TargetMode") === "External") {
                continue;
            }
            const target = relationship.getAttribute("Target");
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

        return issues;
    },
};
