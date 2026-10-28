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
 *
 * The walk judges XML-named parts (`isXmlPartName`), excluding
 * relationships parts and XML-typed non-XML names (#759). It also
 * judges case variants: a `word/customizations.XML` and its
 * `customizations.XML.rels` get the recovery offer (session 54, s54c;
 * #743), and the part the main document's `keyMapCustomizations`
 * relationship reaches, whatever its name or extension (#777; session
 * 57, s57d): Word offered recovery for the refused toolbar fixture
 * with `word/customizations.xml` renamed `word/customizations` (no
 * extension) and `word/customizations.bin` — its Override keeping the
 * keyMapCustomizations type — and opened the same renames of the
 * repaired fixture clean, so the rename itself is not what Word
 * refuses. A glossary's customizations relationship was not shown and
 * is not followed. The tool reads each sidecar by OPC part-name equivalence,
 * resolving escaped and backslash spellings, and compares the sidecar's
 * own `_rels` folder segment by the same equivalence (#772); Word's
 * observation covers only `customizations.xml.RELS`, which opened clean
 * (session 54, s54c; #398, #402, #743). Session 3 recorded Word refusing
 * escaped part names (#408, #436). An empty `r:id` is not judged (s54c).
 */

import type { CheckDefinition } from "../../../lib/check-registry";
import {
    equivalentPartLookup,
    isRelationshipsFolderName,
    isRelationshipsPartName,
    relationshipTargetFolder,
    relationshipTargetPartName,
} from "../../../lib/content-types";
import type { ValidationContext } from "../../../lib/validation-context";
import type { Profile, ValidationIssue } from "../../../lib/types";
import path from "node:path";
import { otherRelationshipIds, registerStoryReferenceIdentity } from "../story-reference-identity";
import { wordExactTypeTargets } from "../../../lib/word-read-parts";
import { getElementsByTagNameNSAll, getRelationshipChildren } from "../../../lib/xml-helpers";
import { isXmlPartName } from "./xml";

const OFFICE_RELATIONSHIPS_NAMESPACE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

/** The only relationship Type Word accepts for `wne:toolbarData r:id` (#716, session 48). */
const ATTACHED_TOOLBARS_RELATIONSHIP_TYPE = "http://schemas.microsoft.com/office/2006/relationships/attachedToolbars";

/** The relationship Type the main document reaches its customizations part through (#777, session 57). */
const KEYMAP_CUSTOMIZATIONS_RELATIONSHIP_TYPE = "http://schemas.microsoft.com/office/2006/relationships/keyMapCustomizations";

/** One `wne:toolbarData r:id`, resolved against its part's own sidecar. */
type ToolbarReference = {
    readonly xmlFile: string;
    /** The stored spelling of the part's sidecar, or `undefined` when the part has none. */
    readonly relsFile: string | undefined;
    readonly elem: Element;
    readonly rid: string;
    /** The relationships namespace the `r:id` was read in; `""` for a literal `r:id`. */
    readonly ridNamespaceURI: string;
    /** The element's other non-empty ids in the relationships namespaces (#1079). */
    readonly otherRids: readonly string[];
    /** The Relationship element `rid` names, or `undefined` when the sidecar holds none. */
    readonly relationship: Element | undefined;
};

/**
 * Each `wne:toolbarData r:id` in the package's XML-named parts — excluding
 * relationships parts, so a `.XML` spelling is judged (session
 * 54, s54c; #743) — and in the part the main document's
 * `keyMapCustomizations` relationship reaches, whatever its name or
 * extension (#777, session 57), a part both walks reach judged once —
 * with the relationship its part's own sidecar gives it, the sidecar
 * found through the shared OPC part-name index,
 * so the tool reads a `.RELS` spelling (s54c). A part with no sidecar
 * resolves every `r:id` to `undefined`; an unparseable sidecar yields
 * nothing, left for `xml-syntax` to report.
 */
const toolbarReferences = async function* (
    ctx: ValidationContext,
    lookup: (partName: string) => string | undefined,
    profile?: Profile,
): AsyncGenerator<ToolbarReference> {
    const wordmlNs = "http://schemas.microsoft.com/office/word/2006/wordml";
    const relationshipNamespaces = [OFFICE_RELATIONSHIPS_NAMESPACE, "http://purl.oclc.org/ooxml/officeDocument/relationships"];
    const xmlParts = ctx.partNames.filter((part) => isXmlPartName(part) && !isRelationshipsPartName(part));
    // Word reaches the customizations part through the main document's
    // keyMapCustomizations relationship, not by its name (#777, session
    // 57), so that part is judged whatever its name or extension; a part
    // both walks reach stays in the set once.
    const judgedParts = new Set(xmlParts);
    for (const reached of await wordExactTypeTargets(ctx, KEYMAP_CUSTOMIZATIONS_RELATIONSHIP_TYPE, profile)) {
        judgedParts.add(reached);
    }
    // The shared index and this list keep a bounded number of passes over
    // the part names, including when a sidecar is absent (#358, #758).
    for (const xmlFile of judgedParts) {
        const parsed = ctx.effectiveXml(xmlFile);
        if (!parsed?.ok) {
            continue;
        }
        const toolbarData = getElementsByTagNameNSAll(parsed.document, wordmlNs, "toolbarData");
        if (toolbarData.length === 0) {
            continue;
        }
        const relsFile = lookup(path.posix.join(path.posix.dirname(xmlFile), "_rels", `${path.posix.basename(xmlFile)}.rels`));
        const parsedRels = relsFile === undefined ? undefined : ctx.effectiveXml(relsFile);
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
                // #1079: the legacy walk's copy for either id pairs.
                otherRids: otherRelationshipIds(elem, relationshipNamespaces, rid),
                relationship: relationships.find((rel) => rel.getAttribute("Id") === rid),
            };
        }
    }
};

const resolveToolbarTarget = (relsFile: string, target: string): string | null => {
    // The target reads as the read model and both file-reference walks read
    // it (#991): `\` reads as `/`, as it does everywhere in the package view
    // (#350); a query, like a fragment, is not part of the path (#359); a
    // climb past the package root stops at the root (#984, #986); and a
    // target naming a folder names no part (#962, #975). The sidecar's own
    // `_rels` folder segment compares canonically, as the sidecar lookup
    // does (#772).
    const relsDir = path.posix.dirname(relsFile);
    const baseDir = isRelationshipsFolderName(path.posix.basename(relsDir)) ? path.posix.dirname(relsDir) : relsDir;
    const resolved = relationshipTargetPartName(baseDir, target);
    if (resolved === undefined || relationshipTargetFolder(baseDir, target) !== undefined) {
        return null;
    }
    return resolved.replace(/^\//, "");
};

export const WORD_ATTACHED_TOOLBAR_RELATIONSHIP = "word-attached-toolbar-relationship";

export const wordAttachedToolbarRelationshipCheck: CheckDefinition = {
    id: WORD_ATTACHED_TOOLBAR_RELATIONSHIP,
    description: "wne:toolbarData references name an attachedToolbars relationship.",
    groups: ["settings"],
    formats: ["docx"],
    requires: ["xml-syntax"],
    run: async (ctx: ValidationContext, profile?: Profile): Promise<ValidationIssue[]> => {
        const issues: ValidationIssue[] = [];
        const lookup = equivalentPartLookup(ctx.partNames);
        for await (const { xmlFile, elem, rid, ridNamespaceURI, otherRids, relationship } of toolbarReferences(ctx, lookup, profile)) {
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
                otherRids,
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
    run: async (ctx: ValidationContext, profile?: Profile): Promise<ValidationIssue[]> => {
        const issues: ValidationIssue[] = [];
        // One finding per referenced part, however many toolbarData elements
        // (in however many parts) name it (Jules on #323). A part is marked
        // seen before its bytes are read, so a clean part named again is not
        // scanned again (Jules on #345).
        const seen = new Set<string>();
        const lookup = equivalentPartLookup(ctx.partNames);

        const needle = Buffer.from([0xef, 0xbf, 0xbd]);
        for await (const { relsFile, relationship } of toolbarReferences(ctx, lookup, profile)) {
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
            const resolved = relsFile === undefined ? null : resolveToolbarTarget(relsFile, target);
            // The entry the target names is both the key and the reported
            // path, so two spellings of one part yield one finding under
            // its real name (#350).
            // One index per run serves sidecars and targets (#358, #758).
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
