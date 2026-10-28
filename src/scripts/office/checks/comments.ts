/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

/** Legacy comment integrity rules shared by full and selected DOCX runs. */
import path from "node:path";
import type { CheckDefinition } from "../../../lib/check-registry";
import { findEquivalentPart } from "../../../lib/content-types";
import type { Profile, ValidationIssue } from "../../../lib/types";
import type { ValidationContext } from "../../../lib/validation-context";
import { registerCommentReferenceMissing } from "../comment-finding-facts";
import {
    COMMENTS_EXTENDED_RELATIONSHIP_TYPE,
    COMMENTS_EXTENSIBLE_RELATIONSHIP_TYPE,
    COMMENTS_IDS_RELATIONSHIP_TYPE,
    wordExactTypeTargets,
    wordTypedRelationshipTargets,
} from "../../../lib/word-read-parts";
import { getElementsByTagNameNSAll } from "../../../lib/xml-helpers";
import { xmlPartNames } from "./xml";
import { WORD_2006_NAMESPACE, WORD_PARAGRAPH_NAMESPACES } from "../word-namespaces";

const W14_NAMESPACE = "http://schemas.microsoft.com/office/word/2010/wordml";
const W15_NAMESPACE = "http://schemas.microsoft.com/office/word/2012/wordml";
const W16CEX_NAMESPACE = "http://schemas.microsoft.com/office/word/2018/wordml/cex";
const W16CID_NAMESPACE = "http://schemas.microsoft.com/office/word/2016/wordml/cid";

/** Keep the legacy parse diagnostic rather than replacing it with an XML prerequisite's message. */
const commentXml = (ctx: ValidationContext, part: string): Document => {
    const parsed = ctx.xml(part)!;
    if (!parsed.ok) throw new Error(parsed.message);
    return parsed.document;
};

/**
 * The main document the comment checks judge (#696): the first part the
 * package's `officeDocument` relationship names, wherever it is stored, so
 * a reached glossary's own `document.xml` never stands in for it. When no
 * relationship names a stored part, or the package view cannot be read,
 * `word/document.xml` or its OPC-equivalent spelling (#711), never a
 * `document.xml` in another folder.
 */
const mainDocumentPart = async (ctx: ValidationContext, profile?: Profile): Promise<string | null> => {
    try {
        const [first] = await wordTypedRelationshipTargets(ctx, new Set(["officeDocument"]), profile);
        if (first !== undefined) {
            return first.part;
        }
    } catch {
        // An unreadable package view is reported by its own checks.
    }
    return inFolderByName(ctx, "word", "document.xml");
};

/**
 * The comments part the comment checks judge (#696): the first `comments`
 * target the main document's own relationships name, so a reached
 * glossary's comments never stand in for the main document's; a target that
 * is not XML leaves the comments unjudged, as #693 leaves the repair's. When
 * no relationship names a stored part, or the package view cannot be read,
 * the last `comments.xml` in the main document's folder, so the fallback
 * never pairs the main document with another story's comments.
 */
const mainCommentsPart = async (
    ctx: ValidationContext,
    document: string | null,
    xmlParts: ReadonlySet<string>,
    profile?: Profile,
): Promise<string | null> => {
    try {
        const [first] = await wordTypedRelationshipTargets(ctx, new Set(["comments"]), profile);
        if (first !== undefined) {
            return xmlParts.has(first.part) ? first.part : null;
        }
    } catch {
        // An unreadable package view is reported by its own checks.
    }
    return inFolderByName(ctx, document === null ? "word" : path.posix.dirname(document), "comments.xml");
};

/** The last part named `name` in `folder`, else the OPC-equivalent spelling of `folder/name`. */
const inFolderByName = (ctx: ValidationContext, folder: string, name: string): string | null => {
    let byName: string | null = null;
    for (const part of ctx.partNames) {
        if (path.posix.basename(part) === name && path.posix.dirname(part) === folder) {
            byName = part;
        }
    }
    return byName ?? findEquivalentPart(ctx.partNames, folder === "." ? name : `${folder}/${name}`) ?? null;
};

/**
 * The comment-thread part of exact Microsoft relationship `type` the
 * threading check judges (#696), as #690's repair reads it: the part the
 * main document's relationship of that Type names (unjudged when it is not
 * XML), else the last file named `name` in the comments part's folder —
 * never one from another story root — and absent when that folder holds none.
 */
const threadPart = async (
    ctx: ValidationContext,
    type: string,
    name: string,
    comments: string | null,
    xmlParts: ReadonlySet<string>,
    profile?: Profile,
): Promise<string | null> => {
    if (comments === null) {
        return null;
    }
    try {
        const [first] = await wordExactTypeTargets(ctx, type, profile);
        if (first !== undefined) {
            return xmlParts.has(first) ? first : null;
        }
    } catch {
        // An unreadable package view is reported by its own checks.
    }
    return inFolderByName(ctx, path.posix.dirname(comments), name);
};

// These checks deliberately have no xml-syntax prerequisite: they report
// their own parse failures, and threading still judges document-independent
// metadata when document.xml cannot be parsed.
export const commentMarkersCheck: CheckDefinition = {
    id: "comment-markers",
    description: "Comment range markers pair and reference existing comments.",
    groups: ["comments"],
    formats: ["docx"],
    reportsParseFailures: true,
    run: async (ctx, profile): Promise<ValidationIssue[]> => {
        const issues: ValidationIssue[] = [];

        const documentXml = await mainDocumentPart(ctx, profile);
        const xmlParts = xmlPartNames(ctx.partNames, (part) => ctx.bytes(part));
        const commentsXml = await mainCommentsPart(ctx, documentXml, xmlParts, profile);

        if (!documentXml) {
            // Mirrors Python: "no document.xml" is a pass.
            return [];
        }

        let docDom: Document;
        try {
            docDom = commentXml(ctx, documentXml);
        } catch (err) {
            issues.push({
                severity: "error",
                message: `Error parsing XML: ${err instanceof Error ? err.message : String(err)}`,
                path: documentXml,
                code: "comment-marker-parse",
            });
            return issues;
        }

        const collectIds = (localName: string): Set<string> => {
            const out = new Set<string>();
            const list = getElementsByTagNameNSAll(docDom, WORD_2006_NAMESPACE, localName);
            for (let i = 0; i < list.length; i += 1) {
                const elem = list[i];
                if (!elem) continue;
                const id = elem.getAttributeNS(WORD_2006_NAMESPACE, "id");
                out.add(id ?? "");
            }
            return out;
        };

        const rangeStarts = collectIds("commentRangeStart");
        const rangeEnds = collectIds("commentRangeEnd");
        const references = collectIds("commentReference");

        const orphanedEnds = setDiff(rangeEnds, rangeStarts);
        for (const id of sortIdsNumeric(orphanedEnds)) {
            issues.push({
                severity: "error",
                message: `commentRangeEnd id="${id}" has no matching commentRangeStart`,
                path: documentXml,
                code: "comment-orphan-end",
            });
        }

        const orphanedStarts = setDiff(rangeStarts, rangeEnds);
        for (const id of sortIdsNumeric(orphanedStarts)) {
            issues.push({
                severity: "error",
                message: `commentRangeStart id="${id}" has no matching commentRangeEnd`,
                path: documentXml,
                code: "comment-orphan-start",
            });
        }

        if (commentsXml) {
            let commentsDom: Document;
            try {
                commentsDom = commentXml(ctx, commentsXml);
            } catch (err) {
                issues.push({
                    severity: "error",
                    message: `Error parsing XML: ${err instanceof Error ? err.message : String(err)}`,
                    path: commentsXml,
                    code: "comment-marker-parse",
                });
                return issues;
            }
            const commentIds = new Set<string>();
            const list = getElementsByTagNameNSAll(commentsDom, WORD_2006_NAMESPACE, "comment");
            for (let i = 0; i < list.length; i += 1) {
                const elem = list[i];
                if (!elem) continue;
                const id = elem.getAttributeNS(WORD_2006_NAMESPACE, "id");
                if (id) commentIds.add(id);
            }

            const markerIds = new Set<string>([...rangeStarts, ...rangeEnds, ...references]);
            const invalidRefs = setDiff(markerIds, commentIds);
            for (const id of sortIdsNumeric(invalidRefs)) {
                if (!id) continue;
                const issue: ValidationIssue = {
                    severity: "error",
                    message: `marker id="${id}" references non-existent comment`,
                    path: documentXml,
                    code: "comment-marker-missing",
                };
                if (references.has(id)) {
                    registerCommentReferenceMissing(issue);
                }
                issues.push(issue);
            }
        }

        return issues;
    },
};

export const commentThreadingCheck: CheckDefinition = {
    id: "comment-threading",
    description: "Comment threading paragraph and durable ids agree across comment parts.",
    groups: ["comments"],
    formats: ["docx"],
    reportsParseFailures: true,
    run: async (ctx, profile): Promise<ValidationIssue[]> => {
        const issues: ValidationIssue[] = [];

        const documentXml = await mainDocumentPart(ctx, profile);
        const xmlParts = xmlPartNames(ctx.partNames, (part) => ctx.bytes(part));
        const commentsXml = await mainCommentsPart(ctx, documentXml, xmlParts, profile);
        const commentsExtendedXml = await threadPart(
            ctx,
            COMMENTS_EXTENDED_RELATIONSHIP_TYPE,
            "commentsExtended.xml",
            commentsXml,
            xmlParts,
            profile,
        );
        const commentsIdsXml = await threadPart(ctx, COMMENTS_IDS_RELATIONSHIP_TYPE, "commentsIds.xml", commentsXml, xmlParts, profile);
        const commentsExtensibleXml = await threadPart(
            ctx,
            COMMENTS_EXTENSIBLE_RELATIONSHIP_TYPE,
            "commentsExtensible.xml",
            commentsXml,
            xmlParts,
            profile,
        );

        // No comments.xml → no threading to check.
        if (!commentsXml) return [];

        let commentsDom: Document;
        try {
            commentsDom = commentXml(ctx, commentsXml);
        } catch (err) {
            issues.push({
                severity: "error",
                message: `Error parsing XML: ${err instanceof Error ? err.message : String(err)}`,
                path: commentsXml,
                code: "comment-thread-parse",
            });
            return issues;
        }

        // Build the comment-id → all-paragraph-paraIds map.
        //
        // Word's threading uses the LAST paragraph's paraId (the
        // thread-anchor paragraph) in commentsExtended.xml, not the first.
        // For a single-paragraph comment they're the same; for multi-
        // paragraph comments (typical of replies) only the last paraId is
        // referenced from the extension. So we collect every paraId on
        // every <w:p> inside each <w:comment> and match against the full
        // set rather than guessing first-vs-last.
        //
        // Both Transitional (`schemas.openxmlformats.org/.../2006/main`)
        // and Strict (`purl.oclc.org/ooxml/wordprocessingml/main`)
        // namespace URIs are queried because OOXML Strict-conformance
        // documents use the latter and the rest of the validator pipeline
        // already supports both.
        const commentParaIds = new Map<string, string[]>();
        for (const wNs of WORD_PARAGRAPH_NAMESPACES) {
            const commentList = getElementsByTagNameNSAll(commentsDom, wNs, "comment");
            for (let i = 0; i < commentList.length; i += 1) {
                const elem = commentList[i];
                if (!elem) continue;
                const id = elem.getAttributeNS(wNs, "id");
                if (!id) continue;
                const ps = getElementsByTagNameNSAll(elem, wNs, "p");
                const paraIds: string[] = [];
                for (let j = 0; j < ps.length; j += 1) {
                    const p = ps[j];
                    if (!p) continue;
                    const paraId = p.getAttributeNS(W14_NAMESPACE, "paraId");
                    if (paraId) paraIds.push(paraId);
                }
                if (!commentParaIds.has(id)) commentParaIds.set(id, paraIds);
            }
        }
        const allCommentParaIds = new Set<string>();
        for (const v of commentParaIds.values()) {
            for (const p of v) allCommentParaIds.add(p);
        }

        // ----- rule 4: marker counts ---------------------------------------
        if (documentXml) {
            try {
                const docDom = commentXml(ctx, documentXml);
                const counts = (local: string): number => {
                    let total = 0;
                    for (const ns of WORD_PARAGRAPH_NAMESPACES) {
                        total += getElementsByTagNameNSAll(docDom, ns, local).length;
                    }
                    return total;
                };
                const startCount = counts("commentRangeStart");
                const endCount = counts("commentRangeEnd");
                const refCount = counts("commentReference");
                const expected = commentParaIds.size;
                if (startCount !== expected) {
                    issues.push({
                        severity: "error",
                        message: `commentRangeStart count (${startCount}) does not match ` + `comment count in comments.xml (${expected})`,
                        code: "comment-thread-count-mismatch",
                        path: documentXml,
                    });
                }
                if (endCount !== expected) {
                    issues.push({
                        severity: "error",
                        message: `commentRangeEnd count (${endCount}) does not match ` + `comment count in comments.xml (${expected})`,
                        path: documentXml,
                        code: "comment-thread-count-mismatch",
                    });
                }
                if (refCount !== expected) {
                    issues.push({
                        severity: "error",
                        message: `commentReference count (${refCount}) does not match ` + `comment count in comments.xml (${expected})`,
                        path: documentXml,
                        code: "comment-thread-count-mismatch",
                    });
                }
            } catch {
                // document.xml parse failures are reported elsewhere.
            }
        }

        const commentIdDurableIds = new Set<string>();
        if (commentsIdsXml) {
            let idsDom: Document;
            try {
                idsDom = commentXml(ctx, commentsIdsXml);
            } catch (err) {
                issues.push({
                    severity: "error",
                    message: `Error parsing XML: ${err instanceof Error ? err.message : String(err)}`,
                    path: commentsIdsXml,
                    code: "comment-thread-parse",
                });
                return issues;
            }

            const commentIds = getElementsByTagNameNSAll(idsDom, W16CID_NAMESPACE, "commentId");
            const paraIdCounts = new Map<string, number>();
            const durableIdCounts = new Map<string, number>();
            for (let i = 0; i < commentIds.length; i += 1) {
                const elem = commentIds[i];
                if (!elem) continue;
                const paraId = elem.getAttributeNS(W16CID_NAMESPACE, "paraId");
                const durableId = elem.getAttributeNS(W16CID_NAMESPACE, "durableId");
                if (paraId) {
                    paraIdCounts.set(paraId, (paraIdCounts.get(paraId) ?? 0) + 1);
                    if (!allCommentParaIds.has(paraId)) {
                        issues.push({
                            severity: "error",
                            message:
                                `<w16cid:commentId w16cid:paraId='${paraId}'> does not match any paragraph paraId ` +
                                `in any <w:comment> in comments.xml`,
                            path: commentsIdsXml,
                            code: "comment-thread-commentid-paraid-orphan",
                        });
                    }
                } else {
                    issues.push({
                        severity: "error",
                        message: `<w16cid:commentId> is missing required w16cid:paraId`,
                        path: commentsIdsXml,
                        code: "comment-thread-commentid-missing-paraid",
                    });
                }

                if (durableId) {
                    durableIdCounts.set(durableId, (durableIdCounts.get(durableId) ?? 0) + 1);
                    commentIdDurableIds.add(durableId);
                } else {
                    issues.push({
                        severity: "error",
                        message: `<w16cid:commentId> is missing required w16cid:durableId`,
                        path: commentsIdsXml,
                        code: "comment-thread-commentid-missing-durableid",
                    });
                }
            }

            for (const [paraId, count] of paraIdCounts) {
                if (count > 1) {
                    issues.push({
                        severity: "error",
                        message: `commentsIds.xml has ${count} entries with duplicate paraId='${paraId}'`,
                        path: commentsIdsXml,
                        code: "comment-thread-commentid-duplicate-paraid",
                    });
                }
            }
            for (const [durableId, count] of durableIdCounts) {
                if (count > 1) {
                    issues.push({
                        severity: "error",
                        message: `commentsIds.xml has ${count} entries with duplicate durableId='${durableId}'`,
                        path: commentsIdsXml,
                        code: "comment-thread-commentid-duplicate-durableid",
                    });
                }
            }
        }

        if (commentsExtensibleXml) {
            let cexDom: Document;
            try {
                cexDom = commentXml(ctx, commentsExtensibleXml);
            } catch (err) {
                issues.push({
                    severity: "error",
                    message: `Error parsing XML: ${err instanceof Error ? err.message : String(err)}`,
                    path: commentsExtensibleXml,
                    code: "comment-thread-parse",
                });
                return issues;
            }

            const extensibleEntries = getElementsByTagNameNSAll(cexDom, W16CEX_NAMESPACE, "commentExtensible");
            const durableIdCounts = new Map<string, number>();
            for (let i = 0; i < extensibleEntries.length; i += 1) {
                const elem = extensibleEntries[i];
                if (!elem) continue;
                const durableId = elem.getAttributeNS(W16CEX_NAMESPACE, "durableId");
                if (!durableId) {
                    issues.push({
                        severity: "error",
                        message: `<w16cex:commentExtensible> is missing required w16cex:durableId`,
                        path: commentsExtensibleXml,
                        code: "comment-thread-durableid-missing",
                    });
                    continue;
                }
                durableIdCounts.set(durableId, (durableIdCounts.get(durableId) ?? 0) + 1);
                if (!commentIdDurableIds.has(durableId)) {
                    issues.push({
                        severity: "error",
                        message:
                            `<w16cex:commentExtensible w16cex:durableId='${durableId}'> does not match any ` +
                            `w16cid:durableId in commentsIds.xml`,
                        path: commentsExtensibleXml,
                        code: "comment-thread-durableid-orphan",
                    });
                }
            }
            for (const [durableId, count] of durableIdCounts) {
                if (count > 1) {
                    issues.push({
                        severity: "error",
                        message: `commentsExtensible.xml has ${count} entries with duplicate durableId='${durableId}'`,
                        path: commentsExtensibleXml,
                        code: "comment-thread-durableid-duplicate",
                    });
                }
            }
        }

        // No commentsExtended.xml → rules 1–3 are vacuous, return now.
        if (!commentsExtendedXml) return issues;

        let extDom: Document;
        try {
            extDom = commentXml(ctx, commentsExtendedXml);
        } catch (err) {
            issues.push({
                severity: "error",
                message: `Error parsing XML: ${err instanceof Error ? err.message : String(err)}`,
                path: commentsExtendedXml,
                code: "comment-thread-parse",
            });
            return issues;
        }

        const extEntries = getElementsByTagNameNSAll(extDom, W15_NAMESPACE, "commentEx");
        const extByParaId = new Map<string, number>();
        const extParents: string[] = [];
        for (let i = 0; i < extEntries.length; i += 1) {
            const elem = extEntries[i];
            if (!elem) continue;
            const paraId = elem.getAttributeNS(W15_NAMESPACE, "paraId");
            const parent = elem.getAttributeNS(W15_NAMESPACE, "paraIdParent");
            if (paraId) {
                extByParaId.set(paraId, (extByParaId.get(paraId) ?? 0) + 1);
            }
            if (parent) extParents.push(parent);
        }

        // The full strict validator promotes this soft rule to an error. Other
        // profiles record a warning so the integrity gap is visible
        // without flunking real-world Word-emitted documents that omit
        // commentsExtended entries on standalone (non-threaded) comments.
        const softSeverity: "error" | "warning" = "warning";

        // ----- rule 3: no duplicate paraId entries (jubarte #153) ---------
        // Always an error — duplicates are unambiguous writer bugs.
        for (const [paraId, count] of extByParaId) {
            if (count > 1) {
                issues.push({
                    severity: "error",
                    message: `commentsExtended.xml has ${count} entries with duplicate paraId='${paraId}'`,
                    path: commentsExtendedXml,
                    code: "comment-thread-duplicate-paraid",
                });
            }
        }

        // ----- rule 1: every comment with paraIds has at least one matching commentEx --
        for (const [commentId, paraIds] of commentParaIds) {
            if (paraIds.length === 0) continue;
            if (!paraIds.some((p) => extByParaId.has(p))) {
                issues.push({
                    severity: softSeverity,
                    message:
                        `comment id='${commentId}' has paraIds=[${paraIds.join(", ")}] in comments.xml ` +
                        `but none match a <w15:commentEx w15:paraId='...'> entry in commentsExtended.xml`,
                    path: commentsExtendedXml,
                    code: "comment-thread-paraid-missing",
                });
            }
        }

        // ----- rule 2: every commentEx points at a real comment paragraph --
        for (const paraId of extByParaId.keys()) {
            if (!allCommentParaIds.has(paraId)) {
                issues.push({
                    severity: "error",
                    message:
                        `<w15:commentEx w15:paraId='${paraId}'> does not match any paragraph paraId ` +
                        `in any <w:comment> in comments.xml`,
                    path: commentsExtendedXml,
                    code: "comment-thread-paraid-orphan",
                });
            }
        }

        // ----- rule 3: paraIdParent resolves -------------------------------
        for (const parent of extParents) {
            if (!extByParaId.has(parent)) {
                issues.push({
                    severity: "error",
                    message:
                        `<w15:commentEx w15:paraIdParent='${parent}'> does not resolve to any ` + `paraId declared in commentsExtended.xml`,
                    path: commentsExtendedXml,
                    code: "comment-thread-orphan-parent",
                });
            }
        }

        return issues;
    },
};

function setDiff(a: Set<string>, b: Set<string>): Set<string> {
    const out = new Set<string>();
    for (const v of a) if (!b.has(v)) out.add(v);
    return out;
}

function sortIdsNumeric(ids: Iterable<string>): string[] {
    return [...ids].sort((x, y) => {
        const a = /^\d+$/.test(x) ? Number.parseInt(x, 10) : 0;
        const b = /^\d+$/.test(y) ? Number.parseInt(y, 10) : 0;
        return a - b;
    });
}
