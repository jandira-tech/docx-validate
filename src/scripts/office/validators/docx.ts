/*
 * Copyright 2026 Jandira Technologies, LLC
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *     http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */

/**
 * Validator for Word document XML files against XSD schemas.
 *
 * 1:1 TypeScript port of `src/docx-validate/scripts/office/validators/docx.py`
 * (task #10). Subclasses `BaseSchemaValidator` and adds DOCX-specific checks:
 *
 *   - paragraph counts (original vs. modified, excluding text-box overlays);
 *   - whitespace preservation on `<w:t>` runs that lead/trail with whitespace;
 *   - tracked-changes nesting (no `<w:t>` inside `<w:del>`, no `<w:delText>`
 *     inside `<w:ins>`);
 *   - comment-marker pairing (`commentRangeStart`/`End`/`commentReference`);
 *   - id constraints (paraId < 0x80000000, durableId < 0x7FFFFFFF — plain
 *     32-bit numbers, well inside Number.MAX_SAFE_INTEGER, no BigInt needed);
 *   - durableId auto-repair when a value blows past the constraint.
 *
 * All `validate*` methods follow the base-class shape: return
 * `ValidationResult` rather than printing + returning bool. The Python source
 * also exposed `compare_paragraph_counts` (purely informational, prints the
 * delta) — that role is filled here by `compareParagraphCounts` returning a
 * `{ original, modified, delta, originalUsesStrictNamespace }` object so
 * callers can render it however they like.
 *
 * @xmldom gotcha: never write `nodeValue` on a Text node — assign via `.data`.
 * This file only mutates element attributes (setAttribute), so we're safe;
 * see `lib/xml-helpers.ts` for the convention.
 */

import { promises as fs, readFileSync } from "node:fs";
import { default as JSZip } from "jszip";
import type { ValidationIssue, ValidationResult } from "../../../lib/types";
import { mergeResults } from "../../../lib/types";
import { makeSelect, parseXml, serializeXml } from "../../../lib/xml-helpers";
import { BaseSchemaValidator, collectDeclaredPrefixes, XML_NAMESPACE } from "./base";

export const WORD_2006_NAMESPACE = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
export const WORD_STRICT_NAMESPACE = "http://purl.oclc.org/ooxml/wordprocessingml/main";
/** Python: `WORD_PARAGRAPH_NAMESPACES = (WORD_2006_NAMESPACE, WORD_STRICT_NAMESPACE)` */
export const WORD_PARAGRAPH_NAMESPACES: readonly [string, string] = [WORD_2006_NAMESPACE, WORD_STRICT_NAMESPACE];
const W14_NAMESPACE = "http://schemas.microsoft.com/office/word/2010/wordml";
const W15_NAMESPACE = "http://schemas.microsoft.com/office/word/2012/wordml";
const W16CID_NAMESPACE = "http://schemas.microsoft.com/office/word/2016/wordml/cid";
const VML_NAMESPACE = "urn:schemas-microsoft-com:vml";

/**
 * Well-known OOXML namespace prefixes Word emits in `mc:Ignorable` and
 * declares on the root element. When `repairIgnorable` finds an undeclared
 * prefix in `mc:Ignorable`, it looks up the URI here and adds
 * `xmlns:prefix="<uri>"` to the document root rather than dropping the
 * Ignorable entry — that mirrors what Word does on save and preserves
 * tolerated-extension semantics for any `<prefix:*>` elements actually
 * present in the body. Prefixes not in this table fall back to drop-from-
 * Ignorable as a safe last resort.
 *
 * Source: Microsoft Office OOXML schema headers + ECMA-376 part 4.
 */
const KNOWN_OOXML_PREFIX_URIS: Readonly<Record<string, string>> = {
    w14: W14_NAMESPACE,
    w15: W15_NAMESPACE,
    w16: "http://schemas.microsoft.com/office/word/2018/wordml",
    w16cex: "http://schemas.microsoft.com/office/word/2018/wordml/cex",
    w16cid: W16CID_NAMESPACE,
    w16du: "http://schemas.microsoft.com/office/word/2023/wordml/word16du",
    w16sdtdh: "http://schemas.microsoft.com/office/word/2020/wordml/sdtdatahash",
    w16sdtfl: "http://schemas.microsoft.com/office/word/2024/wordml/sdtformatlock",
    w16se: "http://schemas.microsoft.com/office/word/2015/wordml/symex",
    wp14: "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing",
    wpc: "http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas",
    wpg: "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup",
    wpi: "http://schemas.microsoft.com/office/word/2010/wordprocessingInk",
    wps: "http://schemas.microsoft.com/office/word/2010/wordprocessingShape",
    wne: "http://schemas.microsoft.com/office/word/2006/wordml",
    cx: "http://schemas.microsoft.com/office/drawing/2014/chartex",
    cx1: "http://schemas.microsoft.com/office/drawing/2015/9/8/chartex",
    cx2: "http://schemas.microsoft.com/office/drawing/2015/10/21/chartex",
    cx3: "http://schemas.microsoft.com/office/drawing/2016/5/9/chartex",
    cx4: "http://schemas.microsoft.com/office/drawing/2016/5/10/chartex",
    cx5: "http://schemas.microsoft.com/office/drawing/2016/5/11/chartex",
    cx6: "http://schemas.microsoft.com/office/drawing/2016/5/12/chartex",
    cx7: "http://schemas.microsoft.com/office/drawing/2016/5/13/chartex",
    cx8: "http://schemas.microsoft.com/office/drawing/2016/5/14/chartex",
    aink: "http://schemas.microsoft.com/office/drawing/2016/ink",
    am3d: "http://schemas.microsoft.com/office/drawing/2017/model3d",
    oel: "http://schemas.microsoft.com/office/2019/extlst",
};

/**
 * Wire-format token prefixes used by HTML-to-DOCX pipelines (e.g. jubarte,
 * Plate's tracked-changes plugin) to round-trip insertions, deletions, and
 * comments through HTML. They MUST be expanded into proper OOXML before the
 * archive ships — any leak indicates the writer skipped a transformation
 * pass. The prefixes are deliberately matched verbatim (not regex-anchored
 * to a specific producer) so the check generalises across producers.
 */
const TRACKING_TOKEN_PREFIXES: readonly string[] = [
    "[[DOCX_INS_START:",
    "[[DOCX_INS_END:",
    "[[DOCX_DEL_START:",
    "[[DOCX_DEL_END:",
    "[[DOCX_CMT_START:",
    "[[DOCX_CMT_END:",
    "[[DOCX_PMARK_DEL:",
    "[[DOCX_PMARK_INS:",
];
const TRACKING_TOKEN_REGEX =
    /\[\[DOCX_(?:INS|DEL|CMT)_(?:START|END):[^\]]*?\]\]|\[\[DOCX_PMARK_(?:DEL|INS):[^\]]*?\]\]/g;

// XPath excluding `<w:p>` inside DML/VML text-box overlays. Matches the
// Python BODY_PARAGRAPH_XPATH.
const BODY_PARAGRAPH_XPATH =
    ".//w:p[not(ancestor::w:txbxContent) and not(ancestor::v:textbox)] " +
    "| " +
    ".//strict:p[not(ancestor::strict:txbxContent) and not(ancestor::v:textbox)]";

const MAX_PARA_ID = 0x80000000;
const MAX_DURABLE_ID = 0x7fffffff;
const MAX_RANDOM_DURABLE = 0x7ffffffe;

export interface ParagraphCounts {
    original: number;
    modified: number;
    delta: number;
    originalUsesStrictNamespace: boolean;
}

export class DOCXSchemaValidator extends BaseSchemaValidator {
    protected readonly elementRelationshipTypes: Record<string, string> = {};

    /**
     * Cached parse of `word/document.xml` from the original `.docx` zip.
     * `null` means "not yet attempted"; `false` means "attempted and failed,
     * don't try again".
     */
    private originalDocumentRoot: Element | null = null;
    private originalDocumentRootFailed = false;

    // ----- top-level entry point ----------------------------------------------

    /**
     * Run every check. `compareParagraphCounts` is informational only and
     * does not affect the merged validity.
     */
    async validate(): Promise<ValidationResult> {
        const xmlOk = await this.validateXml();
        if (!xmlOk.valid) return xmlOk;

        const results = await Promise.all([
            this.validateNamespaces(),
            this.validateUniqueIds(),
            this.validateFileReferences(),
            this.validateRelationshipElements(),
            this.validateContentTypes(),
            this.validateAgainstXsd(),
            this.validateWhitespacePreservation(),
            this.validateDeletions(),
            this.validateInsertions(),
            this.validateAllRelationshipIds(),
            this.validateIdConstraints(),
            this.validateCommentMarkers(),
            this.validateCommentThreading(),
            this.validateNoTrackingTokens(),
            this.validateAllParagraphsHaveParaId(),
            this.validateNoBom(),
            this.validateNoEmptyRelsParts(),
        ]);

        return mergeResults(xmlOk, ...results);
    }

    // ----- whitespace ---------------------------------------------------------

    async validateWhitespacePreservation(): Promise<ValidationResult> {
        const issues: ValidationIssue[] = [];
        for (const xmlFile of this.documentXmlFiles()) {
            let dom: Document;
            try {
                dom = parseXml(await fs.readFile(xmlFile, "utf-8"));
            } catch (err) {
                issues.push({
                    severity: "error",
                    message: `Error: ${err instanceof Error ? err.message : String(err)}`,
                    path: this.relPath(xmlFile),
                    code: "ws-parse",
                });
                continue;
            }
            const tElems = dom.getElementsByTagNameNS(WORD_2006_NAMESPACE, "t");
            for (let i = 0; i < tElems.length; i += 1) {
                const elem = tElems.item(i);
                if (!elem) continue;
                const first = elem.firstChild;
                if (!first || first.nodeType !== 3 /* TEXT_NODE */) continue;
                const text = first.nodeValue ?? "";
                if (!text) continue;
                if (!/^[ \t\n\r]/.test(text) && !/[ \t\n\r]$/.test(text)) continue;
                const xmlSpace = elem.getAttributeNS(XML_NAMESPACE, "space");
                if (xmlSpace === "preserve") continue;
                const preview = previewRepr(text, 50);
                issues.push({
                    severity: "error",
                    message: `w:t element with whitespace missing xml:space='preserve': ${preview}`,
                    path: this.relPath(xmlFile),
                    code: "ws-missing-preserve",
                });
            }
        }
        return finalize(issues);
    }

    // ----- tracked changes ----------------------------------------------------

    async validateDeletions(): Promise<ValidationResult> {
        const issues: ValidationIssue[] = [];
        const $$ = makeSelect();
        for (const xmlFile of this.documentXmlFiles()) {
            let dom: Document;
            try {
                dom = parseXml(await fs.readFile(xmlFile, "utf-8"));
            } catch (err) {
                issues.push({
                    severity: "error",
                    message: `Error: ${err instanceof Error ? err.message : String(err)}`,
                    path: this.relPath(xmlFile),
                    code: "del-parse",
                });
                continue;
            }
            // <w:t> inside <w:del>
            const tInDel = $$(".//w:del//w:t", dom) as Node[];
            for (const node of tInDel) {
                const elem = node as Element;
                const text = elem.firstChild?.nodeValue ?? "";
                issues.push({
                    severity: "error",
                    message: `<w:t> found within <w:del>: ${previewRepr(text, 50)}`,
                    path: this.relPath(xmlFile),
                    code: "del-contains-t",
                });
            }
            // <w:instrText> inside <w:del>
            const instrInDel = $$(".//w:del//w:instrText", dom) as Node[];
            for (const node of instrInDel) {
                const elem = node as Element;
                const text = elem.firstChild?.nodeValue ?? "";
                issues.push({
                    severity: "error",
                    message: `<w:instrText> found within <w:del> (use <w:delInstrText>): ${previewRepr(text, 50)}`,
                    path: this.relPath(xmlFile),
                    code: "del-contains-instrtext",
                });
            }
        }
        return finalize(issues);
    }

    async validateInsertions(): Promise<ValidationResult> {
        const issues: ValidationIssue[] = [];
        const $$ = makeSelect();
        for (const xmlFile of this.documentXmlFiles()) {
            let dom: Document;
            try {
                dom = parseXml(await fs.readFile(xmlFile, "utf-8"));
            } catch (err) {
                issues.push({
                    severity: "error",
                    message: `Error: ${err instanceof Error ? err.message : String(err)}`,
                    path: this.relPath(xmlFile),
                    code: "ins-parse",
                });
                continue;
            }
            const invalid = $$(".//w:ins//w:delText[not(ancestor::w:del)]", dom) as Node[];
            for (const node of invalid) {
                const elem = node as Element;
                const text = elem.firstChild?.nodeValue ?? "";
                issues.push({
                    severity: "error",
                    message: `<w:delText> within <w:ins>: ${previewRepr(text, 50)}`,
                    path: this.relPath(xmlFile),
                    code: "ins-contains-deltext",
                });
            }
        }
        return finalize(issues);
    }

    // ----- comment markers ----------------------------------------------------

    async validateCommentMarkers(): Promise<ValidationResult> {
        const issues: ValidationIssue[] = [];

        let documentXml: string | null = null;
        let commentsXml: string | null = null;
        for (const xmlFile of this.xmlFiles) {
            const base = baseName(xmlFile);
            if (base === "document.xml" && xmlFile.includes("word")) {
                documentXml = xmlFile;
            } else if (base === "comments.xml") {
                commentsXml = xmlFile;
            }
        }

        if (!documentXml) {
            // Mirrors Python: "no document.xml" is a pass.
            return { valid: true, issues: [] };
        }

        let docDom: Document;
        try {
            docDom = parseXml(await fs.readFile(documentXml, "utf-8"));
        } catch (err) {
            issues.push({
                severity: "error",
                message: `Error parsing XML: ${err instanceof Error ? err.message : String(err)}`,
                path: this.relPath(documentXml),
                code: "comment-marker-parse",
            });
            return finalize(issues);
        }

        const collectIds = (localName: string): Set<string> => {
            const out = new Set<string>();
            const list = docDom.getElementsByTagNameNS(WORD_2006_NAMESPACE, localName);
            for (let i = 0; i < list.length; i += 1) {
                const elem = list.item(i);
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
                path: "document.xml",
                code: "comment-orphan-end",
            });
        }

        const orphanedStarts = setDiff(rangeStarts, rangeEnds);
        for (const id of sortIdsNumeric(orphanedStarts)) {
            issues.push({
                severity: "error",
                message: `commentRangeStart id="${id}" has no matching commentRangeEnd`,
                path: "document.xml",
                code: "comment-orphan-start",
            });
        }

        if (commentsXml) {
            let commentsDom: Document;
            try {
                commentsDom = parseXml(await fs.readFile(commentsXml, "utf-8"));
            } catch (err) {
                issues.push({
                    severity: "error",
                    message: `Error parsing XML: ${err instanceof Error ? err.message : String(err)}`,
                    path: this.relPath(commentsXml),
                    code: "comment-marker-parse",
                });
                return finalize(issues);
            }
            const commentIds = new Set<string>();
            const list = commentsDom.getElementsByTagNameNS(WORD_2006_NAMESPACE, "comment");
            for (let i = 0; i < list.length; i += 1) {
                const elem = list.item(i);
                if (!elem) continue;
                const id = elem.getAttributeNS(WORD_2006_NAMESPACE, "id");
                if (id) commentIds.add(id);
            }

            const markerIds = new Set<string>([...rangeStarts, ...rangeEnds, ...references]);
            const invalidRefs = setDiff(markerIds, commentIds);
            for (const id of sortIdsNumeric(invalidRefs)) {
                if (!id) continue;
                issues.push({
                    severity: "error",
                    message: `marker id="${id}" references non-existent comment`,
                    path: "document.xml",
                    code: "comment-marker-missing",
                });
            }
        }

        return finalize(issues);
    }

    // ----- paraId completeness (Word's tracking-anchor convention) ------------

    /**
     * Profile-aware: every `<w:p>` AND every `<w:tr>` should carry a
     * `w14:paraId` attribute. The OOXML schema makes these optional, but
     * Word stamps one on every paragraph and table row to use as the
     * anchor for tracked-changes infrastructure (e.g. a deleted row's
     * revision recorded by paraId, or a comment range that spans an
     * inserted paragraph).
     *
     * Severity:
     *   - `strict`  → `error` (spec-purist + Word-parity).
     *   - `lenient` → `warning` (ours-vs-Word stylistic difference; the
     *     schema allows it, but downstream collaboration tooling may
     *     need it).
     *
     * Surfaced by the comparison against Word's `reallyrepaired.docx`
     * for the sample-document fixture: jubarte's writer leaves table
     * rows without a paraId; Word stamps all 13 of them.
     */
    async validateAllParagraphsHaveParaId(): Promise<ValidationResult> {
        const issues: ValidationIssue[] = [];
        const severity: "error" | "warning" = this.profile === "strict" ? "error" : "warning";
        for (const xmlFile of this.documentXmlFiles()) {
            let dom: Document;
            try {
                dom = parseXml(await fs.readFile(xmlFile, "utf-8"));
            } catch {
                continue;
            }
            for (const local of ["p", "tr"] as const) {
                let missing = 0;
                for (const ns of WORD_PARAGRAPH_NAMESPACES) {
                    const list = dom.getElementsByTagNameNS(ns, local);
                    for (let i = 0; i < list.length; i += 1) {
                        const elem = list.item(i);
                        if (!elem) continue;
                        const paraId = elem.getAttributeNS(W14_NAMESPACE, "paraId");
                        if (!paraId) missing += 1;
                    }
                }
                if (missing > 0) {
                    issues.push({
                        severity,
                        message:
                            `${missing} <w:${local}> element(s) lack a w14:paraId. ` +
                            `Word stamps one on every paragraph and table row as the anchor for ` +
                            `tracked-changes / comment-range infrastructure.`,
                        path: this.relPath(xmlFile),
                        code: "paraid-missing-element",
                    });
                }
            }
        }
        return finalize(issues);
    }

    // ----- comment threading --------------------------------------------------

    /**
     * Verify the cross-references between `word/comments.xml` and
     * `word/commentsExtended.xml` (issue #153 in jubarte was a regression in
     * this exact integrity contract).
     *
     * This is a STRICT validator. Some Word builds emit a comments.xml
     * paraId that has no matching `<w15:commentEx>` entry (notably for
     * non-threaded comments) — we still flag that here, because callers
     * that want to ignore Word's output divergences should run validation
     * with `profile: "lenient"` (where the strict-only `code` is downgraded
     * to a `warning`).
     *
     * Rules enforced (severity = error in `strict`, warning in `lenient`):
     *   1. Every `<w:comment>` in comments.xml that carries a `w14:paraId`
     *      on its first paragraph must have a matching `<w15:commentEx>`
     *      entry in commentsExtended.xml.
     *   2. Every `<w15:commentEx>` must have a `w15:paraId` matching the
     *      paraId of some `<w:comment>` first paragraph (the inverse of #1).
     *   3. No two `<w15:commentEx>` entries may share the same `w15:paraId`
     *      (this was the surface symptom of jubarte's #153 regression).
     *   4. Every non-null `w15:paraIdParent` must resolve to a paraId
     *      present somewhere else in commentsExtended.xml.
     *   5. The number of `w:commentRangeStart`, `w:commentRangeEnd`, and
     *      `w:commentReference` elements in document.xml must each equal
     *      the number of comments in comments.xml.
     *
     * `validateCommentMarkers` already covers orphan range start/end and
     * missing-comment references — this validator is strictly about the
     * threaded-comments extension surface.
     */
    async validateCommentThreading(): Promise<ValidationResult> {
        const issues: ValidationIssue[] = [];

        let documentXml: string | null = null;
        let commentsXml: string | null = null;
        let commentsExtendedXml: string | null = null;
        for (const xmlFile of this.xmlFiles) {
            const base = baseName(xmlFile);
            if (base === "document.xml" && xmlFile.includes("word")) {
                documentXml = xmlFile;
            } else if (base === "comments.xml") {
                commentsXml = xmlFile;
            } else if (base === "commentsExtended.xml") {
                commentsExtendedXml = xmlFile;
            }
        }

        // No comments.xml → no threading to check.
        if (!commentsXml) return { valid: true, issues: [] };

        let commentsDom: Document;
        try {
            commentsDom = parseXml(await fs.readFile(commentsXml, "utf-8"));
        } catch (err) {
            issues.push({
                severity: "error",
                message: `Error parsing XML: ${err instanceof Error ? err.message : String(err)}`,
                path: this.relPath(commentsXml),
                code: "comment-thread-parse",
            });
            return finalize(issues);
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
            const commentList = commentsDom.getElementsByTagNameNS(wNs, "comment");
            for (let i = 0; i < commentList.length; i += 1) {
                const elem = commentList.item(i);
                if (!elem) continue;
                const id = elem.getAttributeNS(wNs, "id");
                if (!id) continue;
                const ps = elem.getElementsByTagNameNS(wNs, "p");
                const paraIds: string[] = [];
                for (let j = 0; j < ps.length; j += 1) {
                    const p = ps.item(j);
                    if (!p) continue;
                    const paraId = p.getAttributeNS(W14_NAMESPACE, "paraId");
                    if (paraId) paraIds.push(paraId);
                }
                if (!commentParaIds.has(id)) commentParaIds.set(id, paraIds);
            }
        }

        // ----- rule 4: marker counts ---------------------------------------
        if (documentXml) {
            try {
                const docDom = parseXml(await fs.readFile(documentXml, "utf-8"));
                const counts = (local: string): number => {
                    let total = 0;
                    for (const ns of WORD_PARAGRAPH_NAMESPACES) {
                        total += docDom.getElementsByTagNameNS(ns, local).length;
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
                        message:
                            `commentRangeStart count (${startCount}) does not match ` +
                            `comment count in comments.xml (${expected})`,
                        path: "document.xml",
                        code: "comment-thread-count-mismatch",
                    });
                }
                if (endCount !== expected) {
                    issues.push({
                        severity: "error",
                        message:
                            `commentRangeEnd count (${endCount}) does not match ` +
                            `comment count in comments.xml (${expected})`,
                        path: "document.xml",
                        code: "comment-thread-count-mismatch",
                    });
                }
                if (refCount !== expected) {
                    issues.push({
                        severity: "error",
                        message:
                            `commentReference count (${refCount}) does not match ` +
                            `comment count in comments.xml (${expected})`,
                        path: "document.xml",
                        code: "comment-thread-count-mismatch",
                    });
                }
            } catch {
                // document.xml parse failures are reported elsewhere.
            }
        }

        // No commentsExtended.xml → rules 1–3 are vacuous, return now.
        if (!commentsExtendedXml) return finalize(issues);

        let extDom: Document;
        try {
            extDom = parseXml(await fs.readFile(commentsExtendedXml, "utf-8"));
        } catch (err) {
            issues.push({
                severity: "error",
                message: `Error parsing XML: ${err instanceof Error ? err.message : String(err)}`,
                path: this.relPath(commentsExtendedXml),
                code: "comment-thread-parse",
            });
            return finalize(issues);
        }

        const extEntries = extDom.getElementsByTagNameNS(W15_NAMESPACE, "commentEx");
        const extByParaId = new Map<string, number>();
        const extParents: string[] = [];
        for (let i = 0; i < extEntries.length; i += 1) {
            const elem = extEntries.item(i);
            if (!elem) continue;
            const paraId = elem.getAttributeNS(W15_NAMESPACE, "paraId");
            const parent = elem.getAttributeNS(W15_NAMESPACE, "paraIdParent");
            if (paraId) {
                extByParaId.set(paraId, (extByParaId.get(paraId) ?? 0) + 1);
            }
            if (parent) extParents.push(parent);
        }

        // Severity for soft-strict rules: strict profile reports an error,
        // lenient profile records a warning so the integrity gap is visible
        // without flunking real-world Word-emitted documents that omit
        // commentsExtended entries on standalone (non-threaded) comments.
        const softSeverity: "error" | "warning" = this.profile === "strict" ? "error" : "warning";

        // ----- rule 3: no duplicate paraId entries (jubarte #153) ---------
        // Always an error — duplicates are unambiguous writer bugs.
        for (const [paraId, count] of extByParaId) {
            if (count > 1) {
                issues.push({
                    severity: "error",
                    message: `commentsExtended.xml has ${count} entries with duplicate paraId='${paraId}'`,
                    path: this.relPath(commentsExtendedXml),
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
                    path: this.relPath(commentsExtendedXml),
                    code: "comment-thread-paraid-missing",
                });
            }
        }

        // ----- rule 2: every commentEx points at a real comment paragraph --
        const allCommentParaIds = new Set<string>();
        for (const v of commentParaIds.values()) {
            for (const p of v) allCommentParaIds.add(p);
        }
        for (const paraId of extByParaId.keys()) {
            if (!allCommentParaIds.has(paraId)) {
                issues.push({
                    severity: "error",
                    message:
                        `<w15:commentEx w15:paraId='${paraId}'> does not match any paragraph paraId ` +
                        `in any <w:comment> in comments.xml`,
                    path: this.relPath(commentsExtendedXml),
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
                        `<w15:commentEx w15:paraIdParent='${parent}'> does not resolve to any ` +
                        `paraId declared in commentsExtended.xml`,
                    path: this.relPath(commentsExtendedXml),
                    code: "comment-thread-orphan-parent",
                });
            }
        }

        return finalize(issues);
    }

    // ----- tracking-token leak detection --------------------------------------

    /**
     * Detect HTML-to-DOCX wire-format tracking tokens (e.g.
     * `[[DOCX_INS_START:{...}]]`) that leaked into the OOXML output. These
     * are placeholder strings emitted by html-to-docx-style pipelines for
     * tracked-changes round-tripping; they MUST be expanded into proper
     * `<w:ins>` / `<w:del>` / comment markers before serialisation. A leak
     * is a writer-pipeline bug — the document will render the literal
     * `[[DOCX_…]]` text in Word.
     *
     * Only scans the document XML files (document.xml + headers / footers /
     * footnotes / endnotes), not the relationship sidecars or content
     * types — those parts can never legitimately contain user-visible text.
     */
    async validateNoTrackingTokens(): Promise<ValidationResult> {
        const issues: ValidationIssue[] = [];
        for (const xmlFile of this.userTextXmlFiles()) {
            let raw: string;
            try {
                raw = await fs.readFile(xmlFile, "utf-8");
            } catch {
                continue;
            }
            // Cheap pre-filter — if no `[[DOCX_` substring, skip the regex.
            if (!TRACKING_TOKEN_PREFIXES.some((p) => raw.includes(p))) continue;

            const tokenRegex = new RegExp(TRACKING_TOKEN_REGEX);
            let match: RegExpExecArray | null;
            const seen = new Set<string>();
            while ((match = tokenRegex.exec(raw)) !== null) {
                const token = match[0];
                if (seen.has(token)) continue;
                seen.add(token);
                issues.push({
                    severity: "error",
                    message: `Wire-format tracking token leaked into OOXML output: ${previewRepr(token, 80)}`,
                    path: this.relPath(xmlFile),
                    code: "tracking-token-leak",
                });
            }
        }
        return finalize(issues);
    }

    // ----- id constraints -----------------------------------------------------

    async validateIdConstraints(): Promise<ValidationResult> {
        const issues: ValidationIssue[] = [];
        for (const xmlFile of this.xmlFiles) {
            let dom: Document;
            try {
                dom = parseXml(await fs.readFile(xmlFile, "utf-8"));
            } catch {
                // Mirrors Python's bare except — silently skip.
                continue;
            }
            const all = dom.getElementsByTagName("*");
            const base = baseName(xmlFile);
            for (let i = 0; i < all.length; i += 1) {
                const elem = all.item(i);
                if (!elem) continue;

                const paraId = elem.getAttributeNS(W14_NAMESPACE, "paraId");
                if (paraId) {
                    if (parseIdValue(paraId, 16) >= MAX_PARA_ID) {
                        issues.push({
                            severity: "error",
                            message: `${base}: paraId=${paraId} >= 0x80000000`,
                            path: this.relPath(xmlFile),
                            code: "id-paraid-overflow",
                        });
                    }
                }

                const durableId = elem.getAttributeNS(W16CID_NAMESPACE, "durableId");
                if (durableId) {
                    if (base === "numbering.xml") {
                        const v = parseIdValue(durableId, 10);
                        if (Number.isNaN(v)) {
                            issues.push({
                                severity: "error",
                                message: `${base}: durableId=${durableId} must be decimal in numbering.xml`,
                                path: this.relPath(xmlFile),
                                code: "id-durable-decimal",
                            });
                        } else if (v >= MAX_DURABLE_ID) {
                            issues.push({
                                severity: "error",
                                message: `${base}: durableId=${durableId} >= 0x7FFFFFFF`,
                                path: this.relPath(xmlFile),
                                code: "id-durable-overflow",
                            });
                        }
                    } else {
                        if (parseIdValue(durableId, 16) >= MAX_DURABLE_ID) {
                            issues.push({
                                severity: "error",
                                message: `${base}: durableId=${durableId} >= 0x7FFFFFFF`,
                                path: this.relPath(xmlFile),
                                code: "id-durable-overflow",
                            });
                        }
                    }
                }
            }
        }
        return finalize(issues);
    }

    // ----- paragraph counts (informational) -----------------------------------

    countParagraphsInUnpacked(): number {
        let count = 0;
        for (const xmlFile of this.xmlFiles) {
            if (baseName(xmlFile) !== "document.xml") continue;
            try {
                const dom = parseXml(readFileSync(xmlFile, "utf-8"));
                count = countParagraphsInRoot(dom);
            } catch {
                // mirrors Python catch-and-print; we just swallow
            }
        }
        return count;
    }

    async countParagraphsInOriginal(): Promise<number> {
        if (!this.originalFile) return 0;
        const root = await this.loadOriginalDocumentRoot();
        if (!root) return 0;
        try {
            // Wrap the root element in its owner document for xpath evaluation.
            const doc = root.ownerDocument!;
            return countParagraphsInRoot(doc);
        } catch {
            return 0;
        }
    }

    async compareParagraphCounts(): Promise<ParagraphCounts> {
        const original = await this.countParagraphsInOriginal();
        const modified = this.countParagraphsInUnpacked();
        const delta = modified - original;
        const strict = await this.originalUsesStrictNamespace();
        if (this.verbose) {
            const diffStr = delta > 0 ? `+${delta}` : String(delta);
            process.stdout.write(`\nParagraphs: ${original} → ${modified} (${diffStr})\n`);
            if (strict) {
                process.stdout.write(
                    "Note: input document uses the ECMA-376 Strict OOXML " + "conformance class; output uses Transitional.\n",
                );
            }
        }
        return { original, modified, delta, originalUsesStrictNamespace: strict };
    }

    private async loadOriginalDocumentRoot(): Promise<Element | null> {
        if (this.originalDocumentRoot) return this.originalDocumentRoot;
        if (this.originalDocumentRootFailed) return null;
        if (!this.originalFile) {
            this.originalDocumentRootFailed = true;
            return null;
        }
        try {
            const buf = await fs.readFile(this.originalFile);
            const zip = await JSZip.loadAsync(buf);
            const entry = zip.file("word/document.xml");
            if (!entry) {
                this.originalDocumentRootFailed = true;
                return null;
            }
            const text = await entry.async("text");
            const dom = parseXml(text);
            this.originalDocumentRoot = dom.documentElement;
            return this.originalDocumentRoot;
        } catch {
            this.originalDocumentRootFailed = true;
            return null;
        }
    }

    private async originalUsesStrictNamespace(): Promise<boolean> {
        const root = await this.loadOriginalDocumentRoot();
        if (!root) return false;
        return root.namespaceURI === WORD_STRICT_NAMESPACE;
    }

    // ----- repair -------------------------------------------------------------

    async repair(): Promise<number> {
        const baseRepairs = await super.repair();
        const durableRepairs = await this.repairDurableId();
        const paraIdRepairs = await this.repairParaId();
        const ignorableRepairs = await this.repairIgnorable();
        const stampRepairs = await this.repairMissingParaIds();
        return baseRepairs + durableRepairs + paraIdRepairs + ignorableRepairs + stampRepairs;
    }

    /**
     * Stamp a fresh in-range `w14:paraId` on every `<w:p>` and `<w:tr>`
     * that lacks one. The new IDs are random and avoid colliding with any
     * paraId already present on either side (incl. the renumbered ones
     * from `repairParaId`).
     *
     * Without this pass, Word tolerates the file but its tracked-changes
     * infrastructure has no anchor on the affected elements — a comment
     * spanning an inserted row may not survive a save round-trip.
     */
    async repairMissingParaIds(): Promise<number> {
        let repairs = 0;
        const usedIds = new Set<string>();

        // First pass: collect every existing paraId across all document
        // XML parts so the new ones we generate can't collide.
        for (const xmlFile of this.documentXmlFiles()) {
            try {
                const dom = parseXml(await fs.readFile(xmlFile, "utf-8"));
                const all = dom.getElementsByTagName("*");
                for (let i = 0; i < all.length; i += 1) {
                    const elem = all.item(i);
                    if (!elem) continue;
                    const id = elem.getAttributeNS(W14_NAMESPACE, "paraId");
                    if (id) usedIds.add(id.toUpperCase());
                }
            } catch {
                // pass
            }
        }

        const allocate = (): string => {
            for (;;) {
                const value = 1 + Math.floor(Math.random() * (MAX_PARA_ID - 1));
                const newId = value.toString(16).toUpperCase().padStart(8, "0");
                if (!usedIds.has(newId)) {
                    usedIds.add(newId);
                    return newId;
                }
            }
        };

        // Second pass: stamp.
        for (const xmlFile of this.documentXmlFiles()) {
            try {
                const dom = parseXml(await fs.readFile(xmlFile, "utf-8"));
                let modified = false;
                for (const local of ["p", "tr"] as const) {
                    for (const ns of WORD_PARAGRAPH_NAMESPACES) {
                        const list = dom.getElementsByTagNameNS(ns, local);
                        for (let i = 0; i < list.length; i += 1) {
                            const elem = list.item(i);
                            if (!elem) continue;
                            const paraId = elem.getAttributeNS(W14_NAMESPACE, "paraId");
                            if (paraId) continue;
                            elem.setAttributeNS(W14_NAMESPACE, "w14:paraId", allocate());
                            repairs += 1;
                            modified = true;
                        }
                    }
                }
                if (modified) {
                    await fs.writeFile(xmlFile, serializeXml(dom, "UTF-8"), "utf-8");
                }
            } catch {
                // pass
            }
        }
        return repairs;
    }

    /**
     * Repair `mc:Ignorable` entries that name an undeclared namespace
     * prefix. For each undeclared prefix:
     *
     *   - If the prefix is in `KNOWN_OOXML_PREFIX_URIS` (the table of
     *     well-known Word/OOXML namespaces), declare it on the document
     *     root: `xmlns:prefix="<canonical-uri>"`. Mirrors what Word does
     *     on save and preserves the tolerated-extension semantics for any
     *     `<prefix:*>` elements present elsewhere in the document.
     *   - Otherwise (truly unknown prefix), drop it from the
     *     `mc:Ignorable` token list. If that empties the attribute, the
     *     attribute is removed entirely.
     *
     * Returns the count of mutations made (one per declaration added or
     * Ignorable token dropped).
     */
    async repairIgnorable(): Promise<number> {
        let repairs = 0;
        for (const xmlFile of this.xmlFiles) {
            try {
                const dom = parseXml(await fs.readFile(xmlFile, "utf-8"));
                const root = dom.documentElement;
                if (!root) continue;

                // Find every element on the root carrying an Ignorable attr —
                // the MC namespace gives `mc:Ignorable`, but the lxml-style
                // collector in `validateNamespaces` matches anything ending
                // in `Ignorable` so we mirror that to stay symmetric with
                // the validator.
                const declared = collectDeclaredPrefixes(root);
                const attrs = root.attributes;
                const ignorableAttrs: Array<{ name: string; value: string }> = [];
                for (let i = 0; i < attrs.length; i += 1) {
                    const attr = attrs.item(i);
                    if (!attr) continue;
                    if (!attr.name.endsWith("Ignorable")) continue;
                    ignorableAttrs.push({ name: attr.name, value: attr.value });
                }
                if (ignorableAttrs.length === 0) continue;

                let perFileMutations = 0;
                for (const { name, value } of ignorableAttrs) {
                    const tokens = value.split(/\s+/).filter(Boolean);
                    const surviving: string[] = [];
                    for (const prefix of tokens) {
                        if (declared.has(prefix)) {
                            surviving.push(prefix);
                            continue;
                        }
                        const uri = KNOWN_OOXML_PREFIX_URIS[prefix];
                        if (uri !== undefined) {
                            // Declare on the root and keep the Ignorable entry.
                            root.setAttribute(`xmlns:${prefix}`, uri);
                            declared.add(prefix);
                            surviving.push(prefix);
                            perFileMutations += 1;
                            repairs += 1;
                        } else {
                            // Truly unknown — drop from Ignorable rather than
                            // guessing a URI.
                            perFileMutations += 1;
                            repairs += 1;
                        }
                    }
                    const next = surviving.join(" ");
                    if (next === value) continue;
                    if (next.length === 0) {
                        root.removeAttribute(name);
                    } else {
                        root.setAttribute(name, next);
                    }
                }

                if (perFileMutations > 0) {
                    await fs.writeFile(xmlFile, serializeXml(dom, "UTF-8"), "utf-8");
                }
            } catch {
                // swallow — corrupted XML files surface elsewhere
            }
        }
        return repairs;
    }

    async repairDurableId(): Promise<number> {
        let repairs = 0;
        for (const xmlFile of this.xmlFiles) {
            try {
                const content = await fs.readFile(xmlFile, "utf-8");
                const dom = parseXml(content);
                let modified = false;
                const base = baseName(xmlFile);
                const all = dom.getElementsByTagName("*");
                for (let i = 0; i < all.length; i += 1) {
                    const elem = all.item(i);
                    if (!elem) continue;
                    const durableId = elem.getAttributeNS(W16CID_NAMESPACE, "durableId");
                    if (!durableId) continue;

                    let needsRepair: boolean;
                    if (base === "numbering.xml") {
                        const v = parseIdValue(durableId, 10);
                        needsRepair = Number.isNaN(v) || v >= MAX_DURABLE_ID;
                    } else {
                        const v = parseIdValue(durableId, 16);
                        needsRepair = Number.isNaN(v) || v >= MAX_DURABLE_ID;
                    }

                    if (needsRepair) {
                        const value = 1 + Math.floor(Math.random() * MAX_RANDOM_DURABLE);
                        const newId = base === "numbering.xml" ? String(value) : value.toString(16).toUpperCase().padStart(8, "0");
                        // setAttributeNS keeps the prefix binding intact.
                        elem.setAttributeNS(W16CID_NAMESPACE, "w16cid:durableId", newId);
                        repairs += 1;
                        modified = true;
                    }
                }
                if (modified) {
                    await fs.writeFile(xmlFile, serializeXml(dom, "UTF-8"), "utf-8");
                }
            } catch {
                // swallow — mirrors Python bare except
            }
        }
        return repairs;
    }

    /**
     * Repair every `w14:paraId` whose value is `>= 0x80000000` (the spec
     * cap), and propagate the renumbering into the threaded-comments
     * extension.
     *
     * Two-pass implementation:
     *
     *  Pass 1: scan all XML files, build a `Map<oldId, newId>` for every
     *          over-cap `w14:paraId`. Each old ID gets one new value
     *          (deterministic per-run via random allocation, with a
     *          collision-avoidance set so two different over-cap IDs can't
     *          both land on the same fresh value).
     *  Pass 2: walk every XML file again and rewrite (a) every
     *          `w14:paraId` referenced in the map, AND (b) every
     *          `w15:paraId` / `w15:paraIdParent` in `commentsExtended.xml`
     *          that names an old value. Without (b), the cross-reference
     *          between `comments.xml` and `commentsExtended.xml` breaks
     *          and Word's threading view shows no replies — surfaced by
     *          `validateCommentThreading`'s `comment-thread-paraid-orphan`
     *          check.
     */
    async repairParaId(): Promise<number> {
        let repairs = 0;
        const remap = new Map<string, string>();
        const usedNewIds = new Set<string>();

        const allocateNewId = (): string => {
            // Match Python's random.randint(1, 0x7FFFFFFE) range and
            // retry on the (astronomically rare) collision against an
            // already-allocated repair value.
            for (;;) {
                const value = 1 + Math.floor(Math.random() * (MAX_PARA_ID - 1));
                const newId = value.toString(16).toUpperCase().padStart(8, "0");
                if (!usedNewIds.has(newId)) {
                    usedNewIds.add(newId);
                    return newId;
                }
            }
        };

        // Pass 1: discover every over-cap w14:paraId and pick its replacement.
        const parsedDoms = new Map<string, Document>();
        for (const xmlFile of this.xmlFiles) {
            try {
                const dom = parseXml(await fs.readFile(xmlFile, "utf-8"));
                parsedDoms.set(xmlFile, dom);
                const all = dom.getElementsByTagName("*");
                for (let i = 0; i < all.length; i += 1) {
                    const elem = all.item(i);
                    if (!elem) continue;
                    const paraId = elem.getAttributeNS(W14_NAMESPACE, "paraId");
                    if (!paraId) continue;
                    if (remap.has(paraId)) continue;
                    const v = parseIdValue(paraId, 16);
                    if (Number.isNaN(v) || v >= MAX_PARA_ID) {
                        remap.set(paraId, allocateNewId());
                    }
                }
            } catch {
                // pass — corrupted XML files surface elsewhere.
            }
        }

        // Also pre-allocate replacements for any w15:paraId / paraIdParent
        // values that are over-cap *and* don't appear as a w14:paraId
        // (rare — would happen if commentsExtended.xml has a stale entry).
        // Keeping these in the same remap means cross-file references stay
        // self-consistent.
        for (const [xmlFile, dom] of parsedDoms) {
            if (baseName(xmlFile) !== "commentsExtended.xml") continue;
            const exts = dom.getElementsByTagNameNS(W15_NAMESPACE, "commentEx");
            for (let i = 0; i < exts.length; i += 1) {
                const elem = exts.item(i);
                if (!elem) continue;
                for (const attr of ["paraId", "paraIdParent"] as const) {
                    const val = elem.getAttributeNS(W15_NAMESPACE, attr);
                    if (!val) continue;
                    if (remap.has(val)) continue;
                    const v = parseIdValue(val, 16);
                    if (Number.isNaN(v) || v >= MAX_PARA_ID) {
                        remap.set(val, allocateNewId());
                    }
                }
            }
        }

        if (remap.size === 0) return 0;

        // Pass 2: rewrite all references.
        for (const [xmlFile, dom] of parsedDoms) {
            try {
                let modified = false;
                const all = dom.getElementsByTagName("*");
                for (let i = 0; i < all.length; i += 1) {
                    const elem = all.item(i);
                    if (!elem) continue;
                    const w14Para = elem.getAttributeNS(W14_NAMESPACE, "paraId");
                    if (w14Para && remap.has(w14Para)) {
                        elem.setAttributeNS(W14_NAMESPACE, "w14:paraId", remap.get(w14Para)!);
                        repairs += 1;
                        modified = true;
                    }
                    const w15Para = elem.getAttributeNS(W15_NAMESPACE, "paraId");
                    if (w15Para && remap.has(w15Para)) {
                        elem.setAttributeNS(W15_NAMESPACE, "w15:paraId", remap.get(w15Para)!);
                        repairs += 1;
                        modified = true;
                    }
                    const w15Parent = elem.getAttributeNS(W15_NAMESPACE, "paraIdParent");
                    if (w15Parent && remap.has(w15Parent)) {
                        elem.setAttributeNS(W15_NAMESPACE, "w15:paraIdParent", remap.get(w15Parent)!);
                        repairs += 1;
                        modified = true;
                    }
                }
                if (modified) {
                    await fs.writeFile(xmlFile, serializeXml(dom, "UTF-8"), "utf-8");
                }
            } catch {
                // swallow — mirrors Python bare except in repairDurableId
            }
        }
        return repairs;
    }

    // ----- internal helpers ---------------------------------------------------

    private *documentXmlFiles(): IterableIterator<string> {
        for (const f of this.xmlFiles) {
            if (baseName(f) === "document.xml") yield f;
        }
    }

    /**
     * XML parts that can hold user-visible text and therefore can leak
     * wire-format tracking tokens: the body, headers, footers, footnotes,
     * endnotes, comments, and glossary part. Excludes settings / styles /
     * rels / numbering since they never contain user text.
     */
    private *userTextXmlFiles(): IterableIterator<string> {
        for (const f of this.xmlFiles) {
            const name = baseName(f);
            if (name === "document.xml") yield f;
            else if (name === "comments.xml") yield f;
            else if (name === "footnotes.xml") yield f;
            else if (name === "endnotes.xml") yield f;
            else if (/^header\d*\.xml$/i.test(name)) yield f;
            else if (/^footer\d*\.xml$/i.test(name)) yield f;
        }
    }
}

// ===== module-level helpers ===================================================

function finalize(issues: ValidationIssue[]): ValidationResult {
    return { valid: issues.every((i) => i.severity !== "error"), issues };
}

function baseName(p: string): string {
    const idx = Math.max(p.lastIndexOf("/"), p.lastIndexOf("\\"));
    return idx >= 0 ? p.slice(idx + 1) : p;
}

function previewRepr(text: string, max: number): string {
    // Mirrors Python's `repr(text)` truncation: wraps in single quotes and
    // escapes embedded backslashes / quotes the same way repr() does.
    const repr = pythonRepr(text);
    return repr.length > max ? `${repr.slice(0, max)}...` : repr;
}

function pythonRepr(s: string): string {
    let out = "";
    for (const ch of s) {
        if (ch === "\\") out += "\\\\";
        else if (ch === "'") out += "\\'";
        else if (ch === "\n") out += "\\n";
        else if (ch === "\r") out += "\\r";
        else if (ch === "\t") out += "\\t";
        else out += ch;
    }
    return `'${out}'`;
}

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

/**
 * Python's `int(val, base)` returns a value or throws ValueError on garbage.
 * We return NaN for "could not parse" so callers can branch.
 */
function parseIdValue(val: string, base: number): number {
    const trimmed = val.trim();
    if (!trimmed) return Number.NaN;
    // Reject anything that's not pure digits / hex chars for the requested base.
    const re = base === 16 ? /^[+-]?[0-9A-Fa-f]+$/ : /^[+-]?[0-9]+$/;
    if (!re.test(trimmed)) return Number.NaN;
    const v = Number.parseInt(trimmed, base);
    return Number.isFinite(v) ? v : Number.NaN;
}

function countParagraphsInRoot(doc: Document): number {
    const $$ = makeSelect({ strict: WORD_STRICT_NAMESPACE, v: VML_NAMESPACE });
    const nodes = $$(BODY_PARAGRAPH_XPATH, doc) as Node[];
    return nodes.length;
}
