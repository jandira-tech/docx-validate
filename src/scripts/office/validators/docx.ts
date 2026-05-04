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

import { parseXml, serializeXml, makeSelect } from "../../../lib/xml-helpers.ts";
import { mergeResults } from "../../../lib/types.ts";
import type { ValidationIssue, ValidationResult } from "../../../lib/types.ts";
import { BaseSchemaValidator, XML_NAMESPACE } from "./base.ts";

export const WORD_2006_NAMESPACE = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
export const WORD_STRICT_NAMESPACE = "http://purl.oclc.org/ooxml/wordprocessingml/main";
/** Python: `WORD_PARAGRAPH_NAMESPACES = (WORD_2006_NAMESPACE, WORD_STRICT_NAMESPACE)` */
export const WORD_PARAGRAPH_NAMESPACES: readonly [string, string] = [WORD_2006_NAMESPACE, WORD_STRICT_NAMESPACE];
const W14_NAMESPACE = "http://schemas.microsoft.com/office/word/2010/wordml";
const W16CID_NAMESPACE = "http://schemas.microsoft.com/office/word/2016/wordml/cid";
const VML_NAMESPACE = "urn:schemas-microsoft-com:vml";

// XPath excluding `<w:p>` inside DML/VML text-box overlays. Matches the
// Python BODY_PARAGRAPH_XPATH.
const BODY_PARAGRAPH_XPATH =
    ".//w:p[not(ancestor::w:txbxContent) and not(ancestor::v:textbox)] " +
    "| " +
    ".//strict:p[not(ancestor::strict:txbxContent) and not(ancestor::v:textbox)]";

const MAX_PARA_ID = 0x80000000;
const MAX_DURABLE_ID = 0x7fffffff;
const MAX_RANDOM_DURABLE = 0x7ffffffe;

interface ParagraphCounts {
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
            this.validateNoBom(),
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
        return baseRepairs + durableRepairs + paraIdRepairs;
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

    async repairParaId(): Promise<number> {
        let repairs = 0;
        for (const xmlFile of this.xmlFiles) {
            try {
                const content = await fs.readFile(xmlFile, "utf-8");
                const dom = parseXml(content);
                let modified = false;
                const all = dom.getElementsByTagName("*");
                for (let i = 0; i < all.length; i += 1) {
                    const elem = all.item(i);
                    if (!elem) continue;
                    const paraId = elem.getAttributeNS(W14_NAMESPACE, "paraId");
                    if (!paraId) continue;

                    const v = parseIdValue(paraId, 16);
                    if (Number.isNaN(v) || v >= MAX_PARA_ID) {
                        // Safe range: [1, 0x7FFFFFFF]. Match Python's random.randint(1, 0x7FFFFFFE).
                        const value = 1 + Math.floor(Math.random() * (MAX_PARA_ID - 1));
                        const newId = value.toString(16).toUpperCase().padStart(8, "0");
                        elem.setAttributeNS(W14_NAMESPACE, "w14:paraId", newId);
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
