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
 * Validator for PowerPoint presentation XML files against XSD schemas.
 *
 * Port of `src/docx-validate/scripts/office/validators/pptx.py` (task #11).
 * Subclass of `BaseSchemaValidator` (task #7) — composes the shared checks
 * with three PPTX-specific ones: UUID-shaped IDs are well-formed, slide
 * masters' `sldLayoutId` r:id values point at real slideLayout relationships,
 * each slide has at most one slideLayout reference, and notes slides aren't
 * shared between multiple slides.
 */

import { existsSync, promises as fs } from "node:fs";
import path from "node:path";

import type { ValidationIssue, ValidationResult } from "../../../lib/types";
import { mergeResults } from "../../../lib/types";
import { parseXml } from "../../../lib/xml-helpers";
import type { BaseSchemaValidatorOptions } from "./base";
import { BaseSchemaValidator, OFFICE_RELATIONSHIPS_NAMESPACE, PACKAGE_RELATIONSHIPS_NAMESPACE, walkFiles } from "./base";

export const PRESENTATIONML_NAMESPACE = "http://schemas.openxmlformats.org/presentationml/2006/main";

const PPTX_ELEMENT_RELATIONSHIP_TYPES: Record<string, string> = {
    sldid: "slide",
    sldmasterid: "slidemaster",
    notesmasterid: "notesmaster",
    sldlayoutid: "slidelayout",
    themeid: "theme",
    tablestyleid: "tablestyles",
};

const UUID_PATTERN = /^[{(]?[0-9A-Fa-f]{8}-?[0-9A-Fa-f]{4}-?[0-9A-Fa-f]{4}-?[0-9A-Fa-f]{4}-?[0-9A-Fa-f]{12}[})]?$/;

export class PPTXSchemaValidator extends BaseSchemaValidator {
    protected readonly elementRelationshipTypes = PPTX_ELEMENT_RELATIONSHIP_TYPES;

    constructor(opts: BaseSchemaValidatorOptions) {
        super(opts);
    }

    /**
     * Run the full PPTX check suite. Mirrors `PPTXSchemaValidator.validate()`
     * in Python — same order, same short-circuit on `validateXml` failure.
     */
    async validate(): Promise<ValidationResult> {
        const xml = await this.validateXml();
        if (!xml.valid) {
            return xml;
        }

        const checks = [
            xml,
            await this.validateNamespaces(),
            await this.validateUniqueIds(),
            await this.validateUuidIds(),
            await this.validateFileReferences(),
            await this.validateRelationshipElements(),
            await this.validateSlideLayoutIds(),
            await this.validateContentTypes(),
            await this.validateAgainstXsd(),
            await this.validateNotesSlideReferences(),
            await this.validateAllRelationshipIds(),
            await this.validateNoDuplicateSlideLayouts(),
            await this.validateNoBom(),
            await this.validateNoEmptyRelsParts(),
        ];
        return mergeResults(...checks);
    }

    /**
     * Reject any attribute whose name ends in "id" and whose value *looks* like
     * a UUID (32 alphanumeric chars after stripping `{}()-`) but isn't valid
     * hex. Mirrors `validate_uuid_ids` — `_looks_like_uuid` accepts non-hex
     * alnum chars so the regex check has something to fail.
     */
    async validateUuidIds(): Promise<ValidationResult> {
        const issues: ValidationIssue[] = [];

        for (const xmlFile of this.xmlFiles) {
            let dom: Document;
            try {
                dom = parseXml(await fs.readFile(xmlFile, "utf-8"));
            } catch (err) {
                issues.push({
                    severity: "error",
                    message: `Error: ${err instanceof Error ? err.message : String(err)}`,
                    path: this.relPath(xmlFile),
                    code: "uuid-parse",
                });
                continue;
            }

            const root = dom.documentElement;
            if (!root) continue;

            for (const elem of iterElements(root)) {
                const attrs = elem.attributes;
                for (let i = 0; i < attrs.length; i += 1) {
                    const attr = attrs.item(i);
                    if (!attr) continue;
                    const attrLocal = attr.name.split("}").pop()?.toLowerCase() ?? "";
                    if (attrLocal !== "id" && !attrLocal.endsWith("id")) continue;
                    if (!looksLikeUuid(attr.value)) continue;
                    if (UUID_PATTERN.test(attr.value)) continue;

                    issues.push({
                        severity: "error",
                        message: `ID '${attr.value}' appears to be a UUID but contains invalid hex characters`,
                        path: this.relPath(xmlFile),
                        code: "uuid-invalid",
                    });
                }
            }
        }

        return finalize(issues);
    }

    /**
     * Verify each `sldLayoutId/@r:id` inside a slide master resolves to a
     * Relationship of type slideLayout in the master's `.rels` sidecar.
     * Mirrors `validate_slide_layout_ids`.
     */
    async validateSlideLayoutIds(): Promise<ValidationResult> {
        const issues: ValidationIssue[] = [];

        const slideMastersDir = path.join(this.unpackedDir, "ppt", "slideMasters");
        if (!existsSync(slideMastersDir)) {
            return { valid: true, issues: [] };
        }

        const slideMasters = walkFiles(slideMastersDir, [".xml"]).filter((f) => {
            const rel = path.relative(this.unpackedDir, f);
            const parts = rel.split(path.sep);
            // glob "ppt/slideMasters/*.xml" — one level only, no _rels subdir.
            return parts.length === 3 && parts[0] === "ppt" && parts[1] === "slideMasters" && parts[2].endsWith(".xml");
        });

        if (slideMasters.length === 0) {
            return { valid: true, issues: [] };
        }

        for (const slideMaster of slideMasters) {
            let masterDom: Document;
            try {
                masterDom = parseXml(await fs.readFile(slideMaster, "utf-8"));
            } catch (err) {
                issues.push({
                    severity: "error",
                    message: `Error: ${err instanceof Error ? err.message : String(err)}`,
                    path: this.relPath(slideMaster),
                    code: "sldlayout-parse",
                });
                continue;
            }

            const masterDir = path.dirname(slideMaster);
            const masterBase = path.basename(slideMaster);
            const relsFile = path.join(masterDir, "_rels", `${masterBase}.rels`);

            if (!existsSync(relsFile)) {
                issues.push({
                    severity: "error",
                    message: `Missing relationships file: ${this.relPath(relsFile)}`,
                    path: this.relPath(slideMaster),
                    code: "sldlayout-missing-rels",
                });
                continue;
            }

            let relsDom: Document;
            try {
                relsDom = parseXml(await fs.readFile(relsFile, "utf-8"));
            } catch (err) {
                issues.push({
                    severity: "error",
                    message: `Error: ${err instanceof Error ? err.message : String(err)}`,
                    path: this.relPath(slideMaster),
                    code: "sldlayout-rels-parse",
                });
                continue;
            }

            const validLayoutRids = new Set<string>();
            const rels = relsDom.getElementsByTagNameNS(PACKAGE_RELATIONSHIPS_NAMESPACE, "Relationship");
            for (let i = 0; i < rels.length; i += 1) {
                const rel = rels.item(i);
                if (!rel) continue;
                const relType = rel.getAttribute("Type") ?? "";
                if (relType.includes("slideLayout")) {
                    const id = rel.getAttribute("Id");
                    if (id) validLayoutRids.add(id);
                }
            }

            const sldLayoutIds = masterDom.getElementsByTagNameNS(PRESENTATIONML_NAMESPACE, "sldLayoutId");
            for (let i = 0; i < sldLayoutIds.length; i += 1) {
                const sli = sldLayoutIds.item(i);
                if (!sli) continue;
                const rId = sli.getAttributeNS(OFFICE_RELATIONSHIPS_NAMESPACE, "id");
                const layoutId = sli.getAttribute("id");
                if (rId && !validLayoutRids.has(rId)) {
                    issues.push({
                        severity: "error",
                        message:
                            `sldLayoutId with id='${layoutId}' references r:id='${rId}' ` +
                            `which is not found in slide layout relationships`,
                        path: this.relPath(slideMaster),
                        code: "sldlayout-missing-rid",
                    });
                }
            }
        }

        return finalize(issues);
    }

    /**
     * Each slide should reference exactly one slideLayout. Mirrors
     * `validate_no_duplicate_slide_layouts`.
     */
    async validateNoDuplicateSlideLayouts(): Promise<ValidationResult> {
        const issues: ValidationIssue[] = [];

        const slideRelsDir = path.join(this.unpackedDir, "ppt", "slides", "_rels");
        if (!existsSync(slideRelsDir)) {
            return { valid: true, issues: [] };
        }

        const relsFiles = walkFiles(slideRelsDir, [".rels"]).filter((f) => {
            const rel = path.relative(this.unpackedDir, f);
            const parts = rel.split(path.sep);
            return (
                parts.length === 4 && parts[0] === "ppt" && parts[1] === "slides" && parts[2] === "_rels" && parts[3].endsWith(".xml.rels")
            );
        });

        for (const relsFile of relsFiles) {
            let dom: Document;
            try {
                dom = parseXml(await fs.readFile(relsFile, "utf-8"));
            } catch (err) {
                issues.push({
                    severity: "error",
                    message: `Error: ${err instanceof Error ? err.message : String(err)}`,
                    path: this.relPath(relsFile),
                    code: "slide-rels-parse",
                });
                continue;
            }

            const rels = dom.getElementsByTagNameNS(PACKAGE_RELATIONSHIPS_NAMESPACE, "Relationship");
            let layoutCount = 0;
            for (let i = 0; i < rels.length; i += 1) {
                const rel = rels.item(i);
                if (!rel) continue;
                const relType = rel.getAttribute("Type") ?? "";
                if (relType.includes("slideLayout")) layoutCount += 1;
            }

            if (layoutCount > 1) {
                issues.push({
                    severity: "error",
                    message: `has ${layoutCount} slideLayout references`,
                    path: this.relPath(relsFile),
                    code: "slide-layout-duplicate",
                });
            }
        }

        return finalize(issues);
    }

    /**
     * Notes slides may be optional — but if present, each must back exactly
     * one parent slide. Mirrors `validate_notes_slide_references`.
     */
    async validateNotesSlideReferences(): Promise<ValidationResult> {
        const issues: ValidationIssue[] = [];
        const notesReferences = new Map<string, Array<{ slideName: string; relsFile: string }>>();

        const slideRelsDir = path.join(this.unpackedDir, "ppt", "slides", "_rels");
        if (!existsSync(slideRelsDir)) {
            return { valid: true, issues: [] };
        }

        const relsFiles = walkFiles(slideRelsDir, [".rels"]).filter((f) => {
            const rel = path.relative(this.unpackedDir, f);
            const parts = rel.split(path.sep);
            return (
                parts.length === 4 && parts[0] === "ppt" && parts[1] === "slides" && parts[2] === "_rels" && parts[3].endsWith(".xml.rels")
            );
        });

        if (relsFiles.length === 0) {
            return { valid: true, issues: [] };
        }

        for (const relsFile of relsFiles) {
            let dom: Document;
            try {
                dom = parseXml(await fs.readFile(relsFile, "utf-8"));
            } catch (err) {
                issues.push({
                    severity: "error",
                    message: `Error: ${err instanceof Error ? err.message : String(err)}`,
                    path: this.relPath(relsFile),
                    code: "notes-rels-parse",
                });
                continue;
            }

            const rels = dom.getElementsByTagNameNS(PACKAGE_RELATIONSHIPS_NAMESPACE, "Relationship");
            for (let i = 0; i < rels.length; i += 1) {
                const rel = rels.item(i);
                if (!rel) continue;
                const relType = rel.getAttribute("Type") ?? "";
                if (!relType.includes("notesSlide")) continue;
                const target = rel.getAttribute("Target");
                if (!target) continue;
                const normalizedTarget = target.replace(/\.\.\//g, "");
                // Python: rels_file.stem (foo.xml.rels → foo.xml) then .replace(".xml", "") → foo
                const baseStem = path.basename(relsFile, ".rels"); // "foo.xml"
                const slideName = baseStem.replace(/\.xml$/i, "");
                const bucket = notesReferences.get(normalizedTarget) ?? [];
                bucket.push({ slideName, relsFile });
                notesReferences.set(normalizedTarget, bucket);
            }
        }

        for (const [target, references] of notesReferences.entries()) {
            if (references.length <= 1) continue;
            const slideNames = references.map((r) => r.slideName).join(", ");
            issues.push({
                severity: "error",
                message: `Notes slide '${target}' is referenced by multiple slides: ${slideNames}`,
                code: "notes-shared",
            });
            for (const ref of references) {
                issues.push({
                    severity: "error",
                    message: `    - ${this.relPath(ref.relsFile)}`,
                    path: this.relPath(ref.relsFile),
                    code: "notes-shared-detail",
                });
            }
        }

        return finalize(issues);
    }
}

function looksLikeUuid(value: string): boolean {
    const cleaned = value.replace(/[{}()]/g, "").replace(/-/g, "");
    if (cleaned.length !== 32) return false;
    for (let i = 0; i < cleaned.length; i += 1) {
        const code = cleaned.charCodeAt(i);
        const isAlnum =
            (code >= 0x30 && code <= 0x39) /* 0-9 */ ||
            (code >= 0x41 && code <= 0x5a) /* A-Z */ ||
            (code >= 0x61 && code <= 0x7a) /* a-z */;
        if (!isAlnum) return false;
    }
    return true;
}

function* iterElements(root: Element | null): IterableIterator<Element> {
    if (!root) return;
    yield root;
    const stack: Element[] = [root];
    while (stack.length > 0) {
        const node = stack.pop()!;
        const children = node.childNodes;
        for (let i = children.length - 1; i >= 0; i -= 1) {
            const child = children.item(i);
            if (child && child.nodeType === 1) {
                yield child as Element;
                stack.push(child as Element);
            }
        }
    }
}

function finalize(issues: ValidationIssue[]): ValidationResult {
    return { valid: issues.every((i) => i.severity !== "error"), issues };
}
