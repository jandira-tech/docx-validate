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

import { canonicalPartName, findEquivalentPart, relationshipTargetFolder, relationshipTargetPartName } from "../../../lib/content-types";
import type { ValidationIssue, ValidationResult } from "../../../lib/types";
import { mergeResults } from "../../../lib/types";
import { officeRelationshipTypeSegment } from "../../../lib/word-read-parts";
import { getElementsByTagNameNSAll, parseXmlBytes } from "../../../lib/xml-helpers";
import type { BaseSchemaValidatorOptions } from "./base";
import { BaseSchemaValidator, OFFICE_RELATIONSHIPS_NAMESPACE, packageEntries, withoutRestatedIgnorablePrefixes } from "./base";
import { openDirectoryContext, packageExtractFailedIssue, storedNamesRun } from "../package-context";

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
     * in Python — same order, same short-circuit on `validateXml` failure,
     * except that word-valid and strict still judge the content-types stream.
     */
    async validate(): Promise<ValidationResult> {
        return this.withDeadlineFinding(await this.withinBaselineRun(() => this.runEveryCheck()));
    }

    private async runEveryCheck(): Promise<ValidationResult> {
        this.resetExecutedChecks();
        const opened = openDirectoryContext(this.unpackedDir);
        if (opened.openError !== undefined) {
            const diagnostic = await storedNamesRun(opened.context, "pptx", this.profile, this.deadline);
            this.recordExecutedChecks(diagnostic.executed);
            return { valid: false, issues: [...diagnostic.issues, packageExtractFailedIssue(this.unpackedDir, opened.openError)] };
        }
        const context = opened.context;
        const rule = this.ruleScheduler();
        const xml = await rule("PPTX validateXml", () => this.validateXml());
        const malformed = new Set(
            xml.issues.flatMap((issue) => (issue.severity === "error" && issue.path !== undefined ? [issue.path] : [])),
        );
        if (!xml.valid) {
            // Lenient stops at a malformed part, and so do its partial runs
            // (#391). Word-valid and strict partial runs still judge the
            // content-types stream, so the full run reports what they do
            // rather than dropping it (#456).
            if (this.profile === "lenient") return xml;
            return withoutRestatedIgnorablePrefixes(
                mergeResults(
                    xml,
                    await rule("PPTX validateDuplicateDeclarations", () => this.validateDuplicateDeclarations(context, malformed)),
                    await rule("PPTX validateOverridePartNames", () => this.validateOverridePartNames(context, "pptx", malformed)),
                    await rule("PPTX validateContentTypesMarkupCompatibility", () =>
                        this.validateContentTypesMarkupCompatibility(context, "pptx", malformed),
                    ),
                ),
            );
        }

        const checks = [
            xml,
            await rule("PPTX validateNamespaces", () => this.validateNamespaces()),
            await rule("PPTX validateUniqueIds", () => this.validateUniqueIds()),
            await rule("PPTX validateUuidIds", () => this.validateUuidIds()),
            await rule("PPTX validateFileReferences", () => this.validateFileReferences()),
            await rule("PPTX validateRelationshipElements", () => this.validateRelationshipElements()),
            await rule("PPTX validateSlideLayoutIds", () => this.validateSlideLayoutIds()),
            await rule("PPTX validateContentTypes", () => this.validateContentTypes()),
            await rule("PPTX validateOverridePartNames", () => this.validateOverridePartNames(context, "pptx", malformed)),
            await rule("PPTX validateContentTypesMarkupCompatibility", () =>
                this.validateContentTypesMarkupCompatibility(context, "pptx", malformed),
            ),
            await rule("PPTX validateAgainstXsd", () => this.validateAgainstXsd()),
            await rule("PPTX validateNotesSlideReferences", () => this.validateNotesSlideReferences()),
            await rule("PPTX validateAllRelationshipIds", () => this.validateAllRelationshipIds()),
            await rule("PPTX validateNoDuplicateSlideLayouts", () => this.validateNoDuplicateSlideLayouts()),
            await rule("PPTX validateNoBom", () => this.validateNoBom()),
            await rule("PPTX validateNoEmptyRelsParts", () => this.validateNoEmptyRelsParts()),
        ];
        return withoutRestatedIgnorablePrefixes(mergeResults(...checks));
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
                dom = parseXmlBytes(await fs.readFile(xmlFile));
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
                    const attr = attrs[i];
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

        // glob "ppt/slideMasters/*.xml" — one level only, no _rels subdir,
        // ASCII case-insensitively (#1017) — and every master the
        // presentation relates, wherever stored (#1013).
        const slideMasters = (await this.judgedParts(/^ppt\/slideMasters\/[^/]+\.xml$/i, "slideMaster")).map((part) => part.file);

        if (slideMasters.length === 0) {
            return { valid: true, issues: [] };
        }

        for (const slideMaster of slideMasters) {
            let masterDom: Document;
            try {
                masterDom = parseXmlBytes(await fs.readFile(slideMaster));
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
                relsDom = parseXmlBytes(await fs.readFile(relsFile));
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
            const rels = this.relationshipElements(relsDom, this.entryName(relsFile));
            for (let i = 0; i < rels.length; i += 1) {
                const rel = rels[i];
                if (!rel) continue;
                const relType = rel.getAttribute("Type") ?? "";
                if (relType.includes("slideLayout")) {
                    const id = rel.getAttribute("Id");
                    if (id) validLayoutRids.add(id);
                }
            }

            const sldLayoutIds = getElementsByTagNameNSAll(masterDom, PRESENTATIONML_NAMESPACE, "sldLayoutId");
            for (let i = 0; i < sldLayoutIds.length; i += 1) {
                const sli = sldLayoutIds[i];
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

        for (const { file: relsFile } of await this.slideRelationshipsParts()) {
            let dom: Document;
            try {
                dom = parseXmlBytes(await fs.readFile(relsFile));
            } catch (err) {
                issues.push({
                    severity: "error",
                    message: `Error: ${err instanceof Error ? err.message : String(err)}`,
                    path: this.relPath(relsFile),
                    code: "slide-rels-parse",
                });
                continue;
            }

            const rels = this.relationshipElements(dom, this.entryName(relsFile));
            let layoutCount = 0;
            for (let i = 0; i < rels.length; i += 1) {
                const rel = rels[i];
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
        // Keyed by the canonical name of the part the targets name, and
        // reported under the first name a slide gave it (#1007).
        const notesReferences = new Map<string, { partName: string; references: Array<{ slideName: string; relsFile: string }> }>();

        const relsFiles = await this.slideRelationshipsParts();

        if (relsFiles.length === 0) {
            return { valid: true, issues: [] };
        }

        for (const { file: relsFile, folder } of relsFiles) {
            let dom: Document;
            try {
                dom = parseXmlBytes(await fs.readFile(relsFile));
            } catch (err) {
                issues.push({
                    severity: "error",
                    message: `Error: ${err instanceof Error ? err.message : String(err)}`,
                    path: this.relPath(relsFile),
                    code: "notes-rels-parse",
                });
                continue;
            }

            const rels = this.relationshipElements(dom, this.entryName(relsFile));
            for (let i = 0; i < rels.length; i += 1) {
                const rel = rels[i];
                if (!rel) continue;
                const relType = rel.getAttribute("Type") ?? "";
                if (!relType.includes("notesSlide")) continue;
                const target = rel.getAttribute("Target");
                if (!target) continue;
                // The target names the part the read model reads: a climb
                // stops at the package root, a query or fragment is dropped,
                // `\` reads as `/`, and equivalent spellings of one part share
                // a key, so two slides sharing one notes slide are found
                // however each spells it (#1007). A URI or an empty reference
                // names no part, and a target naming a folder names none
                // either (#962, #975), so it shares no notes slide with the
                // part its path spells (#991). Each target resolves against
                // its slide's folder, `ppt/slides` for the historical sidecars
                // (#1013).
                const resolved = relationshipTargetPartName(folder, target);
                if (resolved === undefined || relationshipTargetFolder(folder, target) !== undefined) continue;
                const partName = resolved.replace(/^\//, "");
                const key = canonicalPartName(partName);
                // Python: rels_file.stem (foo.xml.rels → foo.xml) then .replace(".xml", "") → foo
                const baseStem = path.basename(relsFile, ".rels"); // "foo.xml"
                const slideName = baseStem.replace(/\.xml$/i, "");
                const entry = notesReferences.get(key) ?? { partName, references: [] };
                entry.references.push({ slideName, relsFile });
                notesReferences.set(key, entry);
            }
        }

        for (const { partName, references } of notesReferences.values()) {
            if (references.length <= 1) continue;
            const slideNames = references.map((r) => r.slideName).join(", ");
            issues.push({
                severity: "error",
                message: `Notes slide '${partName}' is referenced by multiple slides: ${slideNames}`,
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

    /**
     * The relationships parts of the slides the slide checks judge, each
     * stored name once, in code-unit order, with the folder of the slide
     * whose targets they hold: every stored `ppt/slides/_rels/*.xml.rels`, as
     * the Python port listed them, and the sidecar of every slide the
     * presentation relates, wherever stored (#1013). The historical pattern
     * matches ASCII case-insensitively, as OPC compares part names, and every
     * stored spelling of two equivalent names is judged, as the file-system
     * walk before #1013 judged them (#1017). A sidecar's folder is the parent
     * of its `_rels` folder, as its stored name spells it.
     */
    private async slideRelationshipsParts(): Promise<Array<{ file: string; folder: string }>> {
        const entries = packageEntries(this.unpackedDir);
        const names = [...entries.keys()];
        const judged = new Set(names.filter((name) => /^ppt\/slides\/_rels\/[^/]+\.xml\.rels$/i.test(name)));
        for (const slide of await this.presentationRelatedParts(names, "slide")) {
            const name = findEquivalentPart(
                names,
                path.posix.join(path.posix.dirname(slide), "_rels", `${path.posix.basename(slide)}.rels`),
            );
            if (name !== undefined) judged.add(name);
        }
        return [...judged]
            .sort((a, b) => (a < b ? -1 : a > b ? 1 : 0))
            .map((name) => ({ file: entries.get(name)!, folder: path.posix.dirname(path.posix.dirname(name)) }));
    }

    /**
     * The stored parts a check judges, each stored name once, in code-unit
     * order: every one whose name matches `historical`, as the Python port
     * listed them (#1017), and those the presentation relates by a `segment`
     * relationship, wherever stored (#1013).
     */
    private async judgedParts(historical: RegExp, segment: string): Promise<Array<{ name: string; file: string }>> {
        const entries = packageEntries(this.unpackedDir);
        const names = [...entries.keys()];
        const judged = new Set(names.filter((name) => historical.test(name)));
        for (const part of await this.presentationRelatedParts(names, segment)) judged.add(part);
        return [...judged].sort((a, b) => (a < b ? -1 : a > b ? 1 : 0)).map((name) => ({ name, file: entries.get(name)! }));
    }

    /**
     * The stored parts the presentation relates by an internal relationship
     * of the `segment` type, matched exactly (`officeRelationshipTypeSegment`,
     * so `slideLayout` is not `slide`) (#1013). The presentation is the part
     * the package's `officeDocument` relationship names, the one PowerPoint
     * opens. Each target resolves against the folder of the part relating it,
     * a climb stopping at the package root, and names a stored part by OPC
     * part-name equivalence. Empty when the package names no stored
     * presentation.
     */
    private async presentationRelatedParts(names: readonly string[], segment: string): Promise<string[]> {
        const [presentation] = await this.relatedParts(names, "_rels/.rels", ".", "officeDocument");
        if (presentation === undefined) return [];
        const folder = path.posix.dirname(presentation);
        return this.relatedParts(names, path.posix.join(folder, "_rels", `${path.posix.basename(presentation)}.rels`), folder, segment);
    }

    /** The stored parts the relationships part `relsPart`, whose targets resolve against `folder`, names by internal `segment` relationships. */
    private async relatedParts(names: readonly string[], relsPart: string, folder: string, segment: string): Promise<string[]> {
        const relsName = findEquivalentPart(names, relsPart);
        if (relsName === undefined) return [];
        let dom: Document;
        try {
            dom = parseXmlBytes(await fs.readFile(path.join(this.unpackedDir, relsName)));
        } catch {
            // The XML and relationships checks report a part that cannot be read.
            return [];
        }
        const parts: string[] = [];
        for (const rel of this.relationshipElements(dom, relsName)) {
            if (officeRelationshipTypeSegment(rel.getAttribute("Type") ?? "") !== segment) continue;
            if (rel.getAttribute("TargetMode") === "External") continue;
            const resolved = relationshipTargetPartName(folder, rel.getAttribute("Target") ?? "");
            const part = resolved === undefined ? undefined : findEquivalentPart(names, resolved.replace(/^\//, ""));
            if (part !== undefined) parts.push(part);
        }
        return parts;
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
            (code >= 0x61 && code <= 0x7a); /* a-z */
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
            const child = children[i];
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
