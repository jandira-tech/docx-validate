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

import { promises as fs } from "node:fs";
import path from "node:path";

import { canonicalPartName, findEquivalentPart, relationshipTargetFolder, relationshipTargetPartName } from "../../../lib/content-types";
import type { ValidationIssue, ValidationResult } from "../../../lib/types";
import { mergeResults } from "../../../lib/types";
import { officeRelationshipTypeSegment, sidecarOf } from "../../../lib/word-read-parts";
import { getElementsByTagNameNSAll, parseXmlBytes } from "../../../lib/xml-helpers";
import type { BaseSchemaValidatorOptions } from "./base";
import { BaseSchemaValidator, OFFICE_RELATIONSHIPS_NAMESPACE, packageEntries, withoutRestatedIgnorablePrefixes } from "./base";
import { openDirectoryContext, packageExtractFailedIssue, storedNamesRun } from "../package-context";
import { registerStoryReferenceIdentity, relationshipAttributeName, withoutHandedOffCopies } from "../story-reference-identity";

export const PRESENTATIONML_NAMESPACE = "http://schemas.openxmlformats.org/presentationml/2006/main";
/** The ISO Strict PresentationML namespace, `PRESENTATIONML_NAMESPACE`'s Strict counterpart. */
const STRICT_PRESENTATIONML_NAMESPACE = "http://purl.oclc.org/ooxml/presentationml/main";

const PPTX_ELEMENT_RELATIONSHIP_TYPES: Record<string, string> = {
    sldid: "slide",
    sldmasterid: "slidemaster",
    notesmasterid: "notesmaster",
    sldlayoutid: "slidelayout",
    themeid: "theme",
    tablestyleid: "tablestyles",
};

/**
 * The real casing of each lower-case relationship-type name a PresentationML
 * element carrying an `r:id` expects: `rels-id-mismatch` compares these
 * exactly outside lenient, as the slide checks compare types (#1058). No
 * PresentationML element expects `theme` or `tableStyles` (#1061).
 */
const PPTX_RELATIONSHIP_SEGMENTS: Readonly<Record<string, string>> = {
    slide: "slide",
    slidemaster: "slideMaster",
    notesmaster: "notesMaster",
    handoutmaster: "handoutMaster",
    slidelayout: "slideLayout",
};

/** The finding a slide master's `sldLayoutId` copy is handed off to outside lenient (#1093). */
const SLIDE_LAYOUT_ID_PRODUCERS: ReadonlySet<string> = new Set(["sldlayout-missing-rid"]);

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
        // #1093: outside lenient a slide master's sldLayoutId that
        // sldlayout-missing-rid reported is one finding, that one, so the
        // relationship-id walk's copy of the same element and id is dropped,
        // as DOCX hands a story reference's copy to its Word rule (#616).
        // Lenient keeps both, as the Python validator reported them.
        const merged = mergeResults(...checks);
        return withoutRestatedIgnorablePrefixes(
            this.profile === "lenient" ? merged : withoutHandedOffCopies(merged, merged, SLIDE_LAYOUT_ID_PRODUCERS),
        );
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
        const slideMasters = await this.judgedParts(/^ppt\/slidemasters\/[^/]+\.xml$/, "slideMaster");

        if (slideMasters.length === 0) {
            return { valid: true, issues: [] };
        }

        for (const { file: slideMaster, sidecar: relsFile } of slideMasters) {
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

            // The master's relationships part, found by OPC part-name
            // equivalence as a slide's is (#1018), and reported missing under
            // its plain spelling.
            if (relsFile === undefined) {
                const plain = path.join(path.dirname(slideMaster), "_rels", `${path.basename(slideMaster)}.rels`);
                issues.push({
                    severity: "error",
                    message: `Missing relationships file: ${this.relPath(plain)}`,
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
                if (this.isSlideRelationship(rel, "slideLayout")) {
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
                    // #1109: outside lenient the attribute is named as the
                    // master spells it, as rels-id-missing does (#1090);
                    // lenient keeps the Python validator's literal r:id.
                    const attribute = this.profile === "lenient" ? "r:id" : relationshipAttributeName(sli, OFFICE_RELATIONSHIPS_NAMESPACE);
                    const issue: ValidationIssue = {
                        severity: "error",
                        message:
                            `sldLayoutId with id='${layoutId}' references ${attribute}='${rId}' ` +
                            `which is not found in slide layout relationships`,
                        path: this.relPath(slideMaster),
                        code: "sldlayout-missing-rid",
                    };
                    // #1093: the full run pairs the relationship-id walk's
                    // copy with this finding through the element identity,
                    // never the prose. Only the Transitional r:id is judged
                    // here, so a Strict rs:id beside it keeps its own copy.
                    registerStoryReferenceIdentity(issue, {
                        part: this.relPath(slideMaster),
                        namespaceURI: sli.namespaceURI ?? "",
                        localName: "sldLayoutId",
                        rid: rId,
                        ridNamespaceURI: OFFICE_RELATIONSHIPS_NAMESPACE,
                    });
                    issues.push(issue);
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
                if (this.isSlideRelationship(rel, "slideLayout")) layoutCount += 1;
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
                if (!this.isSlideRelationship(rel, "notesSlide")) continue;
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
                // Python: rels_file.stem (foo.xml.rels → foo.xml) then .replace(".xml", "") → foo.
                // Both strips read ASCII case-insensitively, as the patterns
                // that judge `foo.xml.RELS` match, and `\` reads as `/` (#1038).
                const stored = path.posix.basename(this.entryName(relsFile).replaceAll("\\", "/"));
                const slideName = stored.replace(/\.rels$/i, "").replace(/\.xml$/i, "");
                const entry = notesReferences.get(key) ?? { partName, references: [] };
                entry.references.push({ slideName, relsFile });
                notesReferences.set(key, entry);
            }
        }

        for (const { partName, references } of notesReferences.values()) {
            if (references.length <= 1) continue;
            // The count is of relationships, so one sidecar can give several
            // references: the message and the detail lines name each
            // relationships part once (#1045), and a part relating the notes
            // slide more than once carries its count (#1050).
            const counts = new Map<string, number>();
            for (const r of references) counts.set(r.relsFile, (counts.get(r.relsFile) ?? 0) + 1);
            const sidecars = [...new Map(references.map((r) => [r.relsFile, r])).values()];
            // Two equivalent spellings of one sidecar give one stem twice:
            // name each sidecar by its stored relationships part then (#1041).
            const stems = sidecars.map((r) => r.slideName);
            const names = new Set(stems).size === stems.length ? stems : sidecars.map((r) => this.relPath(r.relsFile));
            const counted = names.map((name, i) => {
                const count = counts.get(sidecars[i]?.relsFile ?? "") ?? 1;
                return count > 1 ? `${name} (${count} times)` : name;
            });
            const [only] = sidecars;
            issues.push({
                severity: "error",
                message:
                    sidecars.length === 1 && only
                        ? `Notes slide '${partName}' is referenced ${references.length} times by ${this.relPath(only.relsFile)}`
                        : `Notes slide '${partName}' is referenced by multiple slides: ${counted.join(", ")}`,
                code: "notes-shared",
            });
            for (const ref of sidecars) {
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
     * matches each stored name's canonical form, ASCII-lowercased with its
     * escaped unreserved characters decoded, as OPC compares part names, so
     * `ppt/slides/%5Frels/` is inside it (#1037); every stored spelling of two
     * equivalent names is judged, as the file-system walk before #1013 judged
     * them (#1017). A sidecar's folder is the parent of its `_rels` folder, as
     * its stored name spells it.
     */
    private async slideRelationshipsParts(): Promise<Array<{ file: string; folder: string }>> {
        const entries = packageEntries(this.unpackedDir);
        const names = [...entries.keys()];
        const judged = new Set(names.filter((name) => /^ppt\/slides\/_rels\/[^/]+\.xml\.rels$/.test(canonicalPartName(name))));
        for (const slide of await this.presentationRelatedParts(names, "slide")) {
            const name = sidecarOf(names, slide);
            if (name !== undefined) judged.add(name);
        }
        return [...judged]
            .sort((a, b) => (a < b ? -1 : a > b ? 1 : 0))
            .map((name) => ({ file: entries.get(name)!, folder: path.posix.dirname(path.posix.dirname(name)) }));
    }

    /**
     * The stored parts a check judges, each stored name once, in code-unit
     * order: every one whose canonical name (`canonicalPartName`, ASCII
     * lowercase) matches `historical`, as the Python port listed them (#1017,
     * #1037), and those the presentation relates by a `segment`
     * relationship, wherever stored (#1013). Each comes with the file of its
     * relationships part, found by OPC part-name equivalence (#1018), when
     * one is stored.
     */
    private async judgedParts(
        historical: RegExp,
        segment: string,
    ): Promise<Array<{ name: string; file: string; sidecar: string | undefined }>> {
        const entries = packageEntries(this.unpackedDir);
        const names = [...entries.keys()];
        const judged = new Set(names.filter((name) => historical.test(canonicalPartName(name))));
        for (const part of await this.presentationRelatedParts(names, segment)) judged.add(part);
        return [...judged]
            .sort((a, b) => (a < b ? -1 : a > b ? 1 : 0))
            .map((name) => {
                const sidecar = sidecarOf(names, name);
                return { name, file: entries.get(name)!, sidecar: sidecar === undefined ? undefined : entries.get(sidecar) };
            });
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

    /**
     * Outside lenient, a relationship satisfies a PresentationML element's
     * expected type only when its type is exactly one Office
     * relationship-type segment, Transitional or Strict, equal to the
     * expected one in its real casing (#1058). So a `p:sldId` naming a
     * `slideLayout` relationship no longer passes as `slide`, and a
     * `p:sldLayoutId` is judged as `sldlayout-missing-rid` judges it (#1051).
     * Every other element keeps the historical substring reading, as lenient
     * does: one in another namespace, such as a foreign element in a `p:ext`
     * that shares a PresentationML local name, and one whose expectation
     * `PPTX_RELATIONSHIP_SEGMENTS` does not name (#1061).
     */
    protected override relationshipTypeSatisfies(type: string, expected: string, elementNamespace: string): boolean {
        const known = Object.hasOwn(PPTX_RELATIONSHIP_SEGMENTS, expected) ? PPTX_RELATIONSHIP_SEGMENTS[expected] : undefined;
        if (
            this.profile === "lenient" ||
            known === undefined ||
            (elementNamespace !== PRESENTATIONML_NAMESPACE && elementNamespace !== STRICT_PRESENTATIONML_NAMESPACE)
        ) {
            return super.relationshipTypeSatisfies(type, expected, elementNamespace);
        }
        return officeRelationshipTypeSegment(type) === known;
    }

    /**
     * Whether the slide checks read `rel` as a `segment` relationship. Outside
     * lenient that is the relationship the slide walk reads (`relatedParts`,
     * #1013): exactly that Office type segment, Transitional or Strict, and
     * internal, as an External target names no part (ECMA-376 Part 2 §9.3).
     * Lenient keeps the historical reading, any type containing the segment,
     * External or not, as it keeps nested relationships (#434) (#1051).
     */
    private isSlideRelationship(rel: Element, segment: string): boolean {
        const type = rel.getAttribute("Type") ?? "";
        if (this.profile === "lenient") return type.includes(segment);
        return officeRelationshipTypeSegment(type) === segment && rel.getAttribute("TargetMode") !== "External";
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
