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
 * Base validator with common validation logic for OOXML document files.
 *
 * 1:1 TypeScript port of `src/docx-validate/scripts/office/validators/base.py`
 * (task #7). Substitutions versus the Python source:
 *
 *   - `defusedxml.minidom` / `lxml.etree` → `@xmldom/xmldom` + `xpath` via
 *     `lib/xml-helpers.ts`.
 *   - `lxml.etree.XMLSchema` → `libxmljs2`'s `parseXml` + `Document.validate`.
 *   - `pathlib.Path.rglob` → custom recursive walk in `walkFiles` since the
 *     Node `fs` module has no built-in recursive glob with patterns.
 *   - Each `validate_*` method that returned `bool` + side-effecting `print`
 *     in Python now returns a `ValidationResult` (`{ valid, issues }`) per
 *     the task spec. Diagnostic prints are gone — callers (the per-format
 *     subclasses in tasks #10/#11) decide what to do with the issues.
 *   - `lxml`'s `sourceline` is not exposed by `@xmldom`; line numbers are
 *     omitted from issues. Tests assert on file path + message instead.
 *
 * Subclasses (DOCX, PPTX) compose these primitives in their own
 * `validate()` and merge the per-check results with `mergeResults`.
 */

import { existsSync, promises as fs, readdirSync, readFileSync, statSync } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

import * as libxmljs from "libxmljs2";

import { withTempDir } from "../../../lib/run-cli";
import type { Profile, ValidationIssue, ValidationResult } from "../../../lib/types";
import { DEFAULT_PROFILE, OK_RESULT } from "../../../lib/types";
import { parseXml, serializeXml } from "../../../lib/xml-helpers";

/**
 * Schema mapping table — file basename / suffix → relative path inside the
 * `schemas/` directory. Mirrors `BaseSchemaValidator.SCHEMA_MAPPINGS`.
 */
const SCHEMA_MAPPINGS: Record<string, string> = {
    word: "ISO-IEC29500-4_2016/wml.xsd",
    ppt: "ISO-IEC29500-4_2016/pml.xsd",
    xl: "ISO-IEC29500-4_2016/sml.xsd",
    "[Content_Types].xml": "ecma/fouth-edition/opc-contentTypes.xsd",
    "app.xml": "ISO-IEC29500-4_2016/shared-documentPropertiesExtended.xsd",
    "core.xml": "ecma/fouth-edition/opc-coreProperties.xsd",
    "custom.xml": "ISO-IEC29500-4_2016/shared-documentPropertiesCustom.xsd",
    ".rels": "ecma/fouth-edition/opc-relationships.xsd",
    "people.xml": "microsoft/wml-2012.xsd",
    "commentsIds.xml": "microsoft/wml-cid-2016.xsd",
    "commentsExtensible.xml": "microsoft/wml-cex-2018.xsd",
    "commentsExtended.xml": "microsoft/wml-2012.xsd",
    chart: "ISO-IEC29500-4_2016/dml-chart.xsd",
    theme: "ISO-IEC29500-4_2016/dml-main.xsd",
    drawing: "ISO-IEC29500-4_2016/dml-main.xsd",
    "bibliography.xml": "ISO-IEC29500-4_2016/shared-bibliography.xsd",
    "additionalCharacteristics.xml": "ISO-IEC29500-4_2016/shared-additionalCharacteristics.xsd",
    itemProps: "ISO-IEC29500-4_2016/shared-customXmlDataProperties.xsd",
    sig: "ecma/fouth-edition/opc-digSig.xsd",
};

const IGNORED_VALIDATION_ERRORS: readonly string[] = [
    "hyphenationZone",
    "purl.org/dc/terms",
    // libxmljs2 swallows the underlying schema-load failure (e.g. unresolved
    // <xs:import namespace="purl.org/dc/terms"/> in opc-coreProperties.xsd)
    // into a single "Invalid XSD schema" string. lxml emits the actual import
    // error which the entry above already filters, so this keeps parity with
    // the Python validator's behaviour for the same broken-XSD-import case.
    "Invalid XSD schema",
];

const UNIQUE_ID_REQUIREMENTS: Record<string, [attr: string, scope: "file" | "global"]> = {
    comment: ["id", "file"],
    commentrangestart: ["id", "file"],
    commentrangeend: ["id", "file"],
    bookmarkstart: ["id", "file"],
    bookmarkend: ["id", "file"],
    sldid: ["id", "file"],
    sldmasterid: ["id", "global"],
    sldlayoutid: ["id", "global"],
    cm: ["authorid", "file"],
    sheet: ["sheetid", "file"],
    definedname: ["id", "file"],
    cxnsp: ["id", "file"],
    sp: ["id", "file"],
    pic: ["id", "file"],
    grpsp: ["id", "file"],
};

const EXCLUDED_ID_CONTAINERS = new Set<string>(["sectionlst"]);

const MC_NAMESPACE = "http://schemas.openxmlformats.org/markup-compatibility/2006";
export const XML_NAMESPACE = "http://www.w3.org/XML/1998/namespace";
export const PACKAGE_RELATIONSHIPS_NAMESPACE = "http://schemas.openxmlformats.org/package/2006/relationships";
export const OFFICE_RELATIONSHIPS_NAMESPACE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const CONTENT_TYPES_NAMESPACE = "http://schemas.openxmlformats.org/package/2006/content-types";

const MAIN_CONTENT_FOLDERS = new Set(["word", "ppt", "xl"]);

const OOXML_NAMESPACES = new Set<string>([
    "http://schemas.openxmlformats.org/officeDocument/2006/math",
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "http://schemas.openxmlformats.org/schemaLibrary/2006/main",
    "http://schemas.openxmlformats.org/drawingml/2006/main",
    "http://schemas.openxmlformats.org/drawingml/2006/chart",
    "http://schemas.openxmlformats.org/drawingml/2006/chartDrawing",
    "http://schemas.openxmlformats.org/drawingml/2006/diagram",
    "http://schemas.openxmlformats.org/drawingml/2006/picture",
    "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing",
    "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
    "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "http://schemas.openxmlformats.org/presentationml/2006/main",
    "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "http://schemas.openxmlformats.org/officeDocument/2006/sharedTypes",
    "http://www.w3.org/XML/1998/namespace",
]);

/**
 * URI prefix shared by all ISO OOXML Strict conformance class namespaces.
 * Strict documents use `purl.oclc.org/ooxml/` instead of the Transitional
 * `schemas.openxmlformats.org/…` prefix. Any root element namespace that
 * starts with this prefix belongs to the Strict conformance class.
 *
 * The XSD schemas bundled with this validator are Transitional-only, so Strict
 * files cannot be validated against them — `validateAgainstXsd` detects these
 * via `_isStrictXmlFile` and emits an info-level `xsd-strict-skipped` issue
 * instead of false-positive xsd-* errors.
 *
 * See PORT_NOTES.md § "ISO OOXML Strict XSD validation" for rationale.
 * Adding full Strict schema support is tracked as a future improvement.
 */
const STRICT_OOXML_NS_PREFIX = "http://purl.oclc.org/ooxml/";

// Use a non-`g` regex for the predicate (`.test()` against a stateful `g`
// pattern advances `lastIndex` between calls and silently mis-skips matches),
// and build a fresh `g`-flagged regex inside the replace path so the two are
// independent. Keep the literal in one place to keep them in sync.
const TEMPLATE_TAG_PATTERN_SOURCE = String.raw`\{\{[^}]*\}\}`;
const TEMPLATE_TAG_PATTERN = new RegExp(TEMPLATE_TAG_PATTERN_SOURCE);
const TEMPLATE_TAG_PATTERN_GLOBAL = new RegExp(TEMPLATE_TAG_PATTERN_SOURCE, "g");

export interface BaseSchemaValidatorOptions {
    /** Path to the unpacked OOXML directory. */
    unpackedDir: string;
    /** Path to the original .docx/.pptx for "ignore pre-existing errors" diffing. */
    originalFile?: string;
    /**
     * Where to find the XSD files. Defaults to the bundled schemas under
     * `scripts/office/schemas/` — pass an explicit path if you have copied
     * them somewhere else.
     */
    schemasDir?: string;
    verbose?: boolean;
    /** Validation profile. Defaults to `"lenient"`. See {@link Profile}. */
    profile?: Profile;
}

/**
 * Default XSD root — points at the schemas directory shipped alongside
 * the validator source under `scripts/office/schemas/`.
 *
 * Two layouts are supported because `vp pack` collapses the source tree
 * to `dist/index.mjs`:
 *   - dev / source layout: this file sits at `src/scripts/office/validators/`,
 *     so `../schemas` resolves to `src/scripts/office/schemas/`.
 *   - published / dist layout: bundled file sits at `dist/index.mjs`, and
 *     `vite.config.ts` copies the schemas to `dist/schemas/`, so
 *     `./schemas` (relative to the bundle) is the right path.
 *
 * We probe in order and return whichever exists; if neither does we fall
 * back to the source path so the error message points somewhere sensible.
 */
export function defaultSchemasDir(): string {
    const here = path.dirname(fileURLToPath(import.meta.url));
    const candidates = [path.resolve(here, "..", "schemas"), path.resolve(here, "schemas")];
    for (const candidate of candidates) {
        if (existsSync(candidate)) {
            return candidate;
        }
    }
    return candidates[0];
}

interface IdHit {
    file: string;
    tag: string;
}

export interface XsdValidationOutcome {
    /** `null` mirrors Python's `(None, None)` skip — no schema configured. */
    valid: boolean | null;
    errors: Set<string>;
}

export class BaseSchemaValidator {
    readonly unpackedDir: string;
    readonly originalFile: string | null;
    readonly verbose: boolean;
    readonly schemasDir: string;
    readonly xmlFiles: string[];
    readonly profile: Profile;

    /**
     * Subclasses override to map an element name → expected relationship type
     * (used by `validateAllRelationshipIds`). Mirrors
     * `BaseSchemaValidator.ELEMENT_RELATIONSHIP_TYPES`.
     */
    protected readonly elementRelationshipTypes: Record<string, string> = {};

    constructor(opts: BaseSchemaValidatorOptions) {
        this.unpackedDir = path.resolve(opts.unpackedDir);
        this.originalFile = opts.originalFile ? path.resolve(opts.originalFile) : null;
        this.verbose = opts.verbose ?? false;
        this.schemasDir = opts.schemasDir ?? defaultSchemasDir();
        this.xmlFiles = walkFiles(this.unpackedDir, [".xml", ".rels"]);
        this.profile = opts.profile ?? DEFAULT_PROFILE;
    }

    /**
     * Strict-profile-only check: every XML part must NOT begin with a UTF-8
     * BOM (U+FEFF, encoded as bytes EF BB BF). The XML spec permits the BOM
     * and Microsoft Office routinely emits it; the lenient profile silently
     * strips it before parsing. Strict callers want to know it's there.
     *
     * Severity is profile-aware:
     *   - `strict`  → reported as `error` (fails the document).
     *   - `lenient` → reported as `warning` (visible to anyone inspecting
     *     `result.issues` but does NOT flip `valid` to false).
     *
     * The check NEVER becomes a silent no-op — even in lenient mode the
     * gap is surfaced as a warning so callers can audit it later.
     */
    async validateNoBom(): Promise<ValidationResult> {
        const issues: ValidationIssue[] = [];
        const severity: "error" | "warning" = this.profile === "strict" ? "error" : "warning";
        for (const xmlFile of this.xmlFiles) {
            let bytes: Buffer;
            try {
                bytes = await fs.readFile(xmlFile);
            } catch {
                continue;
            }
            if (bytes.length >= 3 && bytes[0] === 0xef && bytes[1] === 0xbb && bytes[2] === 0xbf) {
                issues.push({
                    severity,
                    message:
                        "XML part begins with a UTF-8 BOM (U+FEFF). Permitted by the XML spec but " +
                        "non-canonical for OOXML; reported as error under strict, warning under lenient.",
                    path: this.relPath(xmlFile),
                    code: "xml-bom-leading",
                });
            }
        }
        return finalize(issues);
    }

    /**
     * Profile-aware check: every `_rels/*.rels` part must contain at
     * least one `<Relationship>` child. Per OPC §9.3 / ISO 29500-2, a
     * relationship part exists *because* its source part has outgoing
     * relationships — an empty `<Relationships/>` is non-canonical and
     * Word strips them on save.
     *
     * Severity:
     *   - `strict`  → `error` (spec-purist: flag the non-canonical state).
     *   - `lenient` → `warning` (real-world unpack→pack pipelines and
     *     third-party writers routinely emit empty rels parts; visible
     *     but doesn't fail validation).
     *
     * This is the canonical example of a "strict-only" gap surfaced by
     * the comparison against Word's own save output: jubarte / our pack
     * helper preserves empty rels sidecars from the input, while Word
     * strips them. The repaired-file regression of the
     * `sample-document.id-overflow.docx` fixture surfaces here.
     */
    async validateNoEmptyRelsParts(): Promise<ValidationResult> {
        const issues: ValidationIssue[] = [];
        const severity: "error" | "warning" = this.profile === "strict" ? "error" : "warning";
        for (const xmlFile of this.xmlFiles) {
            if (!xmlFile.endsWith(".rels")) continue;
            let dom: Document;
            try {
                dom = parseXml(await fs.readFile(xmlFile, "utf-8"));
            } catch {
                // Malformed rels files surface elsewhere via xml-syntax /
                // rels-broken codes; don't double-report here.
                continue;
            }
            const root = dom.documentElement;
            if (!root) continue;
            const rels = root.getElementsByTagNameNS(
                "http://schemas.openxmlformats.org/package/2006/relationships",
                "Relationship",
            );
            if (rels.length === 0) {
                // Additional check: only report as empty if the root also has no
                // element children. Malformed rels files with Relationship elements
                // in the wrong namespace will return zero from getElementsByTagNameNS
                // but should NOT be classified as empty parts (they have element
                // children, just in the wrong namespace).
                const hasElementChildren = hasAnyElementChild(root);
                if (!hasElementChildren) {
                    issues.push({
                        severity,
                        message:
                            `Relationships part has zero <Relationship> children. ` +
                            `OPC §9.3 says a rels part should exist only when its source part has outgoing relationships; ` +
                            `Word strips empty rels sidecars on save.`,
                        path: this.relPath(xmlFile),
                        code: "rels-empty-part",
                    });
                }
            }
        }
        return finalize(issues);
    }

    /** Subclass entry point. Returns merged result of all checks. */
    validate(): Promise<ValidationResult> {
        throw new Error("Subclasses must implement validate()");
    }

    // ----- repair --------------------------------------------------------------

    async repair(): Promise<number> {
        const wsRepairs = await this.repairWhitespacePreservation();
        const relsRepairs = await this.repairEmptyRelsParts();
        return wsRepairs + relsRepairs;
    }

    /**
     * Delete `_rels/*.rels` parts whose `<Relationships>` element has zero
     * children. Mirrors what Word does on save — see
     * `validateNoEmptyRelsParts` for the spec citation.
     *
     * Returns the count of files deleted. The deleted files are also
     * removed from `this.xmlFiles` so subsequent passes (and validation
     * re-runs in the same instance) don't see them.
     */
    async repairEmptyRelsParts(): Promise<number> {
        let repairs = 0;
        const surviving: string[] = [];
        for (const xmlFile of this.xmlFiles) {
            if (!xmlFile.endsWith(".rels")) {
                surviving.push(xmlFile);
                continue;
            }
            let isEmpty = false;
            try {
                const dom = parseXml(await fs.readFile(xmlFile, "utf-8"));
                const root = dom.documentElement;
                if (root) {
                    const rels = root.getElementsByTagNameNS(
                        "http://schemas.openxmlformats.org/package/2006/relationships",
                        "Relationship",
                    );
                    if (rels.length === 0) {
                        // Additional check: only mark as empty if the root also has no
                        // element children. Malformed rels files with Relationship elements
                        // in the wrong namespace should not be deleted.
                        const hasElementChildren = hasAnyElementChild(root);
                        isEmpty = !hasElementChildren;
                    }
                }
            } catch {
                // malformed — leave it alone, surfaced elsewhere
            }
            if (isEmpty) {
                try {
                    await fs.rm(xmlFile);
                    repairs += 1;
                    // intentionally not added to `surviving`
                } catch {
                    // best-effort
                    surviving.push(xmlFile);
                }
            } else {
                surviving.push(xmlFile);
            }
        }
        // Mutate the cached list in place so subsequent passes within
        // the same validator instance don't try to read the deleted files.
        this.xmlFiles.length = 0;
        for (const f of surviving) this.xmlFiles.push(f);
        return repairs;
    }

    async repairWhitespacePreservation(): Promise<number> {
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
                    if (!elem.tagName.endsWith(":t")) continue;
                    const first = elem.firstChild;
                    if (!first) continue;
                    const text = first.nodeValue ?? "";
                    if (!text) continue;
                    const startsWS = text.startsWith(" ") || text.startsWith("\t");
                    const endsWS = text.endsWith(" ") || text.endsWith("\t");
                    if (!startsWS && !endsWS) continue;
                    if (elem.getAttribute("xml:space") !== "preserve") {
                        elem.setAttribute("xml:space", "preserve");
                        repairs += 1;
                        modified = true;
                        if (this.verbose) {
                            const preview = text.length > 30 ? JSON.stringify(text.slice(0, 30)) + "..." : JSON.stringify(text);
                            process.stdout.write(
                                `  Repaired: ${path.basename(xmlFile)}: Added xml:space='preserve' to ${elem.tagName}: ${preview}\n`,
                            );
                        }
                    }
                }

                if (modified) {
                    await fs.writeFile(xmlFile, serializeXml(dom, "UTF-8"), "utf-8");
                }
            } catch {
                // pass — mirrors Python's bare except
            }
        }

        return repairs;
    }

    // ----- well-formedness & namespaces ---------------------------------------

    async validateXml(): Promise<ValidationResult> {
        const issues: ValidationIssue[] = [];
        for (const xmlFile of this.xmlFiles) {
            try {
                const content = await fs.readFile(xmlFile, "utf-8");
                parseXml(content);
            } catch (err) {
                const message = err instanceof Error ? err.message : String(err);
                issues.push({
                    severity: "error",
                    message,
                    path: this.relPath(xmlFile),
                    code: "xml-syntax",
                });
            }
        }
        return finalize(issues);
    }

    async validateNamespaces(): Promise<ValidationResult> {
        const issues: ValidationIssue[] = [];
        for (const xmlFile of this.xmlFiles) {
            let dom: Document;
            try {
                dom = parseXml(await fs.readFile(xmlFile, "utf-8"));
            } catch {
                continue;
            }
            const root = dom.documentElement;
            if (!root) continue;

            const declared = collectDeclaredPrefixes(root);

            const attrs = root.attributes;
            for (let i = 0; i < attrs.length; i += 1) {
                const attr = attrs.item(i);
                if (!attr) continue;
                if (!attr.name.endsWith("Ignorable")) continue;
                const tokens = attr.value.split(/\s+/).filter(Boolean);
                for (const ns of tokens) {
                    if (!declared.has(ns)) {
                        issues.push({
                            severity: "error",
                            message: `Namespace '${ns}' in Ignorable but not declared`,
                            path: this.relPath(xmlFile),
                            code: "ignorable-undeclared",
                        });
                    }
                }
            }
        }
        return finalize(issues);
    }

    // ----- unique IDs ---------------------------------------------------------

    async validateUniqueIds(): Promise<ValidationResult> {
        const issues: ValidationIssue[] = [];
        const globalIds = new Map<string, IdHit>();

        for (const xmlFile of this.xmlFiles) {
            let dom: Document;
            try {
                dom = parseXml(await fs.readFile(xmlFile, "utf-8"));
            } catch (err) {
                issues.push({
                    severity: "error",
                    message: `Error: ${err instanceof Error ? err.message : String(err)}`,
                    path: this.relPath(xmlFile),
                    code: "id-parse",
                });
                continue;
            }

            // Python: drop AlternateContent subtrees first.
            removeMcAlternateContent(dom);

            const fileIds = new Map<string, Map<string, true>>();
            const root = dom.documentElement;
            if (!root) continue;

            for (const elem of iterElements(root)) {
                const tag = localName(elem.tagName ?? elem.nodeName).toLowerCase();
                const requirement = UNIQUE_ID_REQUIREMENTS[tag];
                if (!requirement) continue;

                if (hasExcludedAncestor(elem)) continue;

                const [attrName, scope] = requirement;
                const idValue = findIdAttribute(elem, attrName);
                if (idValue === null) continue;

                if (scope === "global") {
                    const prev = globalIds.get(idValue);
                    if (prev) {
                        issues.push({
                            severity: "error",
                            message: `Global ID '${idValue}' in <${tag}> already used in ${prev.file} in <${prev.tag}>`,
                            path: this.relPath(xmlFile),
                            code: "id-duplicate-global",
                        });
                    } else {
                        globalIds.set(idValue, { file: this.relPath(xmlFile), tag });
                    }
                } else {
                    const key = `${tag}|${attrName}`;
                    let bucket = fileIds.get(key);
                    if (!bucket) {
                        bucket = new Map();
                        fileIds.set(key, bucket);
                    }
                    if (bucket.has(idValue)) {
                        issues.push({
                            severity: "error",
                            message: `Duplicate ${attrName}='${idValue}' in <${tag}>`,
                            path: this.relPath(xmlFile),
                            code: "id-duplicate-file",
                        });
                    } else {
                        bucket.set(idValue, true);
                    }
                }
            }
        }

        return finalize(issues);
    }

    // ----- file references / .rels --------------------------------------------

    async validateFileReferences(): Promise<ValidationResult> {
        const issues: ValidationIssue[] = [];
        const relsFiles = walkFiles(this.unpackedDir, [".rels"]);

        if (relsFiles.length === 0) {
            return OK_RESULT;
        }

        const allFiles: string[] = [];
        for (const f of walkFiles(this.unpackedDir)) {
            const base = path.basename(f);
            if (base === "[Content_Types].xml" || f.endsWith(".rels")) continue;
            allFiles.push(path.resolve(f));
        }

        const allReferenced = new Set<string>();

        for (const relsFile of relsFiles) {
            let dom: Document;
            try {
                dom = parseXml(await fs.readFile(relsFile, "utf-8"));
            } catch (err) {
                issues.push({
                    severity: "error",
                    message: `Error parsing ${this.relPath(relsFile)}: ${err instanceof Error ? err.message : String(err)}`,
                    path: this.relPath(relsFile),
                    code: "rels-parse",
                });
                continue;
            }
            const root = dom.documentElement;
            if (!root) continue;

            const relsDir = path.dirname(relsFile);
            const referenced: string[] = [];
            const broken: string[] = [];

            const rels = dom.getElementsByTagNameNS(PACKAGE_RELATIONSHIPS_NAMESPACE, "Relationship");
            for (let i = 0; i < rels.length; i += 1) {
                const rel = rels.item(i);
                if (!rel) continue;
                const target = rel.getAttribute("Target");
                if (!target) continue;
                if (target.startsWith("http") || target.startsWith("mailto:")) continue;

                let targetPath: string;
                if (target.startsWith("/")) {
                    targetPath = path.join(this.unpackedDir, target.replace(/^\/+/, ""));
                } else if (path.basename(relsFile) === ".rels") {
                    targetPath = path.join(this.unpackedDir, target);
                } else {
                    targetPath = path.join(path.dirname(relsDir), target);
                }

                try {
                    targetPath = path.resolve(targetPath);
                    if (existsSync(targetPath) && statSync(targetPath).isFile()) {
                        referenced.push(targetPath);
                        allReferenced.add(targetPath);
                    } else {
                        broken.push(target);
                    }
                } catch {
                    broken.push(target);
                }
            }

            for (const ref of broken) {
                issues.push({
                    severity: "error",
                    message: `Broken reference to ${ref}`,
                    path: this.relPath(relsFile),
                    code: "rels-broken",
                });
            }
        }

        const unreferenced = allFiles.filter((f) => !allReferenced.has(f)).sort();
        for (const unref of unreferenced) {
            issues.push({
                severity: "error",
                message: `Unreferenced file: ${this.relPath(unref)}`,
                path: this.relPath(unref),
                code: "rels-unreferenced",
            });
        }

        return finalize(issues);
    }

    async validateRelationshipElements(): Promise<ValidationResult> {
        const issues: ValidationIssue[] = [];
        const relsFiles = walkFiles(this.unpackedDir, [".rels"]);

        for (const relsFile of relsFiles) {
            let dom: Document;
            try {
                dom = parseXml(await fs.readFile(relsFile, "utf-8"));
            } catch {
                continue;
            }

            const rels = dom.getElementsByTagNameNS(PACKAGE_RELATIONSHIPS_NAMESPACE, "Relationship");
            for (let i = 0; i < rels.length; i += 1) {
                const rel = rels.item(i);
                if (!rel) continue;

                const rid = rel.getAttribute("Id");
                const type = rel.getAttribute("Type");
                const target = rel.getAttribute("Target");

                if (!rid || !type || !target) {
                    issues.push({
                        severity: "error",
                        message:
                            `<Relationship> missing required attribute(s): ` +
                            [!rid && "Id", !type && "Type", !target && "Target"].filter(Boolean).join(", "),
                        path: this.relPath(relsFile),
                        code: "rels-empty-element",
                    });
                    continue;
                }

                // OPC spec requires <Relationship> to be an empty element (no body content).
                // Flag any non-empty body: text (whitespace-only OR otherwise) and/or child elements.
                let hasBodyContent = false;
                for (let c = rel.firstChild; c; c = c.nextSibling) {
                    if (c.nodeType === 3 /* TEXT_NODE */) {
                        const text = (c as Text).data ?? (c as Text).nodeValue ?? "";
                        if (text.length > 0) {
                            hasBodyContent = true;
                            break;
                        }
                    } else if (c.nodeType === 1 /* ELEMENT_NODE */) {
                        hasBodyContent = true;
                        break;
                    }
                }
                if (hasBodyContent) {
                    issues.push({
                        severity: "error",
                        message: `<Relationship Id="${rid}"> must be a self-closing empty element, not <Relationship>...</Relationship>`,
                        path: this.relPath(relsFile),
                        code: "rels-empty-element",
                    });
                }
            }
        }

        return finalize(issues);
    }

    async validateAllRelationshipIds(): Promise<ValidationResult> {
        const issues: ValidationIssue[] = [];

        for (const xmlFile of this.xmlFiles) {
            if (xmlFile.endsWith(".rels")) continue;

            const relsDir = path.join(path.dirname(xmlFile), "_rels");
            const relsFile = path.join(relsDir, `${path.basename(xmlFile)}.rels`);

            // Gap 4: flag XML parts that use r:id attributes but have no .rels sidecar.
            if (!existsSync(relsFile)) {
                let xmlDomCheck: Document;
                try {
                    xmlDomCheck = parseXml(await fs.readFile(xmlFile, "utf-8"));
                } catch {
                    continue;
                }
                const rNs = OFFICE_RELATIONSHIPS_NAMESPACE;
                const xmlRootCheck = xmlDomCheck.documentElement;
                if (!xmlRootCheck) continue;
                let hasRidRef = false;
                outer: for (const elem of iterElements(xmlRootCheck)) {
                    for (const attrName of ["id", "embed", "link"] as const) {
                        if (elem.getAttributeNS(rNs, attrName)) {
                            hasRidRef = true;
                            break outer;
                        }
                    }
                }
                if (hasRidRef) {
                    issues.push({
                        severity: "error",
                        message: `Part '${this.relPath(xmlFile)}' uses r:id relationship references but has no .rels sidecar file`,
                        path: this.relPath(xmlFile),
                        code: "rels-missing-sidecar",
                    });
                }
                continue;
            }

            let relsDom: Document;
            try {
                relsDom = parseXml(await fs.readFile(relsFile, "utf-8"));
            } catch (err) {
                issues.push({
                    severity: "error",
                    message: `Error processing ${this.relPath(xmlFile)}: ${err instanceof Error ? err.message : String(err)}`,
                    path: this.relPath(xmlFile),
                    code: "rels-id-parse",
                });
                continue;
            }

            const ridToType = new Map<string, string>();
            const rels = relsDom.getElementsByTagNameNS(PACKAGE_RELATIONSHIPS_NAMESPACE, "Relationship");
            for (let i = 0; i < rels.length; i += 1) {
                const rel = rels.item(i);
                if (!rel) continue;
                const rid = rel.getAttribute("Id");
                const relType = rel.getAttribute("Type") ?? "";
                if (!rid) continue;
                if (ridToType.has(rid)) {
                    issues.push({
                        severity: "error",
                        message: `Duplicate relationship ID '${rid}' (IDs must be unique)`,
                        path: this.relPath(relsFile),
                        code: "rels-id-duplicate",
                    });
                }
                const typeName = relType.includes("/") ? (relType.split("/").pop() ?? relType) : relType;
                ridToType.set(rid, typeName);
            }

            let xmlDom: Document;
            try {
                xmlDom = parseXml(await fs.readFile(xmlFile, "utf-8"));
            } catch (err) {
                issues.push({
                    severity: "error",
                    message: `Error processing ${this.relPath(xmlFile)}: ${err instanceof Error ? err.message : String(err)}`,
                    path: this.relPath(xmlFile),
                    code: "rels-id-parse",
                });
                continue;
            }

            const rNs = OFFICE_RELATIONSHIPS_NAMESPACE;
            const ridAttrsToCheck: ReadonlyArray<"id" | "embed" | "link"> = ["id", "embed", "link"];
            const xmlRoot = xmlDom.documentElement;
            if (!xmlRoot) continue;

            for (const elem of iterElements(xmlRoot)) {
                for (const attrName of ridAttrsToCheck) {
                    const ridAttr = elem.getAttributeNS(rNs, attrName);
                    if (!ridAttr) continue;
                    const elemName = localName(elem.tagName ?? elem.nodeName);

                    if (!ridToType.has(ridAttr)) {
                        const valid = [...ridToType.keys()].sort().slice(0, 5).join(", ");
                        const ellipsis = ridToType.size > 5 ? "..." : "";
                        issues.push({
                            severity: "error",
                            message:
                                `<${elemName}> r:${attrName} references non-existent relationship '${ridAttr}' ` +
                                `(valid IDs: ${valid}${ellipsis})`,
                            path: this.relPath(xmlFile),
                            code: "rels-id-missing",
                        });
                    } else if (attrName === "id" && Object.keys(this.elementRelationshipTypes).length > 0) {
                        const expectedType = this._getExpectedRelationshipType(elemName);
                        if (expectedType) {
                            const actualType = ridToType.get(ridAttr) ?? "";
                            if (!actualType.toLowerCase().includes(expectedType)) {
                                issues.push({
                                    severity: "error",
                                    message:
                                        `<${elemName}> references '${ridAttr}' which points to '${actualType}' ` +
                                        `but should point to a '${expectedType}' relationship`,
                                    path: this.relPath(xmlFile),
                                    code: "rels-id-mismatch",
                                });
                            }
                        }
                    }
                }
            }
        }

        return finalize(issues);
    }

    protected _getExpectedRelationshipType(elementName: string): string | null {
        const elemLower = elementName.toLowerCase();
        if (this.elementRelationshipTypes[elemLower]) {
            return this.elementRelationshipTypes[elemLower];
        }
        if (elemLower.endsWith("id") && elemLower.length > 2) {
            const prefix = elemLower.slice(0, -2);
            if (prefix.endsWith("master") || prefix.endsWith("layout")) return prefix;
            if (prefix === "sld") return "slide";
            return prefix;
        }
        if (elemLower.endsWith("reference") && elemLower.length > 9) {
            return elemLower.slice(0, -9);
        }
        return null;
    }

    // ----- content types ------------------------------------------------------

    async validateContentTypes(): Promise<ValidationResult> {
        const issues: ValidationIssue[] = [];
        const contentTypesFile = path.join(this.unpackedDir, "[Content_Types].xml");
        if (!existsSync(contentTypesFile)) {
            return {
                valid: false,
                issues: [
                    {
                        severity: "error",
                        message: "[Content_Types].xml file not found",
                        code: "ct-missing",
                    },
                ],
            };
        }

        let dom: Document;
        try {
            dom = parseXml(await fs.readFile(contentTypesFile, "utf-8"));
        } catch (err) {
            return {
                valid: false,
                issues: [
                    {
                        severity: "error",
                        message: `Error parsing [Content_Types].xml: ${err instanceof Error ? err.message : String(err)}`,
                        path: "[Content_Types].xml",
                        code: "ct-parse",
                    },
                ],
            };
        }

        const declaredParts = new Set<string>();
        const declaredExtensions = new Set<string>();

        const overrides = dom.getElementsByTagNameNS(CONTENT_TYPES_NAMESPACE, "Override");
        for (let i = 0; i < overrides.length; i += 1) {
            const o = overrides.item(i);
            if (!o) continue;
            const partName = o.getAttribute("PartName");
            if (partName) declaredParts.add(partName.replace(/^\/+/, ""));
        }

        const defaults = dom.getElementsByTagNameNS(CONTENT_TYPES_NAMESPACE, "Default");
        for (let i = 0; i < defaults.length; i += 1) {
            const d = defaults.item(i);
            if (!d) continue;
            const ext = d.getAttribute("Extension");
            if (ext) declaredExtensions.add(ext.toLowerCase());
        }

        const declarableRoots = new Set<string>([
            "sld",
            "sldLayout",
            "sldMaster",
            "presentation",
            "document",
            "workbook",
            "worksheet",
            "theme",
        ]);

        const mediaExtensions: Record<string, string> = {
            png: "image/png",
            jpg: "image/jpeg",
            jpeg: "image/jpeg",
            gif: "image/gif",
            bmp: "image/bmp",
            tiff: "image/tiff",
            wmf: "image/x-wmf",
            emf: "image/x-emf",
        };

        const allFiles = walkFiles(this.unpackedDir);

        for (const xmlFile of this.xmlFiles) {
            const pathStr = this.relPath(xmlFile).replace(/\\/g, "/");
            if ([".rels", "[Content_Types]", "docProps/", "_rels/"].some((s) => pathStr.includes(s))) {
                continue;
            }

            try {
                const root = parseXml(await fs.readFile(xmlFile, "utf-8")).documentElement;
                if (!root) continue;
                const rootName = localName(root.tagName ?? root.nodeName);
                if (declarableRoots.has(rootName) && !declaredParts.has(pathStr)) {
                    issues.push({
                        severity: "error",
                        message: `File with <${rootName}> root not declared in [Content_Types].xml`,
                        path: pathStr,
                        code: "ct-undeclared-part",
                    });
                }
            } catch {}
        }

        for (const filePath of allFiles) {
            const ext = path.extname(filePath).toLowerCase();
            if (ext === ".xml" || ext === ".rels") continue;
            if (path.basename(filePath) === "[Content_Types].xml") continue;
            const parts = filePath.split(path.sep);
            if (parts.includes("_rels") || parts.includes("docProps")) continue;

            const extension = ext.replace(/^\./, "");
            if (extension && !declaredExtensions.has(extension) && extension in mediaExtensions) {
                const rel = this.relPath(filePath);
                issues.push({
                    severity: "error",
                    message:
                        `File with extension '${extension}' not declared in [Content_Types].xml - ` +
                        `should add: <Default Extension="${extension}" ContentType="${mediaExtensions[extension]}"/>`,
                    path: rel,
                    code: "ct-undeclared-ext",
                });
            }
        }

        return finalize(issues);
    }

    // ----- XSD validation -----------------------------------------------------

    async validateAgainstXsd(): Promise<ValidationResult> {
        const issues: ValidationIssue[] = [];
        let strictSkipped = false;

        for (const xmlFile of this.xmlFiles) {
            if (this._isStrictXmlFile(xmlFile)) {
                if (!strictSkipped) {
                    strictSkipped = true;
                    issues.push({
                        severity: "info",
                        message:
                            "Document uses ISO OOXML Strict conformance class namespace URIs; " +
                            "XSD validation skipped (only Transitional schemas are bundled). " +
                            "See PORT_NOTES.md § 'ISO OOXML Strict XSD validation'.",
                        code: "xsd-strict-skipped",
                    });
                }
                continue;
            }

            const outcome = await this.validateFileAgainstXsd(xmlFile);
            if (outcome.valid === null) continue;
            if (outcome.valid) continue;
            const rel = this.relPath(xmlFile);
            issues.push({
                severity: "error",
                message: `${rel}: ${outcome.errors.size} new error(s)`,
                path: rel,
                code: "xsd-summary",
            });
            const list = [...outcome.errors].slice(0, 3);
            for (const e of list) {
                const truncated = e.length > 250 ? `${e.slice(0, 250)}...` : e;
                issues.push({
                    severity: "error",
                    message: `  - ${truncated}`,
                    path: rel,
                    code: "xsd-error",
                });
            }
        }

        return finalize(issues);
    }

    /**
     * Returns true when `xmlFile`'s root element namespace URI belongs to the
     * ISO OOXML Strict conformance class. All Strict URIs share the prefix
     * `http://purl.oclc.org/ooxml/`. Used by `validateAgainstXsd` to skip
     * XSD checks that only apply to Transitional documents.
     */
    protected _isStrictXmlFile(xmlFile: string): boolean {
        try {
            const content = readFileSync(xmlFile, "utf-8");
            const dom = parseXml(content);
            const root = dom.documentElement;
            if (!root) return false;
            const ns = root.namespaceURI;
            return !!ns && ns.startsWith(STRICT_OOXML_NS_PREFIX);
        } catch {
            return false;
        }
    }

    async validateFileAgainstXsd(xmlFile: string): Promise<XsdValidationOutcome> {
        const single = this._validateSingleFileXsd(xmlFile);
        if (single.valid === null) {
            return { valid: null, errors: new Set() };
        }
        if (single.valid) {
            return { valid: true, errors: new Set() };
        }

        const originalErrors = await this._getOriginalFileErrors(xmlFile);
        const newErrors = new Set<string>();
        for (const e of single.errors) {
            if (originalErrors.has(e)) continue;
            if (IGNORED_VALIDATION_ERRORS.some((p) => e.includes(p))) continue;
            newErrors.add(e);
        }

        if (newErrors.size > 0) {
            return { valid: false, errors: newErrors };
        }
        return { valid: true, errors: new Set() };
    }

    /**
     * Eagerly verify that `libxmljs2` is loadable and that XSD validation is
     * available on this host. Call this once at program startup if you want
     * loud failures when the native binding is missing — otherwise the per-file
     * pipeline silently turns the same conditions into per-file validation
     * errors (matching Python's bare-`except` behavior, which is needed for
     * known-noisy schemas like `docProps/core.xml` that depend on `dcterms`
     * imports the libxmljs2 build cannot fully resolve).
     *
     * Throws an Error prefixed with `libxmljs2 required` when validation
     * cannot be performed. Returns silently on success.
     */
    static assertLibxmljsAvailable(): void {
        try {
            const xsd = '<xs:schema xmlns:xs="http://www.w3.org/2001/XMLSchema"><xs:element name="r" type="xs:string"/></xs:schema>';
            const xsdDoc = libxmljs.parseXml(xsd);
            const doc = libxmljs.parseXml("<r>ok</r>");
            const ok = doc.validate(xsdDoc);
            if (ok !== true) {
                throw new Error(`validate() returned ${String(ok)} on a known-good doc`);
            }
        } catch (err) {
            const message = err instanceof Error ? err.message : String(err);
            throw new Error(`libxmljs2 required for XSD validation: ${message}`);
        }
    }

    protected _validateSingleFileXsd(xmlFile: string): XsdValidationOutcome {
        const schemaPath = this._getSchemaPath(xmlFile);
        if (!schemaPath) return { valid: null, errors: new Set() };

        // Catch-all that mirrors the Python `except Exception` block. Anything
        // that goes wrong during schema loading, XML preprocessing, or validation
        // becomes a single per-file validation error so the caller can decide
        // whether to ignore it via `IGNORED_VALIDATION_ERRORS` (which uses simple
        // substring matching against this string).
        //
        // For loud failure on a missing/broken libxmljs2 binding, call
        // `BaseSchemaValidator.assertLibxmljsAvailable()` once at startup.
        try {
            const xsdDoc = BaseSchemaValidator._loadXsd(schemaPath);

            const xmlContent = readFileSync(xmlFile, "utf-8");
            const cleanedString = this._preprocessXmlForXsd(xmlContent, xmlFile);
            const xmlLibDoc = libxmljs.parseXml(cleanedString);

            const valid = xmlLibDoc.validate(xsdDoc);
            if (valid) {
                return { valid: true, errors: new Set() };
            }
            const errors = new Set<string>();
            const errs =
                (
                    xmlLibDoc as unknown as {
                        validationErrors?: Array<{ message: string }>;
                    }
                ).validationErrors ?? [];
            for (const e of errs) {
                errors.add((e.message ?? "").trim());
            }
            return { valid: false, errors };
        } catch (err) {
            const message = err instanceof Error ? err.message : String(err);
            return { valid: false, errors: new Set([message]) };
        }
    }

    // Process-wide cache of parsed XSDs. The OOXML schema bundle is ~1.1 MB and
    // gets re-used across every file in a package; without this every file
    // validation re-parses the same XSD.
    private static readonly _xsdCache = new Map<string, libxmljs.Document>();

    private static _loadXsd(schemaPath: string): libxmljs.Document {
        const abs = path.resolve(schemaPath);
        const hit = BaseSchemaValidator._xsdCache.get(abs);
        if (hit) return hit;
        const content = readFileSync(abs, "utf-8");
        // baseUrl lets `<xs:include>` / `<xs:import>` resolve siblings.
        const doc = libxmljs.parseXml(content, { baseUrl: abs });
        BaseSchemaValidator._xsdCache.set(abs, doc);
        return doc;
    }

    /**
     * Pre-process XML before XSD validation: strips `{{template}}` tags from
     * non-`<*:t>` text nodes, removes the `mc:Ignorable` attribute, and (for
     * files inside `word/`, `ppt/`, `xl/`) strips elements/attributes belonging
     * to namespaces not in `OOXML_NAMESPACES`.
     *
     * Mirrors `_remove_template_tags_from_text_nodes` + `_preprocess_for_mc_ignorable`
     * + `_clean_ignorable_namespaces` in Python.
     */
    protected _preprocessXmlForXsd(xmlContent: string, xmlFile: string): string {
        const dom = parseXml(xmlContent);

        // _remove_template_tags_from_text_nodes
        for (const elem of iterElements(dom.documentElement)) {
            const tag = String(elem.tagName ?? elem.nodeName);
            if (tag === "t" || tag.endsWith(":t") || tag.endsWith("}t")) continue;
            const children = elem.childNodes;
            for (let i = 0; i < children.length; i += 1) {
                const child = children.item(i);
                if (!child) continue;
                if (child.nodeType === 3 /* TEXT_NODE */) {
                    const textNode = child as Text;
                    const v = textNode.data ?? "";
                    if (TEMPLATE_TAG_PATTERN.test(v)) {
                        // @xmldom Text nodes ignore writes to nodeValue at serialization
                        // time — assign via .data so the change round-trips.
                        textNode.data = v.replace(TEMPLATE_TAG_PATTERN_GLOBAL, "");
                    }
                }
            }
        }

        // _preprocess_for_mc_ignorable
        const root = dom.documentElement;
        if (root) {
            const attrs = root.attributes;
            for (let i = attrs.length - 1; i >= 0; i -= 1) {
                const attr = attrs.item(i);
                if (!attr) continue;
                if (attr.namespaceURI === MC_NAMESPACE && localName(attr.name) === "Ignorable") {
                    root.removeAttributeNode(attr);
                }
            }
        }

        // _clean_ignorable_namespaces — only for word/ppt/xl files
        const rel = this.relPath(xmlFile);
        const firstSegment = rel.split(/[\\/]/)[0];
        if (MAIN_CONTENT_FOLDERS.has(firstSegment) && root) {
            stripNonOoxmlAttributes(root);
            removeNonOoxmlElements(root);
        }

        return serializeXml(dom, "UTF-8");
    }

    protected _getSchemaPath(xmlFile: string): string | null {
        const base = path.basename(xmlFile);
        if (SCHEMA_MAPPINGS[base]) {
            return path.join(this.schemasDir, SCHEMA_MAPPINGS[base]);
        }
        if (xmlFile.endsWith(".rels")) {
            return path.join(this.schemasDir, SCHEMA_MAPPINGS[".rels"]);
        }
        if (xmlFile.includes("charts/") && base.startsWith("chart")) {
            return path.join(this.schemasDir, SCHEMA_MAPPINGS["chart"]);
        }
        if (xmlFile.includes("theme/") && base.startsWith("theme")) {
            return path.join(this.schemasDir, SCHEMA_MAPPINGS["theme"]);
        }
        if (xmlFile.includes("customXml/") && base.startsWith("itemProps") && base.endsWith(".xml")) {
            return path.join(this.schemasDir, SCHEMA_MAPPINGS["itemProps"]);
        }
        if (xmlFile.includes("_xmlsignatures/") && base.startsWith("sig") && base.endsWith(".xml")) {
            return path.join(this.schemasDir, SCHEMA_MAPPINGS["sig"]);
        }
        const parentName = path.basename(path.dirname(xmlFile));
        if (MAIN_CONTENT_FOLDERS.has(parentName) && SCHEMA_MAPPINGS[parentName]) {
            return path.join(this.schemasDir, SCHEMA_MAPPINGS[parentName]);
        }
        return null;
    }

    protected async _getOriginalFileErrors(xmlFile: string): Promise<Set<string>> {
        if (!this.originalFile) return new Set();

        // Resolve the relative path inside the unpacked dir; locate the same path
        // inside the original .docx zip and validate it standalone.
        const rel = path.relative(this.unpackedDir, xmlFile);

        const { default: JSZip } = await import("jszip");
        let zipContent: Buffer;
        try {
            zipContent = await fs.readFile(this.originalFile);
        } catch {
            return new Set();
        }
        const zip = await JSZip.loadAsync(zipContent);
        const entry = zip.file(rel.replace(/\\/g, "/"));
        if (!entry) return new Set();

        return withTempDir(async (tmpDir) => {
            // Preserve the relative path so `_validateSingleFileXsd()` can still
            // pick the right schema based on directory context (e.g. parts under
            // charts/, theme/, or word/embeddings/). Writing the entry to
            // tmpDir/<basename> alone drops that scope and reports unchanged
            // errors as new on every diff.
            const relPosix = rel.split(path.sep).join("/");
            const tmpFile = path.join(tmpDir, relPosix);
            await fs.mkdir(path.dirname(tmpFile), { recursive: true });
            await fs.writeFile(tmpFile, await entry.async("nodebuffer"));
            // Reuse the same single-file XSD validator — we do NOT diff again here
            // (`_validateSingleFileXsd` returns the raw set of errors).
            const outcome = this._validateSingleFileXsd(tmpFile);
            return outcome.errors;
        });
    }

    // ----- helpers ------------------------------------------------------------

    protected relPath(absPath: string): string {
        return path.relative(this.unpackedDir, absPath) || path.basename(absPath);
    }
}

// ===== module-level helpers ===================================================

/** Walk `dir` recursively, returning absolute paths. Optional `extensions` (with leading dot) filter. */
export function walkFiles(dir: string, extensions?: ReadonlyArray<string>): string[] {
    const results: string[] = [];
    const stack: string[] = [dir];
    while (stack.length > 0) {
        const current = stack.pop();
        if (current === undefined) break;
        let entries: string[];
        try {
            entries = readdirSync(current);
        } catch {
            continue;
        }
        for (const name of entries) {
            const full = path.join(current, name);
            let st;
            try {
                st = statSync(full);
            } catch {
                continue;
            }
            if (st.isDirectory()) {
                stack.push(full);
            } else if (st.isFile()) {
                if (!extensions || extensions.some((e) => full.endsWith(e))) {
                    results.push(full);
                }
            }
        }
    }
    return results.sort();
}

function localName(qname: string): string {
    if (!qname) return qname;
    if (qname.includes("}")) return qname.split("}").pop() as string;
    if (qname.includes(":")) return qname.split(":").pop() as string;
    return qname;
}

function* iterElements(root: Element | null): IterableIterator<Element> {
    if (!root) return;
    yield root;
    const stack: Element[] = [root];
    while (stack.length > 0) {
        const node = stack.pop();
        if (node === undefined) break;
        const children = node.childNodes;
        for (let i = children.length - 1; i >= 0; i -= 1) {
            const child = children.item(i);
            if (child && child.nodeType === 1 /* ELEMENT_NODE */) {
                yield child as Element;
                stack.push(child as Element);
            }
        }
    }
}

export function collectDeclaredPrefixes(root: Element): Set<string> {
    // lxml's nsmap propagates xmlns:* declarations from *all* descendants, not
    // just the root element. Walk every element and collect xmlns:prefix attrs to
    // match `set(root.nsmap.keys()) - {None}` from the Python source.
    const declared = new Set<string>();
    const stack: Element[] = [root];
    while (stack.length > 0) {
        const elem = stack.pop();
        if (elem === undefined) break;
        const attrs = elem.attributes;
        for (let i = 0; i < attrs.length; i += 1) {
            const attr = attrs.item(i);
            if (!attr) continue;
            if (attr.name.startsWith("xmlns:")) {
                declared.add(attr.name.slice("xmlns:".length));
            }
        }
        const children = elem.childNodes;
        for (let i = 0; i < children.length; i += 1) {
            const child = children.item(i);
            if (child && child.nodeType === 1 /* ELEMENT_NODE */) {
                stack.push(child as Element);
            }
        }
    }
    return declared;
}

function removeMcAlternateContent(dom: Document): void {
    const list = dom.getElementsByTagNameNS(MC_NAMESPACE, "AlternateContent");
    const items: Element[] = [];
    for (let i = 0; i < list.length; i += 1) {
        const e = list.item(i);
        if (e) items.push(e);
    }
    for (const elem of items) {
        elem.parentNode?.removeChild(elem);
    }
}

function hasExcludedAncestor(elem: Element): boolean {
    let cur: Node | null = elem.parentNode;
    while (cur && cur.nodeType === 1) {
        const tag = localName((cur as Element).tagName ?? cur.nodeName).toLowerCase();
        if (EXCLUDED_ID_CONTAINERS.has(tag)) return true;
        cur = cur.parentNode;
    }
    return false;
}

function findIdAttribute(elem: Element, attrName: string): string | null {
    const attrs = elem.attributes;
    for (let i = 0; i < attrs.length; i += 1) {
        const attr = attrs.item(i);
        if (!attr) continue;
        const local = localName(attr.name).toLowerCase();
        if (local === attrName) return attr.value;
    }
    return null;
}

function stripNonOoxmlAttributes(root: Element): void {
    for (const elem of iterElements(root)) {
        const attrs = elem.attributes;
        const toRemove: Attr[] = [];
        for (let i = 0; i < attrs.length; i += 1) {
            const attr = attrs.item(i);
            if (!attr) continue;
            const ns = attr.namespaceURI;
            if (ns && !OOXML_NAMESPACES.has(ns)) {
                toRemove.push(attr);
            }
        }
        for (const attr of toRemove) {
            elem.removeAttributeNode(attr);
        }
    }
}

function removeNonOoxmlElements(root: Element): void {
    // DFS, removing element children whose namespace is outside OOXML_NAMESPACES.
    const stack: Element[] = [root];
    while (stack.length > 0) {
        const node = stack.pop();
        if (node === undefined) break;
        const children = node.childNodes;
        const toRemove: Element[] = [];
        for (let i = 0; i < children.length; i += 1) {
            const child = children.item(i);
            if (!child || child.nodeType !== 1) continue;
            const elem = child as Element;
            const ns = elem.namespaceURI;
            if (ns && !OOXML_NAMESPACES.has(ns)) {
                toRemove.push(elem);
            } else {
                stack.push(elem);
            }
        }
        for (const e of toRemove) node.removeChild(e);
    }
}

function finalize(issues: ValidationIssue[]): ValidationResult {
    return { valid: issues.every((i) => i.severity !== "error"), issues };
}

/** Return true if `node` has at least one Element-type child node. */
export function hasAnyElementChild(node: Node): boolean {
    for (let i = 0; i < node.childNodes.length; i += 1) {
        const child = node.childNodes.item(i);
        if (child && child.nodeType === 1 /* ELEMENT_NODE */) return true;
    }
    return false;
}
