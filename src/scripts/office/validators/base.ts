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
 *   - `lxml.etree.XMLSchema` → `libxml2-wasm` via the `XsdValidator`
 *     interface from `lib/xsd-validator.ts`. (Earlier: the native
 *     `libxmljs2` addon — dropped when it stopped building against
 *     Node 26's V8 headers and shipped no ABI-147 prebuild.)
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

import { existsSync, promises as fs, readdirSync, readFileSync, realpathSync, statSync } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

import { normalizePartName, parseContentTypes, partExtension, resolveContentType } from "../../../lib/content-types";
import { applyMarkupCompatibility, STRICT_UNDERSTOOD_NAMESPACES, UNDERSTOOD_NAMESPACES } from "../../../lib/markup-compatibility";
import { withTempDir } from "../../../lib/run-cli";
import type { Profile, ValidationIssue, ValidationResult, XsdWalkTrace } from "../../../lib/types";
import { DEFAULT_PROFILE, NS, OK_RESULT } from "../../../lib/types";
import { allowSchemaResourceRoot, createXsdValidator, type XsdValidator } from "../../../lib/xsd-validator";
import {
    decodeXmlBytes,
    getElementsByTagNameNSAll,
    getElementsByTagNameAll,
    parseXml,
    parseXmlBytes,
    serializeXml,
} from "../../../lib/xml-helpers";
import { checkXmlSyntax } from "../../../lib/xml-syntax";
import { resolveSafeZipEntry } from "../../../lib/zip-path";

/**
 * Schema mapping table — file basename / suffix → relative path inside the
 * `schemas/` directory. Mirrors `BaseSchemaValidator.SCHEMA_MAPPINGS`.
 */
const SCHEMA_MAPPINGS: Record<string, string> = {
    word: "ISO-IEC29500-4_2016/wml.xsd",
    ppt: "ISO-IEC29500-4_2016/pml.xsd",
    xl: "ISO-IEC29500-4_2016/sml.xsd",
    "[Content_Types].xml": "ecma/fourth-edition/opc-contentTypes.xsd",
    "app.xml": "ISO-IEC29500-4_2016/shared-documentPropertiesExtended.xsd",
    "core.xml": "ecma/fourth-edition/opc-coreProperties.xsd",
    "custom.xml": "ISO-IEC29500-4_2016/shared-documentPropertiesCustom.xsd",
    ".rels": "ecma/fourth-edition/opc-relationships.xsd",
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
    sig: "ecma/fourth-edition/opc-digSig.xsd",
};

// These namespaces are understood only when validating their standalone
// mapped part schemas. wml.xsd does not import these Microsoft extensions.
const SCHEMA_NAMESPACE_CAPABILITIES: Readonly<Record<string, string>> = {
    "wml-2012.xsd": NS.W15,
    "wml-cex-2018.xsd": NS.W16CEX,
    "wml-cid-2016.xsd": NS.W16CID,
};

/**
 * Directory inside `schemas/` holding the bundled ISO/IEC 29500 Strict
 * conformance-class schemas (see `schemas/SOURCES.md`).
 */
const STRICT_SCHEMA_DIR = "iso-strict";

/**
 * Subset of {@link SCHEMA_MAPPINGS} keys that have a Strict-conformance
 * schema under {@link STRICT_SCHEMA_DIR}. OPC parts (`[Content_Types].xml`,
 * `.rels`, `docProps/core.xml`, digital signatures) share one namespace across
 * both conformance classes and always resolve to the OPC schemas; Microsoft
 * extension parts (`commentsExtended.xml`, `people.xml`, …) keep their
 * `microsoft/` schemas. A Strict part whose key is absent here has no
 * applicable schema — it is skipped exactly like an unmapped Transitional
 * part.
 */
const STRICT_SCHEMA_MAPPINGS: Record<string, string> = {
    word: `${STRICT_SCHEMA_DIR}/wml.xsd`,
    ppt: `${STRICT_SCHEMA_DIR}/pml.xsd`,
    xl: `${STRICT_SCHEMA_DIR}/sml.xsd`,
    "app.xml": `${STRICT_SCHEMA_DIR}/shared-documentPropertiesExtended.xsd`,
    "custom.xml": `${STRICT_SCHEMA_DIR}/shared-documentPropertiesCustom.xsd`,
    chart: `${STRICT_SCHEMA_DIR}/dml-chart.xsd`,
    theme: `${STRICT_SCHEMA_DIR}/dml-main.xsd`,
    drawing: `${STRICT_SCHEMA_DIR}/dml-main.xsd`,
    "bibliography.xml": `${STRICT_SCHEMA_DIR}/shared-bibliography.xsd`,
    "additionalCharacteristics.xml": `${STRICT_SCHEMA_DIR}/shared-additionalCharacteristics.xsd`,
    itemProps: `${STRICT_SCHEMA_DIR}/shared-customXmlDataProperties.xsd`,
};

const IGNORED_VALIDATION_ERRORS: readonly string[] = [
    // Inherited from the Python validator with no recorded reason or
    // fixture; it hides every engine message that mentions the element.
    // Lenient and word-valid keep it for compatibility. Strict never
    // consults this list: its validity rests on unsuppressed schema issues.
    "hyphenationZone",
    // opc-coreProperties.xsd's Dublin Core imports resolve to the bundled
    // schemas/dublincore/ copies (SOURCES.md), so the libxmljs2-era
    // suppressions for the missing remote import ("purl.org/dc/terms",
    // "dublincore.org", "Invalid XSD schema") are gone — schema-load
    // failures must surface as xsd-schema-load-failed, never a suppressed
    // pass.
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
const MAIN_CONTENT_FOLDERS = new Set(["word", "ppt", "xl"]);

/**
 * URI prefix shared by all ISO OOXML Strict conformance class namespaces.
 * Strict documents use `purl.oclc.org/ooxml/` instead of the Transitional
 * `schemas.openxmlformats.org/…` prefix. Any root element namespace that
 * starts with this prefix belongs to the Strict conformance class.
 *
 * The bundled `iso-strict/` schemas cover these namespaces. Under
 * `profile: "strict"` `validateAgainstXsd` validates Strict parts against
 * them (namespace + part identity dispatch); lenient and word-valid keep
 * the historical info-level `xsd-strict-skipped` skip.
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
    /**
     * Optional XSD validator override. When omitted, BaseSchemaValidator
     * lazy-creates the default wasm-backed validator from
     * `src/lib/xsd-validator.ts`. Test code may inject a fake; consumer
     * code may swap in a different engine.
     *
     * Part of the four-class architecture cutover (PR B). The wasm default
     * replaced the retired libxmljs2 path outright (PR B Task B.2) — the
     * native addon cannot build against Node 26.
     */
    xsdValidator?: XsdValidator;
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

/**
 * Containment check for relationship targets resolved from untrusted OOXML
 * parts (CWE-22). Two layers:
 *
 * 1. Lexical: the resolved path must stay under `unpackedDir`
 *    (`path.relative` must not escape via `..` or go absolute).
 * 2. Symlinks: a purely lexical check is defeated by a symlink inside the
 *    unpacked tree pointing outward (`link/secret` where `link` → `/etc`).
 *    The deepest EXISTING ancestor of the target is resolved with
 *    `realpathSync` and re-checked; non-existent trailing components cannot
 *    change the verdict (callers report missing targets as broken anyway).
 *
 * Returns the original resolved path when contained, `null` on escape.
 * Not race-resistant across the check and a later open — acceptable for a
 * validator scanning an unpacked tree it controls; flagged for any future
 * fd-based (`O_NOFOLLOW`) hardening.
 */
export function containRelationshipTarget(unpackedDir: string, resolvedPath: string): string | null {
    const root = path.resolve(unpackedDir);
    const target = path.resolve(resolvedPath);
    const rel = path.relative(root, target);
    if (rel === ".." || rel.startsWith(`..${path.sep}`) || path.isAbsolute(rel)) return null;

    let realRoot = root;
    try {
        realRoot = realpathSync(root);
    } catch {
        // Root itself unreachable — the lexical verdict stands.
        return target;
    }
    let probe = target;
    for (;;) {
        let real: string | null = null;
        try {
            real = realpathSync(probe);
        } catch {
            real = null;
        }
        if (real !== null) {
            const realRel = path.relative(realRoot, real);
            if (realRel === ".." || realRel.startsWith(`..${path.sep}`) || path.isAbsolute(realRel)) return null;
            break;
        }
        const parent = path.dirname(probe);
        if (parent === probe) break;
        probe = parent;
    }
    return target;
}

interface IdHit {
    file: string;
    tag: string;
}

export interface XsdValidationOutcome {
    /** `null` mirrors Python's `(None, None)` skip — no schema configured. */
    valid: boolean | null;
    errors: Set<string>;
    /** Compatibility diagnostics keep their own stable codes, independent of XSD messages. */
    issues?: ValidationIssue[];
    /**
     * Structured engine findings for the part, one entry per error
     * occurrence with `line` when libxml2 reports it. Populated by
     * `_validateSingleFileXsd` on failure. The `strict` profile emits these
     * verbatim (deduplicated message strings would collapse repeated
     * identical defects onto one report and lose their locations).
     */
    schemaIssues?: ValidationIssue[];
    /**
     * True when the XML sent to the engine differs from the part bytes
     * (template-tag stripping, the markup-compatibility view, or
     * re-serialization). Reported `line` values then refer to the
     * transformed view, not the source part.
     */
    transformedView?: boolean;
    /**
     * True when the failure is an infrastructure finding — the schema failed
     * to load (`xsd-schema-load-failed`) or the engine itself failed — rather
     * than a document defect. Baseline subtraction against `--original` must
     * never erase these: they surface as error-severity issues and keep the
     * part invalid under every profile (#289).
     */
    infrastructure?: boolean;
}

export class BaseSchemaValidator {
    readonly unpackedDir: string;
    readonly originalFile: string | null;
    readonly verbose: boolean;
    readonly schemasDir: string;
    readonly xmlFiles: string[];
    readonly profile: Profile;
    /**
     * Engine trace for the most recent `validateAgainstXsd` on this instance.
     * `DOCXSchemaValidator.validate` resets it at entry. `mergeResults` drops
     * it, so the public `validate()` copies it off the instance afterwards.
     */
    xsdWalk: XsdWalkTrace = { started: false, parts: [], eligible: 0 };

    /**
     * Parts the running `validate()` found not well-formed. Their XML
     * diagnostics already reject the package, so the schema walk does not
     * parse them again.
     */
    protected malformedParts: ReadonlySet<string> = new Set();

    /**
     * Parts whose schema the engine could not load, or whose XSD call threw,
     * during the most recent `validateAgainstXsd`. Reset at that method's
     * entry, like `xsdWalk`.
     *
     * `xsdWalk.parts.length === xsdWalk.eligible` only proves every eligible
     * part was *offered* to the engine; this counter is what proves the engine
     * actually answered for it. A schema that fails to compile still counts as
     * a completed walk, which is why the top-level `validate()` also requires
     * this to be 0 before reporting `schemaCompleted: true`.
     */
    xsdEngineFailures = 0;

    /**
     * Subclasses override to map an element name → expected relationship type
     * (used by `validateAllRelationshipIds`). Mirrors
     * `BaseSchemaValidator.ELEMENT_RELATIONSHIP_TYPES`.
     */
    protected readonly elementRelationshipTypes: Record<string, string> = {};

    /**
     * Injected XSD validator. Lazily resolved to the wasm default by
     * `_getXsdValidator()` on first use. Stored as `XsdValidator | undefined`
     * (not a Promise) so the field can hold either a synchronously-supplied
     * test fake or the lazily-resolved wasm singleton.
     */
    private xsdValidatorRef: XsdValidator | undefined;
    private xsdValidatorPromise: Promise<XsdValidator> | undefined;

    constructor(opts: BaseSchemaValidatorOptions) {
        this.unpackedDir = path.resolve(opts.unpackedDir);
        this.originalFile = opts.originalFile ? path.resolve(opts.originalFile) : null;
        this.verbose = opts.verbose ?? false;
        this.schemasDir = opts.schemasDir ?? defaultSchemasDir();
        this.xmlFiles = walkFiles(this.unpackedDir, [".xml", ".rels"]);
        this.profile = opts.profile ?? DEFAULT_PROFILE;
        this.xsdValidatorRef = opts.xsdValidator;
    }

    /**
     * Returns the active XSD validator. If the constructor was given a
     * `xsdValidator`, that wins. Otherwise the wasm default is lazily
     * created (and memoised across calls within this instance).
     *
     * The single XSD entrypoint since the libxmljs2 → libxml2-wasm cutover
     * (PR B Task B.2): the native addon neither ships an ABI-147 prebuild
     * nor compiles against Node 26's V8 headers, so it was removed.
     */
    protected async _getXsdValidator(): Promise<XsdValidator> {
        if (this.xsdValidatorRef) {
            return this.xsdValidatorRef;
        }
        // The engine's file access is restricted to allowed schema roots; a
        // caller-supplied `schemasDir` must be one of them or its import
        // closure would be refused.
        allowSchemaResourceRoot(this.schemasDir);
        if (!this.xsdValidatorPromise) {
            this.xsdValidatorPromise = createXsdValidator();
        }
        const resolved = await this.xsdValidatorPromise;
        this.xsdValidatorRef = resolved;
        return resolved;
    }

    /**
     * Every XML part should NOT begin with a UTF-8 BOM (U+FEFF, bytes EF BB
     * BF). XML 1.0 §4.3.3 permits the BOM and Microsoft Office emits it
     * routinely, so this is reported at `warning` severity under EVERY
     * profile — it never fails a document on its own. The finding stays
     * visible in `result.issues` so callers who care about canonical bytes can
     * audit it.
     */
    async validateNoBom(): Promise<ValidationResult> {
        const issues: ValidationIssue[] = [];
        for (const xmlFile of this.xmlFiles) {
            let bytes: Buffer;
            try {
                bytes = await fs.readFile(xmlFile);
            } catch {
                continue;
            }
            if (bytes.length >= 3 && bytes[0] === 0xef && bytes[1] === 0xbb && bytes[2] === 0xbf) {
                issues.push({
                    severity: "warning",
                    message:
                        "XML part begins with a UTF-8 BOM (U+FEFF). Permitted by the XML spec (and emitted by Microsoft Office), " +
                        "so it is reported as a warning under every profile and never fails a document on its own.",
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
                dom = parseXmlBytes(await fs.readFile(xmlFile));
            } catch {
                // Malformed rels files surface elsewhere via xml-syntax /
                // rels-broken codes; don't double-report here.
                continue;
            }
            const root = dom.documentElement;
            if (!root) continue;
            const rels = getElementsByTagNameNSAll(root, "http://schemas.openxmlformats.org/package/2006/relationships", "Relationship");
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
                const dom = parseXmlBytes(await fs.readFile(xmlFile));
                const root = dom.documentElement;
                if (root) {
                    const rels = getElementsByTagNameNSAll(
                        root,
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
                const content = decodeXmlBytes(await fs.readFile(xmlFile));
                const dom = parseXml(content);
                let modified = false;

                const all = getElementsByTagNameAll(dom);
                for (let i = 0; i < all.length; i += 1) {
                    const elem = all[i];
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

    /**
     * Well-formedness of every XML part, checked on the ORIGINAL BYTES.
     *
     * libxml2 parses bytes, so the part's own BOM and `encoding` declaration
     * decide how it decodes (a UTF-16 part read as a UTF-8 string is garbage)
     * and every diagnostic carries a 1-based line/column. Parsing runs
     * WITHOUT `XML_PARSE_RECOVER` so content `@xmldom` silently recovers from
     * (a bare `&`, an unquoted attribute value) is reported instead of
     * accepted, and WITH `XML_PARSE_NONET | XML_PARSE_NO_XXE` so external
     * entities are never fetched. See `src/lib/xml-syntax.ts`.
     *
     * Codes: `xml-syntax` (not well-formed), `xml-undeclared-prefix`
     * (undeclared namespace prefix), `xml-encoding` (declared encoding
     * unsupported / bytes do not decode in it).
     */
    async validateXml(): Promise<ValidationResult> {
        const issues: ValidationIssue[] = [];
        for (const xmlFile of this.xmlFiles) {
            let raw: Buffer;
            try {
                raw = await fs.readFile(xmlFile);
            } catch {
                continue;
            }
            issues.push(...(await checkXmlSyntax(raw, this.relPath(xmlFile))));
        }
        return finalize(issues);
    }

    async validateNamespaces(): Promise<ValidationResult> {
        const issues: ValidationIssue[] = [];
        for (const xmlFile of this.xmlFiles) {
            let dom: Document;
            try {
                dom = parseXmlBytes(await fs.readFile(xmlFile));
            } catch {
                continue;
            }
            const root = dom.documentElement;
            if (!root) continue;

            const declared = collectDeclaredPrefixes(root);

            const attrs = root.attributes;
            for (let i = 0; i < attrs.length; i += 1) {
                const attr = attrs[i];
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
                dom = parseXmlBytes(await fs.readFile(xmlFile));
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
                dom = parseXmlBytes(await fs.readFile(relsFile));
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

            const rels = getElementsByTagNameNSAll(dom, PACKAGE_RELATIONSHIPS_NAMESPACE, "Relationship");
            for (let i = 0; i < rels.length; i += 1) {
                const rel = rels[i];
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
                    if (containRelationshipTarget(this.unpackedDir, targetPath) === null) {
                        broken.push(target);
                        continue;
                    }
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
                dom = parseXmlBytes(await fs.readFile(relsFile));
            } catch {
                continue;
            }

            const rels = getElementsByTagNameNSAll(dom, PACKAGE_RELATIONSHIPS_NAMESPACE, "Relationship");
            for (let i = 0; i < rels.length; i += 1) {
                const rel = rels[i];
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
                    xmlDomCheck = parseXmlBytes(await fs.readFile(xmlFile));
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
                relsDom = parseXmlBytes(await fs.readFile(relsFile));
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
            const rels = getElementsByTagNameNSAll(relsDom, PACKAGE_RELATIONSHIPS_NAMESPACE, "Relationship");
            for (let i = 0; i < rels.length; i += 1) {
                const rel = rels[i];
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
                xmlDom = parseXmlBytes(await fs.readFile(xmlFile));
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
            dom = parseXmlBytes(await fs.readFile(contentTypesFile));
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

        // Single OPC resolution table (Override first, then Default —
        // ECMA-376 Part 2 §10.1.2.4) shared by every coverage pass below:
        // declarable XML roots (ct-undeclared-part), known media extensions
        // (ct-undeclared-ext) and the residual invariant (ct-uncovered-part).
        const table = parseContentTypes(dom);

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
                const root = parseXmlBytes(await fs.readFile(xmlFile)).documentElement;
                if (!root) continue;
                const rootName = localName(root.tagName ?? root.nodeName);
                // A declarable root must be covered by an <Override PartName>
                // (OPC-equivalent match — case-insensitive, percent-decoded);
                // coverage through a bare <Default Extension> is not enough
                // for parts Word opens by name.
                if (declarableRoots.has(rootName) && resolveContentType(table, pathStr)?.source !== "override") {
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

            const rel = this.relPath(filePath).replace(/\\/g, "/");
            const extension = ext.replace(/^\./, "");
            // An exact <Override PartName> covers a media part just as well
            // as a <Default Extension> — OPC §10.1.2.4 resolves Override
            // first — so resolve through the table instead of consulting
            // only the Default extensions (jubarte redlines declare every
            // media part by Override; issues #208/#237).
            // Object.hasOwn, not `in`: mediaExtensions is a plain object, so
            // `"constructor" in mediaExtensions` is true and a part named
            // x.constructor would sail past the check (and be "suggested"
            // with ContentType="undefined").
            if (Object.hasOwn(mediaExtensions, extension) && resolveContentType(table, rel) === null) {
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

        // Residual OPC-coverage pass. The two checks above are Word-behaviour
        // heuristics — declarable XML roots (ct-undeclared-part) and known
        // media extensions (ct-undeclared-ext) — and deliberately skip what
        // they don't recognise. This pass enforces the underlying spec
        // invariant instead: every part must be matched by an
        // <Override PartName> or a <Default Extension> in [Content_Types].xml.
        // The recurring real-world case is extensionless media
        // (word/media/image_3094991244). Known media extensions are left to
        // ct-undeclared-ext above, which reports them with a fix-it message.
        for (const filePath of allFiles) {
            if (path.basename(filePath) === "[Content_Types].xml") continue;
            const rel = this.relPath(filePath).replace(/\\/g, "/");
            // One resolver call subsumes the old set lookups: an OPC-
            // equivalent <Override PartName> (case-insensitive, decoded) or
            // a <Default Extension> for the part's OPC extension covers it.
            if (resolveContentType(table, rel) !== null) continue;
            // OPC extension = substring after the last dot of the part name's
            // last segment — unlike POSIX, `_rels/.rels` has extension "rels".
            const extension = partExtension(rel);
            if (Object.hasOwn(mediaExtensions, extension)) continue;
            issues.push({
                severity: "error",
                message:
                    `Part '${rel}' has no content type: matched by no <Override PartName> and no ` +
                    `<Default Extension> in [Content_Types].xml. OPC requires a content type for every ` +
                    `part; Word prompts to repair such packages.`,
                path: rel,
                code: "ct-uncovered-part",
            });
        }

        return finalize(issues);
    }

    // ----- XSD validation -----------------------------------------------------

    async validateAgainstXsd(): Promise<ValidationResult> {
        const issues: ValidationIssue[] = [];
        let strictSkipped = false;
        this.xsdWalk.started = true;
        this.xsdWalk.parts = [];
        this.xsdWalk.eligible = 0;
        this.xsdEngineFailures = 0;

        const noteStrict = (): void => {
            if (strictSkipped) return;
            strictSkipped = true;
            issues.push({
                severity: "info",
                message:
                    "Document uses ISO OOXML Strict conformance class namespace URIs; " +
                    "XSD validation skipped (only Transitional schemas are bundled). " +
                    "See PORT_NOTES.md § 'ISO OOXML Strict XSD validation'.",
                code: "xsd-strict-skipped",
            });
        };

        const countEligible = (xmlFile: string): boolean => {
            if (this.malformedParts.has(this.relPath(xmlFile))) return false;
            // Strict-conformance parts are only skipped outside the strict
            // profile: lenient keeps the documented `xsd-strict-skipped` info
            // issue, while `strict` schema-validates them against the bundled
            // iso-strict schemas. Word-valid never walks the schemas.
            if (this._isStrictXmlFile(xmlFile) && this.profile !== "strict") {
                noteStrict();
                return false;
            }
            if (!this._getSchemaPath(xmlFile)) return false;
            this.xsdWalk.eligible += 1;
            return true;
        };

        for (const xmlFile of this.xmlFiles) {
            if (!countEligible(xmlFile)) continue;

            const outcome = await this.validateFileAgainstXsd(xmlFile);
            issues.push(...(outcome.issues ?? []));
            const rel = this.relPath(xmlFile);
            this.xsdWalk.parts.push(rel);
            if (outcome.valid === null || outcome.valid) continue;

            issues.push({
                severity: "error",
                message: `${rel}: ${outcome.errors.size} new error(s)`,
                path: rel,
                code: "xsd-summary",
            });
            if (this.profile === "strict") {
                // One issue per occurrence, with the engine's own line and
                // code preserved. When preprocessing rewrote the XML the
                // reported line belongs to the transformed view, not the
                // source part — say so instead of inventing source lines.
                for (const issue of outcome.schemaIssues ?? []) {
                    const message = issue.message.trim();
                    const truncated = message.length > 250 ? `${message.slice(0, 250)}...` : message;
                    const suffix =
                        outcome.transformedView && issue.line !== undefined ? " (line refers to the preprocessed validation view)" : "";
                    issues.push({
                        severity: "error",
                        message: `  - ${truncated}${suffix}`,
                        path: rel,
                        ...(issue.line !== undefined ? { line: issue.line } : {}),
                        code: !issue.code || issue.code === "xsd-validation-failed" ? "xsd-error" : issue.code,
                    });
                }
            } else {
                for (const e of [...outcome.errors].slice(0, 3)) {
                    const truncated = e.length > 250 ? `${e.slice(0, 250)}...` : e;
                    issues.push({
                        severity: "error",
                        message: `  - ${truncated}`,
                        path: rel,
                        code: "xsd-error",
                    });
                }
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
        const cached = this.strictRootCache.get(xmlFile);
        if (cached !== undefined) return cached;
        const strict = this.detectStrictRoot(xmlFile);
        this.strictRootCache.set(xmlFile, strict);
        return strict;
    }

    private detectStrictRoot(xmlFile: string): boolean {
        try {
            const content = decodeXmlBytes(readFileSync(xmlFile));
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
        const single = await this._validateSingleFileXsd(xmlFile);
        if (single.valid === null) {
            return { valid: null, errors: new Set() };
        }
        if (single.valid) {
            return single;
        }

        // Infrastructure failures (schema load / engine failure) are never
        // part of the --original baseline subtraction: they surface as
        // error-severity issues and keep the part invalid under every
        // profile (#289).
        if (single.infrastructure) {
            return single;
        }

        if (this.profile === "strict") {
            // Every structured finding survives — no baseline subtraction, no
            // message suppression, no dedup of repeated messages — so repeated
            // identical defects keep their own location.
            const schemaIssues = single.schemaIssues ?? [];
            if (schemaIssues.length === 0) {
                return { ...single, valid: true, errors: new Set(), schemaIssues: [] };
            }
            return {
                ...single,
                valid: false,
                errors: new Set(schemaIssues.map((issue) => issue.message.trim())),
                schemaIssues,
            };
        }

        // Baseline subtraction diffs the modified part against --original so
        // migrations don't fail on pre-existing defects; lenient/word-valid
        // keep the diff.
        const originalErrors = await this._getOriginalFileErrors(xmlFile);
        const newErrors = new Set<string>();
        for (const e of single.errors) {
            if (originalErrors.has(e)) continue;
            if (IGNORED_VALIDATION_ERRORS.some((p) => e.includes(p))) continue;
            newErrors.add(e);
        }

        if (newErrors.size > 0) {
            return { ...single, valid: false, errors: newErrors };
        }
        return { ...single, valid: true, errors: new Set() };
    }

    /**
     * Eagerly verify that the wasm-backed XSD engine is loadable and that XSD
     * validation is available on this host. Call this once at program startup
     * if you want loud failures when the engine is broken — otherwise the
     * per-file pipeline silently turns the same condition into per-file
     * validation errors that look like document corruption.
     *
     * Throws an Error prefixed with `XSD engine unavailable` when validation
     * cannot be performed. Returns silently on success.
     *
     * Successor of the libxmljs2-era `assertLibxmljsAvailable`, retired with
     * the native addon (no Node 26 / ABI-147 prebuild, no source build).
     */
    static async assertXsdValidationAvailable(): Promise<void> {
        try {
            const validator = await createXsdValidator();
            await withTempDir(async (dir) => {
                const schemaPath = path.join(dir, "smoke.xsd");
                const xsd = '<xs:schema xmlns:xs="http://www.w3.org/2001/XMLSchema"><xs:element name="r" type="xs:string"/></xs:schema>';
                await fs.writeFile(schemaPath, xsd, "utf-8");
                const issues = await validator.validate("<r>ok</r>", schemaPath);
                const errors = issues.filter((i) => i.severity === "error");
                if (errors.length > 0) {
                    throw new Error(`validate() flagged a known-good doc: ${errors[0]?.message ?? "(no message)"}`);
                }
            });
        } catch (err) {
            const message = err instanceof Error ? err.message : String(err);
            throw new Error(`XSD engine unavailable (libxml2-wasm): ${message}`);
        }
    }

    /** Identify the required main part by the package relationship, never its basename. */
    private isMainDocumentPart(xmlFile: string): boolean {
        try {
            const relationships = parseXml(readFileSync(path.join(this.unpackedDir, "_rels", ".rels"), "utf-8"));
            const part = normalizePartName(this.relPath(xmlFile).split(path.sep).join("/"));
            return getElementsByTagNameNSAll(relationships, PACKAGE_RELATIONSHIPS_NAMESPACE, "Relationship").some((relationship) => {
                const type = relationship.getAttribute("Type");
                if (
                    relationship.getAttribute("TargetMode") === "External" ||
                    (type !== `${OFFICE_RELATIONSHIPS_NAMESPACE}/officeDocument` &&
                        type !== "http://purl.oclc.org/ooxml/officeDocument/relationships/officeDocument")
                ) {
                    return false;
                }
                return path.posix.normalize(normalizePartName(relationship.getAttribute("Target") || "")) === part;
            });
        } catch {
            // Missing or malformed package relationships are diagnosed by their own checks.
            return false;
        }
    }

    protected async _validateSingleFileXsd(xmlFile: string): Promise<XsdValidationOutcome> {
        const schemaPath = this._getSchemaPath(xmlFile);
        if (!schemaPath) return { valid: null, errors: new Set() };

        const compatibilityIssues: ValidationIssue[] = [];
        try {
            const xmlContent = decodeXmlBytes(readFileSync(xmlFile));
            const cleanedString = this._preprocessXmlForXsd(xmlContent, xmlFile, compatibilityIssues);
            if (compatibilityIssues.some((issue) => issue.code === "mce-root-ignored")) {
                const main = this.isMainDocumentPart(xmlFile);
                const issues = compatibilityIssues.map((issue): ValidationIssue =>
                    main && issue.code === "mce-root-ignored"
                        ? {
                              ...issue,
                              severity: "error",
                              code: "mce-main-root-ignored",
                              message: `Required main document part ${this.relPath(xmlFile)} has no root after markup compatibility processing.`,
                          }
                        : issue,
                );
                return { valid: true, errors: new Set(), issues };
            }
            const transformedView = cleanedString !== xmlContent;

            const validator = await this._getXsdValidator();
            const issues = await validator.validate(cleanedString, schemaPath);

            // Non-error issues from the engine stay non-fatal; schema-load
            // failures arrive as errors and are string-filtered by
            // IGNORED_VALIDATION_ERRORS in `validateFileAgainstXsd`.
            const errorIssues = issues.filter((i) => i.severity === "error");
            // A schema the engine could not load means it produced no verdict
            // for this part, so the walk is not "completed". Apply the same
            // IGNORED_VALIDATION_ERRORS filter used one layer up: the
            // known-noisy bundled case (`opc-coreProperties.xsd`'s remote
            // Dublin Core import) is filtered there, and counting it here
            // would mark every document with docProps/core.xml as having an
            // incomplete schema walk.
            const engineFailure = errorIssues.find(
                (i) => i.code === "xsd-schema-load-failed" && !IGNORED_VALIDATION_ERRORS.some((pattern) => i.message.includes(pattern)),
            );
            if (engineFailure) {
                this.xsdEngineFailures += 1;
            }
            if (errorIssues.length === 0) {
                return {
                    valid: true,
                    errors: new Set(),
                    ...(compatibilityIssues.length ? { issues: compatibilityIssues } : {}),
                };
            }
            const errors = new Set<string>();
            for (const issue of errorIssues) {
                errors.add(issue.message.trim());
            }
            const infrastructure = errorIssues.some(
                (issue) => issue.code === "xsd-schema-load-failed" || issue.message.startsWith("XSD engine failed"),
            );
            return {
                valid: false,
                errors,
                schemaIssues: errorIssues,
                transformedView,
                ...(infrastructure ? { infrastructure } : {}),
                ...(compatibilityIssues.length ? { issues: compatibilityIssues } : {}),
            };
        } catch (err) {
            // The engine never produced a verdict for this part: reading the
            // part, preprocessing it, or the engine call itself failed.
            this.xsdEngineFailures += 1;
            const message = err instanceof Error ? err.message : String(err);
            return {
                valid: false,
                errors: new Set([message]),
                // A throw means the engine never measured the part — unreadable
                // bytes, preprocessing failure, or an engine-level error. Every
                // one of those stays blocking: baseline subtraction must not
                // erase it (#289).
                infrastructure: true,
                schemaIssues: [{ severity: "error", message }],
                ...(compatibilityIssues.length ? { issues: compatibilityIssues } : {}),
            };
        }
    }

    /**
     * Pre-process XML before XSD validation: strips `{{template}}` tags from
     * non-`<*:t>` text nodes, then applies markup compatibility to a validation
     * view of parts inside `word/`, `ppt/`, and `xl/`.
     */
    protected _preprocessXmlForXsd(xmlContent: string, xmlFile: string, issues: ValidationIssue[] = []): string {
        const dom = parseXml(xmlContent);

        // _remove_template_tags_from_text_nodes
        for (const elem of iterElements(dom.documentElement)) {
            const tag = String(elem.tagName ?? elem.nodeName);
            if (tag === "t" || tag.endsWith(":t") || tag.endsWith("}t")) continue;
            const children = elem.childNodes;
            for (let i = 0; i < children.length; i += 1) {
                const child = children[i];
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

        const rel = this.relPath(xmlFile);
        const firstSegment = rel.split(/[\\/]/)[0];
        if (MAIN_CONTENT_FOLDERS.has(firstSegment)) {
            const schemaNamespace = SCHEMA_NAMESPACE_CAPABILITIES[path.basename(this._getSchemaPath(xmlFile) ?? "")];
            // A Strict part understands the purl.oclc.org counterparts of the
            // Transitional namespaces its iso-strict schema validates.
            const baseUnderstood = dom.documentElement?.namespaceURI?.startsWith(STRICT_OOXML_NS_PREFIX)
                ? STRICT_UNDERSTOOD_NAMESPACES
                : UNDERSTOOD_NAMESPACES;
            const understood = schemaNamespace ? new Set([...baseUnderstood, schemaNamespace]) : baseUnderstood;
            const compatible = applyMarkupCompatibility(dom, { understood });
            issues.push(
                ...compatible.issues.map((issue) => ({
                    ...issue,
                    path: rel,
                    message: issue.code === "mce-extension-not-ignorable" ? `${rel}: ${issue.message}` : issue.message,
                })),
            );
            return serializeXml(compatible.document, "UTF-8");
        }

        return serializeXml(dom, "UTF-8");
    }

    protected _getSchemaPath(xmlFile: string): string | null {
        const key = this._schemaMappingKey(xmlFile);
        if (!key) return null;
        // Dispatch by namespace AND part identity: a Strict-conformance root
        // resolves to the iso-strict schema for the same key when one exists.
        // Keys without a Strict schema (OPC parts, Microsoft extensions,
        // customXml itemProps is mapped — see STRICT_SCHEMA_MAPPINGS) fall
        // through to null so unmapped Strict parts keep the usual skip.
        if (this._isStrictXmlFile(xmlFile)) {
            const strictMapping = STRICT_SCHEMA_MAPPINGS[key];
            if (!strictMapping) return null;
            return path.join(this.schemasDir, strictMapping);
        }
        return path.join(this.schemasDir, SCHEMA_MAPPINGS[key]);
    }

    /**
     * Part-identity lookup — the {@link SCHEMA_MAPPINGS} key this XML part
     * maps to, independent of conformance class. Mirrors the pre-strict
     * `_getSchemaPath` heuristics exactly.
     */
    protected _schemaMappingKey(xmlFile: string): string | null {
        const base = path.basename(xmlFile);
        if (SCHEMA_MAPPINGS[base]) {
            return base;
        }
        if (xmlFile.endsWith(".rels")) {
            return ".rels";
        }
        if (xmlFile.includes("charts/") && base.startsWith("chart")) {
            return "chart";
        }
        if (xmlFile.includes("theme/") && base.startsWith("theme")) {
            return "theme";
        }
        if (xmlFile.includes("customXml/") && base.startsWith("itemProps") && base.endsWith(".xml")) {
            return "itemProps";
        }
        if (xmlFile.includes("_xmlsignatures/") && base.startsWith("sig") && base.endsWith(".xml")) {
            return "sig";
        }
        const parentName = path.basename(path.dirname(xmlFile));
        if (MAIN_CONTENT_FOLDERS.has(parentName) && SCHEMA_MAPPINGS[parentName]) {
            return parentName;
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
            const tmpFile = resolveSafeZipEntry(tmpDir, relPosix);
            await fs.mkdir(path.dirname(tmpFile), { recursive: true });
            await fs.writeFile(tmpFile, await entry.async("nodebuffer"));
            // Reuse the same single-file XSD validator — we do NOT diff again here
            // (`_validateSingleFileXsd` returns the raw set of errors).
            const outcome = await this._validateSingleFileXsd(tmpFile);
            return outcome.errors;
        });
    }

    // ----- helpers ------------------------------------------------------------

    protected relPath(absPath: string): string {
        return path.relative(this.unpackedDir, absPath) || path.basename(absPath);
    }

    /** Per-file memo for {@link _isStrictXmlFile} — root-namespace sniffing would otherwise re-read each part for eligibility, schema dispatch, and preprocessing. */
    private readonly strictRootCache = new Map<string, boolean>();
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
            const child = children[i];
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
            const attr = attrs[i];
            if (!attr) continue;
            if (attr.name.startsWith("xmlns:")) {
                declared.add(attr.name.slice("xmlns:".length));
            }
        }
        const children = elem.childNodes;
        for (let i = 0; i < children.length; i += 1) {
            const child = children[i];
            if (child && child.nodeType === 1 /* ELEMENT_NODE */) {
                stack.push(child as Element);
            }
        }
    }
    return declared;
}

const IGNORABLE_UNDECLARED_MESSAGE = /^Namespace '(.+)' in Ignorable but not declared$/;
const UNBOUND_IGNORABLE_MESSAGE = /^Unbound prefix in mc:Ignorable: (.+)$/;

/**
 * `validateNamespaces` reports an undeclared prefix in a part root's
 * `Ignorable` as `ignorable-undeclared`, and the markup-compatibility
 * processor reports the same prefix as `mce-unbound-prefix`. Keep the
 * established code, once: one processor finding is dropped per
 * `ignorable-undeclared` on the same part and prefix (#309). A prefix only the
 * processor sees unbound (an `Ignorable` below the root, or a root prefix
 * bound only on a descendant) keeps its finding.
 */
export function withoutRestatedIgnorablePrefixes(result: ValidationResult): ValidationResult {
    // validateNamespaces reports path.relative paths (backslashes on Windows);
    // the Word rule reports the normalized part name.
    const key = (issue: ValidationIssue, prefix: string): string => `${(issue.path ?? "").replace(/\\/g, "/")}\u0000${prefix}`;
    const undeclared = new Map<string, number>();
    for (const issue of result.issues) {
        const prefix = issue.code === "ignorable-undeclared" ? IGNORABLE_UNDECLARED_MESSAGE.exec(issue.message)?.[1] : undefined;
        if (prefix !== undefined) undeclared.set(key(issue, prefix), (undeclared.get(key(issue, prefix)) ?? 0) + 1);
    }
    if (undeclared.size === 0) return result;
    const issues = result.issues.filter((issue) => {
        const prefix = issue.code === "mce-unbound-prefix" ? UNBOUND_IGNORABLE_MESSAGE.exec(issue.message)?.[1] : undefined;
        const remaining = prefix === undefined ? 0 : (undeclared.get(key(issue, prefix)) ?? 0);
        if (remaining === 0) return true;
        undeclared.set(key(issue, prefix!), remaining - 1);
        return false;
    });
    return finalize(issues);
}

function removeMcAlternateContent(dom: Document): void {
    const list = getElementsByTagNameNSAll(dom, MC_NAMESPACE, "AlternateContent");
    const items: Element[] = [];
    for (let i = 0; i < list.length; i += 1) {
        const e = list[i];
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
        const attr = attrs[i];
        if (!attr) continue;
        const local = localName(attr.name).toLowerCase();
        if (local === attrName) return attr.value;
    }
    return null;
}

function finalize(issues: ValidationIssue[]): ValidationResult {
    return { valid: issues.every((i) => i.severity !== "error"), issues };
}

/** Return true if `node` has at least one Element-type child node. */
export function hasAnyElementChild(node: Node): boolean {
    for (let i = 0; i < node.childNodes.length; i += 1) {
        const child = node.childNodes[i];
        if (child && child.nodeType === 1 /* ELEMENT_NODE */) return true;
    }
    return false;
}
