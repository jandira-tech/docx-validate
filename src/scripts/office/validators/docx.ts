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
 * Validator for Word documents.
 *
 * The default (`word-valid`) profile judges a package by what Microsoft Word
 * reads — the structural checks below plus the registered Word rules — and
 * compiles no schema; the `lenient` and `strict` profiles also walk the XSD
 * schemas.
 *
 * 1:1 TypeScript port of `src/docx-validate/scripts/office/validators/docx.py`
 * (task #10). Subclasses `BaseSchemaValidator` and adds DOCX-specific checks:
 *
 *   - paragraph counts (original vs. modified, excluding text-box overlays);
 *   - whitespace preservation on `<w:t>` runs that lead/trail with whitespace;
 *   - tracked-changes nesting (no `<w:t>` inside `<w:del>`, no `<w:delText>`
 *     inside `<w:ins>`);
 *   - comment-marker pairing (`commentRangeStart`/`End`/`commentReference`);
 *   - id constraints (0 < paraId/textId < 0x80000000, 0 < durableId < 0x7FFFFFFF
 *     — plain 32-bit numbers, well inside Number.MAX_SAFE_INTEGER, no BigInt);
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

import { constants as fsConstants, promises as fs, readFileSync } from "node:fs";
import path from "node:path";
import { canonicalPartName, isRelationshipsFolderName, isRelationshipsPartName, resolvePartName } from "../../../lib/content-types";
import { opcNameIssueSeverity } from "../../../lib/mode-policy";
import { nextSecureLongHexNumber } from "../../../lib/secure-id";
import type { PackageFormat, ValidationIssue, ValidationResult } from "../../../lib/types";
import { withoutConsequentialParseFailures } from "../../../lib/parse-failures";
import { mergeResults, NS } from "../../../lib/types";
import { promoteFullCheckIssue } from "../../../lib/mode-policy";
import { propertiesParts } from "../../../lib/properties-parts";
import {
    COMMENTS_EXTENSIBLE_RELATIONSHIP_TYPE,
    COMMENTS_IDS_RELATIONSHIP_TYPE,
    wordExactTypeTargets,
    wordReferenceLoadedSources,
    wordReadParts,
    wordStoryRootParts,
    wordTypedRelationshipTargets,
    wordUnobservedRelationshipTargets,
    glossaryTypedRelationshipTargets,
    wordMainReachedParts,
} from "../../../lib/word-read-parts";
import {
    hasWordRelationshipReference,
    isWordBlockingIssue,
    unobservedReferenceAttributeValues,
    wordRelationshipReferenceIds,
} from "./word-blocking";
import {
    decodeXmlBytes,
    decodeXmlPart,
    getElementsByTagNameAll,
    getElementsByTagNameNSAll,
    parseXml,
    parseXmlBytes,
} from "../../../lib/xml-helpers";
import { writeXmlPart } from "../../../lib/xml-part-io";
import {
    BaseSchemaValidator,
    collectDeclaredPrefixes,
    containRelationshipTarget,
    packageEntries,
    withoutRestatedIgnorablePrefixes,
    XML_NAMESPACE,
} from "./base";

import { runChecks } from "../../../lib/check-scheduler";
import { normalizePartName, type ValidationContext } from "../../../lib/validation-context";
import { createDefaultCheckRegistry, WORD_RULE_CHECKS } from "../checks/index";
import { WORD_HEADER_FOOTER_RELATIONSHIP } from "../checks/word-header-footer";
import { WORD_PART_REFERENCE_RELATIONSHIP } from "../checks/word-part-reference";
import { WORD_ATTACHED_TOOLBAR_RELATIONSHIP } from "../checks/word-attached-toolbar";
import { type StoryReferenceIdentity, storyReferenceIdentity } from "../story-reference-identity";
import { WORD_PART_CONTENT_TYPE_GENERIC, wordPartContentTypeGenericFindings } from "../checks/word-part-content-type-generic";
import { contextFromDirectory, openDirectoryContext, packageExtractFailedIssue, storedNamesRun } from "../package-context";
import { WORD_2006_NAMESPACE, WORD_PARAGRAPH_NAMESPACES, WORD_STRICT_NAMESPACE } from "../word-namespaces";
import { WORD_PART_CONTENT_TYPE_MISMATCH, wordPartContentTypeMismatchFindings } from "../checks/word-part-content-type-mismatch";
import { WORD_PART_CONTENT_TYPE_INEXACT, wordPartContentTypeInexactFindings } from "../checks/word-part-content-type-inexact";
import { RELATIONSHIPS_UNDERSTOOD_NAMESPACES, understoodNamespacesForPart } from "../../../lib/markup-compatibility";
/**
 * The relationship-type segments whose targets can hold the user-text
 * parts `userTextXmlFiles` and `documentXmlFiles` walk (#698): a reached
 * glossary's story parts, judged against the glossary's own styles part
 * rather than the main document's. A nested `glossaryDocument` stays in
 * the set: reached only through the outer glossary's relationships, its
 * references are that root's, not the main document's.
 */
const STORY_TEXT_SEGMENTS: ReadonlySet<string> = new Set(["comments", "endnotes", "footnotes", "footer", "glossaryDocument", "header"]);

import {
    WORD_PART_CONTENT_TYPE_UNRECOGNIZED,
    wordObservedContentTypeParts,
    wordPartContentTypeUnrecognizedFindings,
} from "../checks/word-part-content-type-unrecognized";

export { WORD_2006_NAMESPACE, WORD_PARAGRAPH_NAMESPACES, WORD_STRICT_NAMESPACE } from "../word-namespaces";
/** OpenXmlPowerTools DocumentBuilder Insert — `PtOpenXml.ptOpenXml` in PtOpenXmlUtil.cs */
const DOCUMENT_BUILDER_INSERT_NAMESPACE = "http://powertools.codeplex.com/documentbuilder/2011/insert";
const DOCUMENT_BUILDER_INSERT_NAMESPACES: ReadonlySet<string> = new Set([DOCUMENT_BUILDER_INSERT_NAMESPACE]);

/** True when `xml` declares the Insert URI as an attribute value (`xmlns` or `xmlns:prefix`). */
function declaresDocumentBuilderInsertNs(xml: string): boolean {
    return xml.includes(`="${DOCUMENT_BUILDER_INSERT_NAMESPACE}"`) || xml.includes(`='${DOCUMENT_BUILDER_INSERT_NAMESPACE}'`);
}

/**
 * Local names of EG_RunInnerContent elements — valid ONLY inside a `<w:r>`,
 * never as a direct child of `<w:p>` (where they are not in EG_PContent and
 * Word drops them on open). Used by `validateRunContentPlacement`. Deliberately
 * excludes elements that ARE valid paragraph-level siblings of runs (w:r,
 * w:hyperlink, w:fldSimple, w:sdt, w:smartTag, w:bookmarkStart/End,
 * w:commentRangeStart/End, w:ins, w:del, w:moveFrom/To, math, …).
 */
const RUN_INNER_CONTENT_LOCALS: ReadonlySet<string> = new Set([
    "t",
    "tab",
    "br",
    "cr",
    "drawing",
    "pict",
    "object",
    "sym",
    "softHyphen",
    "noBreakHyphen",
    "ptab",
    "fldChar",
    "instrText",
    "delText",
    "delInstrText",
    "lastRenderedPageBreak",
    "footnoteReference",
    "endnoteReference",
    "commentReference",
    "pgNum",
    "dayShort",
    "monthShort",
    "yearShort",
    "dayLong",
    "monthLong",
    "yearLong",
    "annotationRef",
    "separator",
    "continuationSeparator",
    "footnoteRef",
    "endnoteRef",
    "ruby",
    "contentPart",
]);
const W14_NAMESPACE = "http://schemas.microsoft.com/office/word/2010/wordml";
const XMLNS_NAMESPACE = "http://www.w3.org/2000/xmlns/";
const W15_NAMESPACE = "http://schemas.microsoft.com/office/word/2012/wordml";
const W16CEX_NAMESPACE = "http://schemas.microsoft.com/office/word/2018/wordml/cex";
const W16CID_NAMESPACE = "http://schemas.microsoft.com/office/word/2016/wordml/cid";
const VML_NAMESPACE = "urn:schemas-microsoft-com:vml";
const WP_NAMESPACE = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
const WP14_DRAWING_NAMESPACE = "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing";
const DRAWINGML_NAMESPACE = "http://schemas.openxmlformats.org/drawingml/2006/main";
const PICTURE_NAMESPACE = "http://schemas.openxmlformats.org/drawingml/2006/picture";
const DEFAULT_PICTURE_EXTENT_EMU = "95250";

const WORD_DRAWING_SCALAR_TEXT_ELEMENTS: ReadonlyArray<{ namespace: string; localName: string }> = [
    { namespace: WP_NAMESPACE, localName: "align" },
    { namespace: WP_NAMESPACE, localName: "posOffset" },
    { namespace: WP14_DRAWING_NAMESPACE, localName: "pctWidth" },
    { namespace: WP14_DRAWING_NAMESPACE, localName: "pctHeight" },
];
const TRACKED_CHANGE_ID_TAGS = ["ins", "del", "rPrChange", "pPrChange", "tblPrChange", "trPrChange", "tcPrChange"] as const;
const TABLE_BORDER_SIDES = ["top", "left", "bottom", "right", "insideH", "insideV"] as const;
const BORDER_COMPARE_ATTRIBUTES = ["val", "sz", "space", "color", "themeColor", "themeTint", "themeShade", "shadow", "frame"] as const;

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
 * Canonical `<w:style>` definitions for well-known styleIds that writers
 * routinely *reference* via `<w:rStyle w:val="…"/>` or `<w:pStyle …/>`
 * without ever *defining*. Drawn from a clean Word save's `styles.xml`
 * for each style. Used by `repairMissingStyleDefinitions` — unknown
 * styleIds are NOT auto-defined (we'd be guessing intent).
 *
 * The fragments are appended verbatim into the `<w:styles>` root, so they
 * must be self-contained and ordered such that any inter-style
 * `basedOn`/`link` references they contain also resolve to either an
 * already-present style OR another fragment we'll append in the same pass.
 */
const WELL_KNOWN_STYLE_DEFINITIONS: Readonly<Record<string, string>> = {
    CommentReference:
        `<w:style w:type="character" w:styleId="CommentReference">` +
        `<w:name w:val="annotation reference"/>` +
        `<w:uiPriority w:val="99"/>` +
        `<w:semiHidden/>` +
        `<w:unhideWhenUsed/>` +
        `<w:rPr><w:sz w:val="16"/><w:szCs w:val="16"/></w:rPr>` +
        `</w:style>`,
    CommentText:
        `<w:style w:type="paragraph" w:styleId="CommentText">` +
        `<w:name w:val="annotation text"/>` +
        `<w:uiPriority w:val="99"/>` +
        `<w:semiHidden/>` +
        `<w:unhideWhenUsed/>` +
        `<w:rPr><w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr>` +
        `</w:style>`,
    CommentSubject:
        `<w:style w:type="paragraph" w:styleId="CommentSubject">` +
        `<w:name w:val="annotation subject"/>` +
        `<w:basedOn w:val="CommentText"/>` +
        `<w:next w:val="CommentText"/>` +
        `<w:uiPriority w:val="99"/>` +
        `<w:semiHidden/>` +
        `<w:unhideWhenUsed/>` +
        `<w:rPr><w:b/><w:bCs/></w:rPr>` +
        `</w:style>`,
    Header:
        `<w:style w:type="paragraph" w:styleId="Header">` +
        `<w:name w:val="header"/>` +
        `<w:uiPriority w:val="99"/>` +
        `<w:unhideWhenUsed/>` +
        `</w:style>`,
    Footer:
        `<w:style w:type="paragraph" w:styleId="Footer">` +
        `<w:name w:val="footer"/>` +
        `<w:uiPriority w:val="99"/>` +
        `<w:unhideWhenUsed/>` +
        `</w:style>`,
    DefaultParagraphFont:
        `<w:style w:type="character" w:default="1" w:styleId="DefaultParagraphFont">` +
        `<w:name w:val="Default Paragraph Font"/>` +
        `<w:uiPriority w:val="1"/>` +
        `<w:semiHidden/>` +
        `<w:unhideWhenUsed/>` +
        `</w:style>`,
    // The four "implied default" styleIds per ECMA-376 §17.7.4.4. Every
    // WordprocessingML document must define a default for each of the four
    // style types (paragraph, character, table, numbering) — Word silently
    // looks these up by name when an element omits an explicit style and
    // pops "this document needs to be repaired" with a "Table Properties"
    // category if the lookup misses.
    Normal: `<w:style w:type="paragraph" w:default="1" w:styleId="Normal">` + `<w:name w:val="Normal"/>` + `<w:qFormat/>` + `</w:style>`,
    TableNormal:
        `<w:style w:type="table" w:default="1" w:styleId="TableNormal">` +
        `<w:name w:val="Normal Table"/>` +
        `<w:uiPriority w:val="99"/>` +
        `<w:semiHidden/>` +
        `<w:unhideWhenUsed/>` +
        `<w:tblPr><w:tblInd w:w="0" w:type="dxa"/>` +
        `<w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:left w:w="108" w:type="dxa"/>` +
        `<w:bottom w:w="0" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar></w:tblPr>` +
        `</w:style>`,
    NoList:
        `<w:style w:type="numbering" w:default="1" w:styleId="NoList">` +
        `<w:name w:val="No List"/>` +
        `<w:uiPriority w:val="99"/>` +
        `<w:semiHidden/>` +
        `<w:unhideWhenUsed/>` +
        `</w:style>`,
};

/**
 * The four "implied default" styleIds per ECMA-376 §17.7.4.4 — one for
 * each of the four `<w:style>` types. Word looks these up by name when
 * an element omits an explicit style; missing defaults trip the
 * "Document Recovery" dialog with a "Table Properties" complaint
 * (because the implicit table-style lookup is the most visible failure).
 */
const REQUIRED_DEFAULT_STYLES: ReadonlyArray<{ styleId: string; type: "paragraph" | "character" | "table" | "numbering" }> = [
    { styleId: "Normal", type: "paragraph" },
    { styleId: "DefaultParagraphFont", type: "character" },
    { styleId: "TableNormal", type: "table" },
    { styleId: "NoList", type: "numbering" },
];

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
const TRACKING_TOKEN_REGEX = /\[\[DOCX_(?:INS|DEL|CMT)_(?:START|END):[^\]]*?\]\]|\[\[DOCX_PMARK_(?:DEL|INS):[^\]]*?\]\]/g;

const MAX_PARA_ID = 0x80000000;
const MAX_DURABLE_ID = 0x7fffffff;

export interface ParagraphCounts {
    original: number;
    modified: number;
    delta: number;
    originalUsesStrictNamespace: boolean;
}

/**
 * Single-pass DFS collection for `.//w:del//w:t`-style queries: tracks
 * nesting context instead of checking `.parentNode` for every element,
 * eliminating O(N*D) worst-case scaling. Prunes drawings inside deletions.
 */
const collectDeletedRunText = (root: Document | Element, localName: string): Element[] => {
    const out: Element[] = [];
    const stack: { node: Element; inDel: boolean }[] = [];

    const rootElem = root.nodeType === 9 ? (root as Document).documentElement : (root as Element);
    if (rootElem) {
        stack.push({ node: rootElem, inDel: false });
    }

    while (stack.length > 0) {
        const { node, inDel } = stack.pop()!;
        const ns = node.namespaceURI;
        const name = node.localName;

        let nextInDel = inDel;
        if (name === "del" && (ns === WORD_2006_NAMESPACE || ns === WORD_STRICT_NAMESPACE)) {
            nextInDel = true;
        } else if (inDel && (name === "drawing" || name === "pict")) {
            // Prune drawings inside deletions
            continue;
        }

        if (name === localName && (ns === WORD_2006_NAMESPACE || ns === WORD_STRICT_NAMESPACE)) {
            if (nextInDel) {
                out.push(node);
            }
        }

        let child = node.lastChild;
        while (child) {
            if (child.nodeType === 1) {
                stack.push({ node: child as Element, inDel: nextInDel });
            }
            child = child.previousSibling;
        }
    }

    return out;
};

export class DOCXSchemaValidator extends BaseSchemaValidator {
    protected override get duplicateContentTypeSeverity(): "error" | "warning" {
        return this.profile === "word-valid" ? "error" : super.duplicateContentTypeSeverity;
    }
    /**
     * Element (lower-cased local name) → required `r:id` relationship-type suffix.
     * Enables the `rels-id-mismatch` check in `validateAllRelationshipIds`: an
     * `r:id` that RESOLVES in the sidecar but to the WRONG kind of part. The XSD
     * cannot catch this (rId target semantics are not schema-constrained), yet
     * real Microsoft Word REFUSES to open such a file ("Word found unreadable
     * content"). Proven word-first: a redline whose inserted `<w:hyperlink r:id>`
     * collided with the destination's `header`/`footer`/`fontTable`/`theme`
     * relationships passed XSD validation with 0 errors but Word rejected it.
     * Only `r:id`-bearing elements are listed (image `<a:blip>` uses `r:embed`,
     * which the existence check already covers).
     */
    protected readonly elementRelationshipTypes: Record<string, string> = {
        // Unambiguous `r:id` invariants only. `<o:OLEObject>` is deliberately
        // omitted: its r:id legitimately targets EITHER an `oleObject` OR a
        // `package` rel (embedded packages), so enforcing one type would false-
        // positive. `<a:blip>` uses `r:embed`, not `r:id`, so it is unaffected.
        hyperlink: "hyperlink",
        // headerReference / footerReference are derived by
        // BaseSchemaValidator._getExpectedRelationshipType (`*Reference` → prefix).
    };

    /**
     * The package as the named checks see it, built once per `validate()` run
     * so its read and parse caches never outlive the run (auto-repair rewrites
     * parts between runs). Outside a run, a check method called directly
     * reads the package as it is now.
     */
    private runContext: ValidationContext | undefined;
    private inRun = false;

    private context(): ValidationContext {
        if (!this.inRun) return contextFromDirectory(this.unpackedDir);
        this.runContext ??= contextFromDirectory(this.unpackedDir);
        return this.runContext;
    }

    /**
     * The main document's part reached by the relationship segment
     * `segment` — Transitional or Strict — so a reached glossary's own
     * part is never judged for the main document, wherever it is stored
     * (#688, #690). `undefined` when such a relationship names a stored
     * part that is not one of the XML parts: nothing is judged, and no
     * other part stands in for it (#693). Only when no relationship names
     * a stored part, or the package view cannot be read, the last file
     * named `fallbackName` in the main document's folder (#699); `null`
     * when there is none. A part stored elsewhere that no relationship
     * names is not judged.
     */
    private async mainRelatedPath(segment: string, fallbackName: string): Promise<string | null | undefined> {
        try {
            const targets = await wordTypedRelationshipTargets(this.context(), new Set([segment]), this.profile);
            if (targets.length > 0) return this.xmlFileOf(targets.map((target) => target.part));
        } catch {
            // An unreadable package view is reported by its own checks.
        }
        const folder = await this.mainDocumentFolder();
        let byName: string | null = null;
        for (const xmlFile of this.xmlFiles) {
            if (
                baseName(xmlFile) === fallbackName &&
                canonicalPartName(path.posix.dirname(normalizePartName(this.entryName(xmlFile)))) === folder
            )
                byName = xmlFile;
        }
        return byName;
    }

    /**
     * The folder the by-name fallback reads (#699): the folder of the
     * first main document `wordTypedRelationshipTargets` names, or `word`
     * when the package relationships name none or cannot be read. It reads
     * the profile's relationship view, so under lenient an `officeDocument`
     * relationship only a markup-compatibility branch carries names no
     * folder and `word` is read.
     */
    private async mainDocumentFolder(): Promise<string> {
        try {
            const [first] = await wordTypedRelationshipTargets(this.context(), new Set(["officeDocument"]), this.profile);
            if (first !== undefined) return canonicalPartName(path.posix.dirname(normalizePartName(first.part)));
        } catch {
            // An unreadable package view is reported by its own checks.
        }
        return "word";
    }

    /** The first of `parts` that names an XML part by OPC equivalence, or `undefined`. */
    private xmlFileOf(parts: readonly string[]): string | undefined {
        for (const part of parts) {
            const file =
                this.xmlFiles.find((xmlFile) => normalizePartName(this.entryName(xmlFile)) === normalizePartName(part)) ??
                this.xmlFiles.find(
                    (xmlFile) =>
                        canonicalPartName(normalizePartName(this.entryName(xmlFile))) === canonicalPartName(normalizePartName(part)),
                );
            if (file !== undefined) return file;
        }
        return undefined;
    }

    private async mainStylesPath(): Promise<string | null | undefined> {
        return this.mainRelatedPath("styles", "styles.xml");
    }

    /**
     * The comment-thread part the main document's relationship of exactly
     * `type` names (#690), as `wordExactTypeTargets` reads it: the Microsoft
     * commentsIds/commentsExtensible types carry no Office relationship-type
     * segment, so `mainRelatedPath` cannot reach them. When such a
     * relationship names a stored part that is not one of the XML parts,
     * `undefined` (#693); only when none names a stored part, the file
     * named `name` in the chosen comments part's folder, or `null`.
     */
    private async mainThreadPath(type: string, commentsXml: string, name: string): Promise<string | null | undefined> {
        try {
            const parts = await wordExactTypeTargets(this.context(), type, this.profile);
            if (parts.length > 0) return this.xmlFileOf(parts);
        } catch {
            // An unreadable package view is reported by its own checks.
        }
        return this.sameFolderByName(commentsXml, name);
    }

    /**
     * The last file named `name` stored in the same folder as `partPath`:
     * the by-name fallback of `repairCommentThreading`, which never pairs a
     * commentsIds/commentsExtensible part with comments from another story
     * root (#690). `null` when that folder holds none.
     */
    private sameFolderByName(partPath: string, name: string): string | null {
        const folder = canonicalPartName(path.posix.dirname(normalizePartName(this.entryName(partPath))));
        let byName: string | null = null;
        for (const xmlFile of this.xmlFiles) {
            if (baseName(xmlFile) === name && canonicalPartName(path.posix.dirname(normalizePartName(this.entryName(xmlFile)))) === folder)
                byName = xmlFile;
        }
        return byName;
    }

    /**
     * The `.rels` part as the non-lenient walks read it (#635): the run
     * context's `effectiveXml` whenever `understoodNamespacesForPart`
     * processes `part` with the same understood set `relationshipElements`
     * applies — exactly the relationships parts — so markup compatibility
     * runs once per part per run (#642). A part the context cannot supply
     * (an unrecognized spelling, a part absent from the package view, or a
     * refused directory) keeps the validator's own processed clone,
     * memoized only within a run (#654). DOCX/DOCM only: PPTX keeps the
     * stored view (#643).
     */
    protected override relationshipsView(document: Document, part: string): Document {
        if (understoodNamespacesForPart(part) === RELATIONSHIPS_UNDERSTOOD_NAMESPACES) {
            try {
                const effective = this.context().effectiveXml(part);
                if (effective?.ok === true) {
                    return effective.document;
                }
            } catch {
                // A refused directory view falls back to the processed clone.
            }
        }
        // Outside a run a check method called directly reads the package as
        // it is now, as `context()` does: a repair such as `repairIgnorable`
        // can change which branch markup compatibility selects (#654).
        if (!this.inRun) this.resetProcessedRelationships();
        return this.processedRelationshipsView(document, part);
    }

    // ----- top-level entry point ----------------------------------------------

    /**
     * Run every check. `compareParagraphCounts` is informational only and
     * does not affect the merged validity.
     */
    async validate(): Promise<ValidationResult> {
        this.inRun = true;
        this.runContext = undefined;
        // A run re-reads the package: the processed relationship views of
        // the previous run must not outlive it, as `runContext` does not.
        this.resetProcessedRelationships();
        try {
            return this.withDeadlineFinding(await this.withinBaselineRun(() => this.runEveryCheck()));
        } finally {
            this.inRun = false;
            this.runContext = undefined;
        }
    }

    private async runEveryCheck(): Promise<ValidationResult> {
        this.xsdWalk = { started: false, parts: [], eligible: 0 };
        this.malformedParts = new Set();
        this.resetExecutedChecks();
        const opened = openDirectoryContext(this.unpackedDir);
        if (opened.openError !== undefined) {
            const diagnostic = await storedNamesRun(opened.context, "docx", this.profile, this.deadline);
            this.recordExecutedChecks(diagnostic.executed);
            return { valid: false, issues: [...diagnostic.issues, packageExtractFailedIssue(this.unpackedDir, opened.openError)] };
        }
        this.runContext = opened.context;
        const rule = this.ruleScheduler();
        const xmlOk = await rule("DOCX validateXml", () => this.validateXml());
        // Only lenient stops at a malformed part. Word-valid and strict keep
        // running the checks of every other part, so strict reports every
        // code word-valid does. The malformed part itself is reported once,
        // by its XML diagnostics.
        if (!xmlOk.valid && this.profile === "lenient") return xmlOk;
        this.malformedParts = new Set(
            xmlOk.issues.flatMap((issue) => (issue.severity === "error" && issue.path !== undefined ? [issue.path] : [])),
        );
        const settle = async (pending: Array<Promise<ValidationResult>>): Promise<ValidationResult[]> =>
            (await Promise.all(pending)).map((result) => ({
                ...result,
                issues: withoutConsequentialParseFailures(result.issues, this.malformedParts),
            }));

        const contentTypes = rule("DOCX validateContentTypes", () => this.validateContentTypes());
        // #612: the relationship-id walk is captured by name because a full
        // non-lenient run hands the story reference elements its Word rule
        // already reported back to that rule, dropping rels-id-mismatch's —
        // and, for an r:id naming no Relationship, rels-id-missing's (#619) —
        // copy of each (withoutRepeatedHeaderFooterFindings). Called on its
        // own, validateAllRelationshipIds keeps every finding.
        const relationshipIds = rule("DOCX validateAllRelationshipIds", () => this.validateAllRelationshipIds());
        const checks: Array<Promise<ValidationResult>> = [
            rule("DOCX validateNamespaces", () => this.validateNamespaces()),
            rule("DOCX validateUniqueIds", () => this.validateUniqueIds()),
            rule("DOCX validateFileReferences", () => this.validateFileReferences()),
            rule("DOCX validateRelationshipElements", () => this.validateRelationshipElements()),
            contentTypes,
            rule("DOCX validateOverridePartNames", () => this.validateOverridePartNames(this.context(), "docx", this.malformedParts)),
            rule("DOCX validateContentTypesMarkupCompatibility", () =>
                this.validateContentTypesMarkupCompatibility(this.context(), "docx", this.malformedParts),
            ),
            rule("DOCX validateWordPartContentTypes", () =>
                this.validateWordPartContentTypes(this.context(), "docx", this.malformedParts, contentTypes),
            ),
            rule("DOCX validateWhitespacePreservation", () => this.validateWhitespacePreservation()),
            rule("DOCX validateDeletions", () => this.validateDeletions()),
            rule("DOCX validateInsertions", () => this.validateInsertions()),
            relationshipIds,
            rule("DOCX validateIdConstraints", () => this.validateIdConstraints()),
            rule("DOCX validateCommentMarkers", () => this.validateCommentMarkers()),
            rule("DOCX validateCommentThreading", () => this.validateCommentThreading()),
            rule("DOCX validateNoTrackingTokens", () => this.validateNoTrackingTokens()),
            rule("DOCX validateAllParagraphsHaveParaId", () => this.validateAllParagraphsHaveParaId()),
            rule("DOCX validateRunContentPlacement", () => this.validateRunContentPlacement()),
            rule("DOCX validateStyleReferences", () => this.validateStyleReferences()),
            rule("DOCX validateStyleDefaults", () => this.validateStyleDefaults()),
            rule("DOCX validateDocumentBuilderInserts", () => this.validateDocumentBuilderInserts()),
            rule("DOCX validateDocPropsBooleans", () => this.validateDocPropsBooleans()),
            rule("DOCX validateNoBom", () => this.validateNoBom()),
            rule("DOCX validateNoEmptyRelsParts", () => this.validateNoEmptyRelsParts()),
            rule("DOCX validateOrphanedRelationships", () => this.validateOrphanedRelationships()),
            rule("DOCX validateFontTable", () => this.validateFontTable()),
            rule("DOCX validateLatentStyles", () => this.validateLatentStyles()),
            rule("DOCX validateTableLook", () => this.validateTableLook()),
            rule("DOCX validateTrackedChangeIds", () => this.validateTrackedChangeIds()),
            rule("DOCX validateRedundantCellBorders", () => this.validateRedundantCellBorders()),
            rule("DOCX validateRelationshipIdStability", () => this.validateRelationshipIdStability()),
            rule("DOCX validateRedundantRunProperties", () => this.validateRedundantRunProperties()),
        ];
        // Strict runs the same Word-proven rules at their native severity.
        const wordRules = this.profile === "lenient" ? undefined : rule("DOCX Word rules", () => this.validateRegisteredWordRules());
        if (wordRules) checks.push(wordRules);
        // #612: in a full non-lenient run each element the Word rule reported
        // yields one finding, the Word rule's, so the settled relationship-id
        // result drops its copy. Lenient runs no Word rule and keeps every
        // mismatch.
        const handOffStoryReferences = (results: ValidationResult[]): ValidationResult[] => {
            if (wordRules === undefined) return results;
            const rules = results[checks.indexOf(wordRules)]!;
            const rels = results[checks.indexOf(relationshipIds)]!;
            return results.map((result) => (result === rels ? withoutRepeatedHeaderFooterFindings(rels, rules) : result));
        };
        if (this.profile === "word-valid") {
            // Word-valid judges what Word reads through structural checks
            // alone: no schema is compiled or walked.
            return this.applyWordValidProfile(
                withoutRestatedIgnorablePrefixes(mergeResults(xmlOk, ...handOffStoryReferences(await settle(checks)))),
            );
        }

        const xsd = rule("DOCX validateAgainstXsd", () => this.validateAgainstXsd());
        checks.push(xsd);
        const results = await settle(checks);
        if (!wordRules) return withoutRestatedIgnorablePrefixes(mergeResults(xmlOk, ...results));
        const rules = results[checks.indexOf(wordRules)]!;
        const schema = results[checks.indexOf(xsd)]!;
        return withoutRestatedIgnorablePrefixes(
            mergeResults(
                xmlOk,
                ...handOffStoryReferences(results).map((result) =>
                    result === schema ? withoutRepeatedCompatibilityFindings(schema, rules) : result,
                ),
            ),
        );
    }

    /**
     * Word rules registered in `WORD_RULE_CHECKS`, run through the shared
     * context. Lenient output stays as it was, so they run under `word-valid`
     * and `strict` only.
     */
    private async validateRegisteredWordRules(): Promise<ValidationResult> {
        if (WORD_RULE_CHECKS.length === 0) return { valid: true, issues: [] };
        const rules = new Set(WORD_RULE_CHECKS.map((check) => check.id));
        const run = await runChecks(this.context(), createDefaultCheckRegistry(), {
            format: "docx",
            select: [...rules],
            profile: this.profile,
            deadline: this.deadline,
        });
        // xml-syntax runs as the rules' prerequisite; validateXml already
        // reports its findings, so keep only the rules' own (#307).
        const issues = run.outcomes.filter((outcome) => rules.has(outcome.id)).flatMap((outcome) => outcome.issues);
        // Disclose the rules that really ran: a rule whose prerequisite
        // failed was skipped, and a skipped rule did not run (#319).
        this.recordExecutedChecks(run.executed);
        return { valid: issues.every((issue) => issue.severity !== "error"), issues };
    }

    /**
     * The `word-part-content-type-generic` rule (#432) and the
     * `word-part-content-type-inexact` rule (#540, widened by #550) over the shared
     * context, at the severity the profile gives them: an error under
     * word-valid (Word 16.115 offers to recover the package) and under
     * strict (ECMA-376 Part 1, clause 11.3, gives every WordprocessingML
     * part its own content type, so strict also judges the Part 1 types a
     * check's probes have not observed; #506), a warning under lenient.
     * Sessions 20 and 21 moved numbering and theme into the generic check's
     * observed set (#571), and numbering, theme, footer, endnotes, footnotes
     * and comments into the inexact check's (#570); sessions 21 and 22 moved
     * a reached glossary's styles part and referenced header into the
     * inexact check's (#579). Session 23 adds the glossary's generic theme
     * (#595) and referenced footer, footnotes, endnotes and comments with
     * inexact types (#596). Session 24 adds generic glossary styles (#601)
     * and read header/footer content-type swaps (#600). Strict mismatch
     * findings cover every mapped related part. Session 46 adds a part typed
     * by an unrecognised media type — neither its own type, nor generic XML,
     * nor another mapped Word type — which word-valid blocks only on the
     * observed reach, the main document's `styles` relationship (#707).
     * Strict's inexact findings cite clause 11.3
     * with no recovery claim for a lazy part `wordReadParts` does not
     * include — session 21 (s21b) observed those open clean (#589). A
     * Default-typed main or theme finding is dropped only when strict
     * actually reports ct-undeclared-part for that same part (#535, #571);
     * a part reached by several relationship types counts as main or theme
     * when any of its reaching segments says so, whichever segment the
     * finding names (#593).
     */
    private async validateWordPartContentTypes(
        context: ValidationContext,
        format: PackageFormat,
        malformed: ReadonlySet<string>,
        contentTypes: Promise<ValidationResult>,
    ): Promise<ValidationResult> {
        const view = malformed.size > 0 ? context.withUnparseable(new Map([...malformed].map((part) => [part, "malformed"]))) : context;
        // Strict judges the Part 1 WordprocessingML part types beyond the
        // observed sets: for the generic check header, footer, endnotes,
        // footnotes, comments and glossaryDocument (#506, #571), for the
        // inexact check glossaryDocument only (#570). Word-valid and lenient
        // keep the observed sets; under strict an unread lazy part keeps the
        // clause-11.3 wording instead of the Word sentence (#589).
        const found = await wordPartContentTypeGenericFindings(view, this.profile === "strict" ? "spec" : "observed", this.profile);
        const inexact = await wordPartContentTypeInexactFindings(view, this.profile === "strict" ? "spec" : "observed", this.profile);
        const mismatch = await wordPartContentTypeMismatchFindings(view, this.profile === "strict" ? "spec" : "observed", this.profile);
        const unrecognized = await wordPartContentTypeUnrecognizedFindings(view, this.profile);
        this.recordExecutedCheck(WORD_PART_CONTENT_TYPE_GENERIC);
        this.recordExecutedCheck(WORD_PART_CONTENT_TYPE_INEXACT);
        this.recordExecutedCheck(WORD_PART_CONTENT_TYPE_MISMATCH);
        this.recordExecutedCheck(WORD_PART_CONTENT_TYPE_UNRECOGNIZED);
        // The base pass skips non-XML names and docProps/. Suppress a
        // Default-typed main or theme only if its own path has a replacement error;
        // otherwise strict must retain the Word-blocking finding (#535). The
        // part counts as main or theme when any reaching segment says so —
        // the walk keeps every segment, so an earlier unrelated relationship
        // cannot make the duplicate return (#593).
        const replacements = this.profile === "strict" ? (await contentTypes).issues : [];
        const judged = [...found, ...inexact].filter(
            (finding) =>
                !(
                    (finding.segments.includes("officeDocument") || finding.segments.includes("theme")) &&
                    finding.source === "default" &&
                    replacements.some(
                        (issue) => issue.code === "ct-undeclared-part" && issue.severity === "error" && issue.path === finding.issue.path,
                    )
                ),
        );
        const severity = opcNameIssueSeverity(this.profile, format);
        return finalize(
            [...judged, ...mismatch, ...unrecognized].map((finding): ValidationIssue =>
                severity === "warning" ? { ...finding.issue, severity: "warning" } : finding.issue,
            ),
        );
    }

    /**
     * The `word-valid` view of a full run: every error that does not predict a
     * Word refusal or recovery offer ({@link isWordBlockingIssue}) becomes a
     * warning, judged against this package's properties parts, Word-read parts
     * and story roots. Runs after the story-reference hand-off, so the
     * pairing in `withoutRepeatedHeaderFooterFindings` sees the producers'
     * own issue objects.
     */
    private async applyWordValidProfile(result: ValidationResult): Promise<ValidationResult> {
        const context = this.context();
        // Resolve each part set once per run and share the blocking predicate
        // with partial validation (#381, #615, #617, #718).
        const [properties, read, storyRoots, observedContentTypes] = await Promise.all([
            propertiesParts(context, this.profile),
            wordReadParts(context, this.profile),
            wordStoryRootParts(context, this.profile),
            wordObservedContentTypeParts(context, this.profile),
        ]);
        const missingSidecarPart = (issue: ValidationIssue): string | undefined =>
            issue.severity === "error" && issue.code === "rels-missing-sidecar" ? issue.path : undefined;
        // Only a missing-sidecar error consults the unobserved-type targets.
        const unobserved = result.issues.some((issue) => missingSidecarPart(issue) !== undefined)
            ? await wordUnobservedRelationshipTargets(context, this.profile)
            : new Set<string>();
        // #645, session 36 s36b/s36d: a part every reaching relationship names
        // through the chart type is loaded by Word only when a Word-read
        // part's effective XML references one of those relationship Ids;
        // #647, session 44: diagramData, oleObject and hdphoto match every
        // attribute naming the Id but a picture's `a:blip r:embed`, which
        // session 45 opened clean (#702 S1; the hdphoto reach again in
        // session 52 D1); an image relationship is never
        // a source, so an image reach alone warns (#706) — session 52's E1
        // and E2 rows opened the same no-sidecar chart clean when a
        // `w:name r:id` names the reach and when a VML `v:imagedata r:id`
        // inside a selected `mc:Fallback` does.
        const referenceLoadedReach = new Map<string, "referenced" | "unreferenced">();
        for (const path of result.issues.flatMap((issue) => {
            const part = missingSidecarPart(issue);
            return part === undefined ? [] : [part];
        })) {
            const sources = await wordReferenceLoadedSources(context, path, this.profile);
            if (sources === undefined) {
                continue;
            }
            const referenced = sources.some(({ source, id, reference }) => {
                if (source === undefined || !read.has(canonicalPartName(source))) {
                    return false;
                }
                const xml = context.effectiveXml(source);
                return (
                    xml?.ok === true &&
                    (reference === "declared"
                        ? wordRelationshipReferenceIds(xml.document)
                        : unobservedReferenceAttributeValues(xml.document)
                    ).has(id)
                );
            });
            referenceLoadedReach.set(path, referenced ? "referenced" : "unreferenced");
        }
        const issues = result.issues.map((issue): ValidationIssue => {
            if (issue.severity !== "error") return issue;
            // #638, session 35, chunks s35a/s35b: only a conservatively
            // counted reference surviving Word's MC view makes this finding
            // blocking. Keep the issue's identity and serialized shape intact.
            const part = missingSidecarPart(issue);
            const effective = part === undefined ? undefined : context.effectiveXml(part);
            const facts = {
                hasEffectiveRelationshipReference: effective?.ok === true && hasWordRelationshipReference(effective.document),
                reachedByUnobservedRelationship: part !== undefined && unobserved.has(canonicalPartName(part.replace(/\\/g, "/"))),
                referenceLoadedReach: part === undefined ? undefined : referenceLoadedReach.get(part),
            };
            if (isWordBlockingIssue(issue, properties, read, storyRoots, observedContentTypes, facts)) return issue;
            return { ...issue, severity: "warning" };
        });
        return finalize(issues);
    }

    // ----- whitespace ---------------------------------------------------------

    async validateWhitespacePreservation(): Promise<ValidationResult> {
        const issues: ValidationIssue[] = [];
        for (const xmlFile of this.documentXmlFiles()) {
            let dom: Document;
            try {
                dom = await this.readPartDom(xmlFile);
            } catch (err) {
                issues.push({
                    severity: "error",
                    message: `Error: ${err instanceof Error ? err.message : String(err)}`,
                    path: this.relPath(xmlFile),
                    code: "ws-parse",
                });
                continue;
            }
            const tElems = getElementsByTagNameNSAll(dom, WORD_2006_NAMESPACE, "t");
            for (let i = 0; i < tElems.length; i += 1) {
                const elem = tElems[i];
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
        for (const xmlFile of this.documentXmlFiles()) {
            let dom: Document;
            try {
                dom = await this.readPartDom(xmlFile);
            } catch (err) {
                issues.push({
                    severity: "error",
                    message: `Error: ${err instanceof Error ? err.message : String(err)}`,
                    path: this.relPath(xmlFile),
                    code: "del-parse",
                });
                continue;
            }
            const rel = this.relPath(xmlFile);
            // Native DOM walk instead of `.//w:del//w:t` XPath (Jules Bolt). Skip
            // <w:t>/<w:instrText> whose ancestor drawing/pict is itself inside a
            // <w:del> — Word treats that shape as one opaque deleted object (#50).
            for (const t of collectDeletedRunText(dom, "t")) {
                issues.push({
                    severity: "error",
                    message: `<w:t> found within <w:del>: ${previewRepr(t.firstChild?.nodeValue ?? "", 50)}`,
                    path: rel,
                    code: "del-contains-t",
                });
            }
            for (const instr of collectDeletedRunText(dom, "instrText")) {
                issues.push({
                    severity: "error",
                    message: `<w:instrText> found within <w:del> (use <w:delInstrText>): ${previewRepr(instr.firstChild?.nodeValue ?? "", 50)}`,
                    path: rel,
                    code: "del-contains-instrtext",
                });
            }
        }
        return finalize(issues);
    }

    async validateInsertions(): Promise<ValidationResult> {
        const issues: ValidationIssue[] = [];

        for (const xmlFile of this.documentXmlFiles()) {
            let dom: Document;
            try {
                dom = await this.readPartDom(xmlFile);
            } catch (err) {
                issues.push({
                    severity: "error",
                    message: `Error: ${err instanceof Error ? err.message : String(err)}`,
                    path: this.relPath(xmlFile),
                    code: "ins-parse",
                });
                continue;
            }

            const invalid: Element[] = [];
            const stack: { node: Element; inDel: boolean; inIns: boolean }[] = [];

            if (dom.documentElement) {
                stack.push({ node: dom.documentElement, inDel: false, inIns: false });
            }

            while (stack.length > 0) {
                const { node, inDel, inIns } = stack.pop()!;
                const ns = node.namespaceURI;
                const name = node.localName;

                let nextInDel = inDel;
                let nextInIns = inIns;

                if (name === "del" && (ns === WORD_2006_NAMESPACE || ns === WORD_STRICT_NAMESPACE)) {
                    nextInDel = true;
                } else if (name === "ins" && (ns === WORD_2006_NAMESPACE || ns === WORD_STRICT_NAMESPACE)) {
                    nextInIns = true;
                }

                if (name === "delText" && (ns === WORD_2006_NAMESPACE || ns === WORD_STRICT_NAMESPACE)) {
                    if (inIns && !nextInDel) {
                        invalid.push(node);
                    }
                }

                let child = node.lastChild;
                while (child) {
                    if (child.nodeType === 1) {
                        stack.push({ node: child as Element, inDel: nextInDel, inIns: nextInIns });
                    }
                    child = child.previousSibling;
                }
            }

            for (const elem of invalid) {
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
        const run = await runChecks(this.context(), createDefaultCheckRegistry(), {
            format: "docx",
            select: ["comment-markers"],
            profile: this.profile,
            deadline: this.deadline,
        });
        this.recordExecutedChecks(run.executed);
        return finalize([...run.issues]);
    }

    // ----- style reference integrity ------------------------------------------

    /**
     * Every `w:val` on a style-reference element (`w:pStyle`, `w:rStyle`,
     * `w:tblStyle`, `w:numStyle`, `w:tblpPr/w:tblStyle`, `w:linkedStyle`)
     * in document/header/footer/footnote/comment XML must resolve to a
     * `<w:style w:styleId="...">` defined in the styles part the story's
     * root relates — the main document's (by name in its folder when
     * none does, #688, #699) or a reached glossary's own (#698). Word
     * silently substitutes the default style for missing references but
     * pops a "this document needs to be repaired" dialog when opening
     * unfamiliar dangling references.
     *
     * Always reported as `error` (real spec violation under both
     * profiles — there's no "lenient" interpretation of a dangling style).
     *
     * Surfaced by the comparison against Word's own save output for the
     * sample-document fixture: jubarte's writer emits
     * `<w:rStyle w:val="CommentReference"/>` around comment anchors but
     * never adds a `CommentReference` style definition to `styles.xml`.
     * Word's repaired version injects the canonical definition.
     */
    /**
     * Verify the styles part the main document relates (by name in its
     * folder when none does, #688, #699) defines the four "implied default" styles
     * required by ECMA-376 §17.7.4.4 — `Normal` (paragraph), `DefaultParagraphFont`
     * (character), `TableNormal` (table), `NoList` (numbering). Word
     * looks these up by name when an element omits an explicit style and
     * pops "Document Recovery" with a "Table Properties" complaint if
     * `TableNormal` (the most visible default) is missing — even when no
     * `<w:tblStyle>` references it.
     *
     * Profile-aware: strict reports `error`, lenient downgrades to
     * `warning` (many real-world templates omit one or more of the
     * implied defaults and Word still opens them — the dialog only
     * fires for `TableNormal`-on-tables, not for absence in general).
     * The repair pass `repairMissingStyleDefinitions` injects canonical
     * definitions for any missing default, so `--auto-repair` resolves
     * this without manual fixes regardless of profile.
     */
    async validateStyleDefaults(): Promise<ValidationResult> {
        const issues: ValidationIssue[] = [];
        const severity: "error" | "warning" = this.profile === "strict" ? "error" : "warning";
        const stylesPath = await this.mainStylesPath();
        if (!stylesPath) return { valid: true, issues: [] };

        let dom: Document;
        try {
            dom = await this.readPartDom(stylesPath);
        } catch {
            return { valid: true, issues: [] };
        }
        const defined = new Set<string>();
        for (const ns of WORD_PARAGRAPH_NAMESPACES) {
            const list = getElementsByTagNameNSAll(dom, ns, "style");
            for (let i = 0; i < list.length; i += 1) {
                const elem = list[i];
                if (!elem) continue;
                const id = elem.getAttributeNS(ns, "styleId");
                if (id) defined.add(id);
            }
        }
        for (const required of REQUIRED_DEFAULT_STYLES) {
            if (!defined.has(required.styleId)) {
                issues.push({
                    severity,
                    message:
                        `${this.relPath(stylesPath)} is missing the implied-default style '${required.styleId}' ` +
                        `(type='${required.type}'). ECMA-376 §17.7.4.4 requires a default for each of the ` +
                        `four style types; missing 'TableNormal' in particular triggers Word's ` +
                        `"Document Recovery — Table Properties" dialog on open.`,
                    path: this.relPath(stylesPath),
                    code: "style-default-missing",
                });
            }
        }
        return finalize(issues);
    }

    /**
     * Detect unresolved OpenXmlPowerTools DocumentBuilder `<Insert>` directives
     * (namespace `http://powertools.codeplex.com/documentbuilder/2011/insert`,
     * the only URI `PtOpenXml.ptOpenXml` publishes) left in a
     * WordprocessingML content part. Word cannot open a document whose
     * header/footer/body content model contains this foreign element — it reports
     * "Word experienced an error trying to open the file" (a hard OPEN_ERROR).
     *
     * Root-caused against REAL Word over the word-invalid-fixtures corpus
     * (Docxodus DB/HeadersFooters/Dest/{Fax,Letter,…}): removing only the
     * `<Insert>` element flips every affected file to clean, and 0/350
     * Word-clean files carry the namespace. Word-blocking in every profile.
     */
    async validateDocumentBuilderInserts(): Promise<ValidationResult> {
        const issues: ValidationIssue[] = [];
        for (const xmlFile of this.xmlFiles) {
            const rel = this.relPath(xmlFile);
            // WML content parts only; customXml/ legitimately carries foreign content.
            if (!/^word\/(?:document|header\d*|footer\d*|footnotes|endnotes|glossary\/document)\.xml$/i.test(rel)) {
                continue;
            }
            let content: string;
            try {
                content = decodeXmlBytes(await fs.readFile(xmlFile));
            } catch {
                continue;
            }
            if (!declaresDocumentBuilderInsertNs(content)) continue;
            let dom: Document;
            try {
                dom = parseXml(content);
            } catch {
                continue;
            }
            const all = getElementsByTagNameAll(dom);
            for (let i = 0; i < all.length; i += 1) {
                const elem = all[i];
                const ns = elem?.namespaceURI ?? "";
                if (DOCUMENT_BUILDER_INSERT_NAMESPACES.has(ns)) {
                    issues.push({
                        severity: "error",
                        message:
                            `<${elem?.nodeName}> is an unresolved OpenXmlPowerTools DocumentBuilder directive ` +
                            `(namespace '${ns}') in ${rel}. Word cannot open a document with this foreign element ` +
                            `in its content model; the directive must be resolved/removed before the document is finalized.`,
                        path: rel,
                        code: "documentbuilder-insert-unresolved",
                    });
                    break; // one finding per part is enough
                }
            }
        }
        return finalize(issues);
    }

    /**
     * Detect docProps/app.xml boolean extended-properties (`ScaleCrop`,
     * `LinksUpToDate`, `SharedDoc`, `HyperlinksChanged`) whose value is not a
     * canonical `xsd:boolean` (`true`/`false`/`0`/`1`) — in practice, whitespace
     * around the value, e.g. `<ScaleCrop>false\n  </ScaleCrop>`. Word reports
     * "unreadable content" and offers recovery.
     *
     * Root-caused against REAL Word: stripping ONLY the boolean whitespace flips
     * every affected file to clean, and 0/350 Word-clean files carry it. Integer
     * properties with whitespace (e.g. `<TotalTime>0\n  </TotalTime>`) are
     * TOLERATED by Word and DO occur in Word-clean files, so they are
     * deliberately NOT flagged here (that would re-introduce false positives).
     */
    async validateDocPropsBooleans(): Promise<ValidationResult> {
        const issues: ValidationIssue[] = [];
        const appXml = path.join(this.unpackedDir, "docProps", "app.xml");
        let dom: Document;
        try {
            dom = await this.readPartDom(appXml);
        } catch {
            return { valid: true, issues: [] };
        }
        const BOOLEAN_PROPS = new Set(["ScaleCrop", "LinksUpToDate", "SharedDoc", "HyperlinksChanged"]);
        const CANONICAL = new Set(["true", "false", "0", "1"]);
        const all = getElementsByTagNameAll(dom);
        for (let i = 0; i < all.length; i += 1) {
            const elem = all[i];
            if (!elem || !BOOLEAN_PROPS.has(elem.localName ?? "")) continue;
            const raw = elem.textContent ?? "";
            if (raw === raw.trim() && CANONICAL.has(raw)) continue;
            issues.push({
                severity: "error",
                message:
                    `docProps/app.xml boolean property '${elem.localName}' has the non-canonical value ` +
                    `${JSON.stringify(raw)} (expected exactly true/false/0/1 with no surrounding whitespace). ` +
                    `Word rejects this as unreadable content and offers to recover the file.`,
                path: "docProps/app.xml",
                code: "docprops-boolean-invalid",
            });
        }
        return finalize(issues);
    }

    async validateStyleReferences(): Promise<ValidationResult> {
        const issues: ValidationIssue[] = [];
        const severity: "error" | "warning" = this.profile === "strict" ? "error" : "warning";

        // If the main document relates no styles part, every main-story
        // reference is dangling — but most documents lacking one are
        // abnormal in other ways already detected. A related styles part
        // that is not XML defines what this check cannot read, so the main
        // story is not judged (#693). A reached glossary's story is judged
        // against the glossary's own styles part, or not judged when it
        // relates none (#698).
        const stylesPath = await this.mainStylesPath();
        const glossaryStyles = await this.glossaryStoryStylesFiles();
        const definedByStyles = new Map<string | null, ReadonlySet<string> | undefined>();
        const definedFor = async (styles: string | null): Promise<ReadonlySet<string> | undefined> => {
            if (!definedByStyles.has(styles)) {
                definedByStyles.set(styles, await this.definedStyleIds(styles));
            }
            return definedByStyles.get(styles);
        };

        const refTags = ["pStyle", "rStyle", "tblStyle", "numStyle", "linkedStyle"] as const;
        // Track each (file, styleId) pair so we don't spam multiple
        // issues for the same dangling reference repeated 100 times.
        const reported = new Set<string>();
        for (const xmlFile of this.userTextXmlFiles()) {
            let styles: string | null;
            let stylesLabel: string;
            if (glossaryStyles.has(xmlFile)) {
                const glossaryPath = glossaryStyles.get(xmlFile);
                if (glossaryPath === undefined) continue;
                styles = glossaryPath;
                stylesLabel = this.relPath(glossaryPath);
            } else {
                if (stylesPath === undefined) continue;
                styles = stylesPath;
                // `null` means no styles part was found, so every
                // reference dangles against none (#695).
                stylesLabel = stylesPath === null ? "any styles part" : this.relPath(stylesPath);
            }
            const defined = await definedFor(styles);
            if (defined === undefined) continue;
            let dom: Document;
            try {
                dom = await this.readPartDom(xmlFile);
            } catch {
                continue;
            }
            for (const tag of refTags) {
                for (const ns of WORD_PARAGRAPH_NAMESPACES) {
                    const list = getElementsByTagNameNSAll(dom, ns, tag);
                    for (let i = 0; i < list.length; i += 1) {
                        const elem = list[i];
                        if (!elem) continue;
                        const val = elem.getAttributeNS(ns, "val");
                        if (!val) continue;
                        if (defined.has(val)) continue;
                        const key = `${xmlFile}|${tag}|${val}`;
                        if (reported.has(key)) continue;
                        reported.add(key);
                        issues.push({
                            severity,
                            message:
                                `<w:${tag} w:val='${val}'/> references a style not defined in ${stylesLabel}. ` +
                                `Word substitutes the default style and triggers a repair dialog.`,
                            path: this.relPath(xmlFile),
                            code: "style-reference-undefined",
                        });
                    }
                }
            }
        }
        return finalize(issues);
    }

    // ----- paraId completeness (Word's tracking-anchor convention) ------------

    /**
     * Profile-aware: every `<w:p>` AND every `<w:tr>` should carry a
     * `w14:paraId` attribute, AND any element that has paraId must
     * also have a `w14:textId`. The OOXML schema makes both optional,
     * but Word's tracked-changes machinery treats them as a pair —
     * empirically, across 226 real-world SuperDoc fixtures, ZERO
     * `<w:tr>` elements were observed with paraId present and textId
     * absent. Elements with NEITHER attribute are common and Word
     * handles them silently. The asymmetric "paraId without textId"
     * shape is what trips Word's "Document Recovery — Table
     * Properties" dialog on open, because the row's paraId cannot be
     * linked to any text-content fingerprint for change tracking.
     *
     * Severity:
     *   - `strict`  → `error` (spec-purist + Word-parity).
     *   - `lenient` → `warning` (ours-vs-Word stylistic difference; the
     *     schema allows it, but downstream collaboration tooling may
     *     need it).
     *
     * Surfaced by the comparison against Word's `reallyrepaired.docx`
     * for the sample-document fixture: jubarte's writer leaves table
     * rows without paraId/textId; Word stamps both on all 13 of them.
     */
    async validateAllParagraphsHaveParaId(): Promise<ValidationResult> {
        const issues: ValidationIssue[] = [];
        const severity: "error" | "warning" = this.profile === "strict" ? "error" : "warning";
        for (const xmlFile of this.userTextXmlFiles()) {
            let dom: Document;
            try {
                dom = await this.readPartDom(xmlFile);
            } catch {
                continue;
            }
            for (const local of ["p", "tr"] as const) {
                let missingParaId = 0;
                let asymmetric = 0;
                for (const ns of WORD_PARAGRAPH_NAMESPACES) {
                    const list = getElementsByTagNameNSAll(dom, ns, local);
                    for (let i = 0; i < list.length; i += 1) {
                        const elem = list[i];
                        if (!elem) continue;
                        const paraId = elem.getAttributeNS(W14_NAMESPACE, "paraId");
                        const textId = elem.getAttributeNS(W14_NAMESPACE, "textId");
                        if (!paraId) missingParaId += 1;
                        else if (!textId) asymmetric += 1;
                    }
                }
                if (missingParaId > 0) {
                    issues.push({
                        severity,
                        message:
                            `${missingParaId} <w:${local}> element(s) lack a w14:paraId. ` +
                            `Word stamps one on every paragraph and table row as the anchor for ` +
                            `tracked-changes / comment-range infrastructure.`,
                        path: this.relPath(xmlFile),
                        code: "paraid-missing-element",
                    });
                }
                if (asymmetric > 0) {
                    issues.push({
                        severity,
                        message:
                            `${asymmetric} <w:${local}> element(s) carry a w14:paraId but no ` +
                            `w14:textId. The pair is what Word's tracked-changes machinery uses ` +
                            `to anchor revisions; an asymmetric paraId-without-textId row is ` +
                            `what triggers Word's "Document Recovery — Table Properties" dialog.`,
                        path: this.relPath(xmlFile),
                        code: "textid-missing-element",
                    });
                }
            }
        }
        return finalize(issues);
    }

    // ----- run-inner content placement ----------------------------------------

    /**
     * EG_RunInnerContent elements (`<w:t>`, `<w:tab>`, `<w:br>`, `<w:cr>`,
     * `<w:drawing>`, `<w:pict>`, `<w:object>`, `<w:sym>`, `<w:softHyphen>`,
     * `<w:noBreakHyphen>`, `<w:ptab>`, `<w:fldChar>`, `<w:instrText>`,
     * `<w:delText>`, `<w:lastRenderedPageBreak>`, the note/comment references,
     * …) are valid ONLY inside a `<w:r>`. When emitted as a DIRECT child of
     * `<w:p>` — a sibling of runs, e.g. `<w:r>…</w:r><w:tab/><w:r>…` produced by
     * a foreign pipeline that forgot to wrap them — they are not in EG_PContent
     * (invalid OOXML). A real Word probe shows Word OPENS such files cleanly and
     * silently relocates/drops the misplaced content — so this is structural
     * invalidity + silent data loss, NOT a Word-openability failure: an error in
     * lenient/strict, tolerated by the word-valid profile. The bare XSD
     * validator surfaces it only generically ("this element is not expected");
     * this dedicated rule names the misplacement.
     *
     * Reports `run-content-misplaced` at "error" severity.
     */
    async validateRunContentPlacement(): Promise<ValidationResult> {
        const issues: ValidationIssue[] = [];
        for (const xmlFile of this.userTextXmlFiles()) {
            let dom: Document;
            try {
                dom = await this.readPartDom(xmlFile);
            } catch {
                continue;
            }
            for (const ns of WORD_PARAGRAPH_NAMESPACES) {
                const paragraphs = getElementsByTagNameNSAll(dom, ns, "p");
                for (let i = 0; i < paragraphs.length; i += 1) {
                    const para = paragraphs[i];
                    if (!para) continue;
                    for (let child = para.firstChild; child; child = child.nextSibling) {
                        if (child.nodeType !== 1) continue;
                        const elem = child as Element;
                        if (elem.namespaceURI !== ns) continue;
                        const local = elem.localName || elem.tagName.split(":").pop() || "";
                        if (!RUN_INNER_CONTENT_LOCALS.has(local)) continue;
                        issues.push({
                            severity: "error",
                            message:
                                `<w:${local}> is run-inner content (EG_RunInnerContent) placed ` +
                                `directly under <w:p>; it must sit inside a <w:r>. Word drops ` +
                                `misplaced run content on open.`,
                            path: this.relPath(xmlFile),
                            code: "run-content-misplaced",
                        });
                    }
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
     *   6. Every `w16cid:commentId/@paraId` in commentsIds.xml must resolve
     *      to a paragraph in comments.xml.
     *   7. Every `w16cex:commentExtensible/@durableId` in
     *      commentsExtensible.xml must resolve to a durableId in
     *      commentsIds.xml.
     *
     * `validateCommentMarkers` already covers orphan range start/end and
     * missing-comment references — this validator is strictly about the
     * threaded-comments extension surface.
     */
    async validateCommentThreading(): Promise<ValidationResult> {
        const run = await runChecks(this.context(), createDefaultCheckRegistry(), {
            format: "docx",
            select: ["comment-threading"],
            profile: this.profile,
            deadline: this.deadline,
        });
        this.recordExecutedChecks(run.executed);
        return finalize(run.issues.map((issue) => promoteFullCheckIssue(issue, this.profile, "docx")));
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
                raw = decodeXmlBytes(await fs.readFile(xmlFile));
            } catch {
                continue;
            }
            // Cheap pre-filter — if no `[[DOCX_` substring, skip the regex.
            if (!TRACKING_TOKEN_PREFIXES.some((p) => raw.includes(p))) continue;

            TRACKING_TOKEN_REGEX.lastIndex = 0;
            let match: RegExpExecArray | null;
            const seen = new Set<string>();
            while ((match = TRACKING_TOKEN_REGEX.exec(raw)) !== null) {
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
                dom = await this.readPartDom(xmlFile);
            } catch {
                // Mirrors Python's bare except — silently skip.
                continue;
            }
            const all = getElementsByTagNameAll(dom);
            const base = baseName(xmlFile);
            for (let i = 0; i < all.length; i += 1) {
                const elem = all[i];
                if (!elem) continue;

                const paraId = elem.getAttributeNS(W14_NAMESPACE, "paraId");
                if (paraId) {
                    const v = parseIdValue(paraId, 16);
                    if (v >= MAX_PARA_ID) {
                        issues.push({
                            severity: "error",
                            message: `${base}: paraId=${paraId} >= 0x80000000`,
                            path: this.relPath(xmlFile),
                            code: "id-paraid-overflow",
                        });
                    } else if (v === 0) {
                        // [MS-OI29500] 2.6.2.3: paraId MUST be greater than 0.
                        issues.push({
                            severity: "error",
                            message: `${base}: paraId=${paraId} must be > 0`,
                            path: this.relPath(xmlFile),
                            code: "id-paraid-zero",
                        });
                    }
                }

                // textId shares the ST_LongHexNumber type with paraId and is
                // bound by the same < 0x80000000 cap. Word's parser silently
                // recovers any over-cap textId on a <w:p>/<w:tr>; the recovery
                // cascades through the row's table-properties machinery and
                // surfaces as a "Document Recovery — Table Properties" dialog
                // on open. Empirically: across 226 SuperDoc real-world fixtures
                // the cap is never exceeded; a Plate-pipeline export had 51
                // textIds over the cap (every paragraph in the body).
                const textId = elem.getAttributeNS(W14_NAMESPACE, "textId");
                if (textId) {
                    const v = parseIdValue(textId, 16);
                    if (v >= MAX_PARA_ID) {
                        issues.push({
                            severity: "error",
                            message: `${base}: textId=${textId} >= 0x80000000`,
                            path: this.relPath(xmlFile),
                            code: "id-textid-overflow",
                        });
                    } else if (v === 0) {
                        // textId shares the ST_LongHexNumber type with paraId; same > 0 floor.
                        issues.push({
                            severity: "error",
                            message: `${base}: textId=${textId} must be > 0`,
                            path: this.relPath(xmlFile),
                            code: "id-textid-zero",
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
                        } else if (v === 0) {
                            // durableId is a positive id; the repair path seeds it from 1.
                            issues.push({
                                severity: "error",
                                message: `${base}: durableId=${durableId} must be > 0`,
                                path: this.relPath(xmlFile),
                                code: "id-durable-zero",
                            });
                        }
                    } else {
                        const v = parseIdValue(durableId, 16);
                        if (v >= MAX_DURABLE_ID) {
                            issues.push({
                                severity: "error",
                                message: `${base}: durableId=${durableId} >= 0x7FFFFFFF`,
                                path: this.relPath(xmlFile),
                                code: "id-durable-overflow",
                            });
                        } else if (v === 0) {
                            // durableId is a positive id; the repair path seeds it from 1.
                            issues.push({
                                severity: "error",
                                message: `${base}: durableId=${durableId} must be > 0`,
                                path: this.relPath(xmlFile),
                                code: "id-durable-zero",
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
                const dom = parseXmlBytes(readFileSync(xmlFile));
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

    /**
     * Parse `word/document.xml` from the `--original` zip, memoized in the
     * current baseline scope: one `validate()` run shares its parse, direct
     * calls share the outside-run slot {@link BaseSchemaValidator.forgetOriginal}
     * clears, and an overlapping run keeps its own (#481, #503, #520).
     */
    private async loadOriginalDocumentRoot(): Promise<Element | null> {
        const state = this.baselineState();
        if (state.documentRoot) return state.documentRoot;
        if (state.documentRootFailed) return null;
        if (!this.originalFile) {
            state.documentRootFailed = true;
            return null;
        }
        try {
            // Shares the archive `_getOriginalFileErrors` memoizes: one read
            // and load per baseline scope instead of one per caller.
            const zip = await this._loadOriginalArchive();
            const entry = zip?.file("word/document.xml");
            if (!entry) {
                state.documentRootFailed = true;
                return null;
            }
            const text = decodeXmlBytes(await entry.async("uint8array"));
            const dom = parseXml(text);
            state.documentRoot = dom.documentElement;
            return state.documentRoot;
        } catch {
            state.documentRootFailed = true;
            return null;
        }
    }

    private async originalUsesStrictNamespace(): Promise<boolean> {
        const root = await this.loadOriginalDocumentRoot();
        if (!root) return false;
        return root.namespaceURI === WORD_STRICT_NAMESPACE;
    }

    // ----- Plan 01: Whole-file preservation -----------------------------------

    /**
     * Check every `.rels` file and verify that each internal `<Relationship>`
     * target resolves to an existing file in the unpacked directory.
     * External (http/https) targets are skipped.
     *
     * Reports `rels-target-missing` (error) when a target file is absent.
     */
    async validateOrphanedRelationships(): Promise<ValidationResult> {
        const issues: ValidationIssue[] = [];
        const entries = packageEntries(this.unpackedDir);
        const names = [...entries.keys()];
        for (const [name, xmlFile] of entries) {
            if (!isRelationshipsPartName(name)) continue;
            let dom: Document;
            try {
                dom = await this.readPartDom(xmlFile);
            } catch {
                continue;
            }
            const list = this.relationshipElements(dom, name);
            for (let i = 0; i < list.length; i += 1) {
                const elem = list[i];
                if (!elem) continue;
                const target = elem.getAttribute("Target");
                if (!target) continue;
                if (isExternalRelationship(elem, target)) continue;
                const resolved = resolveRelationshipTargetPath(this.unpackedDir, path.join(this.unpackedDir, name), target);
                if (!resolved) {
                    issues.push({
                        severity: "error",
                        message: `Relationship target '${target}' escapes the unpacked directory or is invalid`,
                        path: this.relPath(xmlFile),
                        code: "rels-target-missing",
                    });
                    continue;
                }
                // The target names a part by OPC equivalence (ASCII case,
                // escapes of unreserved characters), not by a file-system
                // lookup of its spelling (#346).
                const part = resolvePartName(names, this.entryName(resolved));
                const file = part === undefined ? undefined : entries.get(part);
                if (file === undefined || containRelationshipTarget(this.unpackedDir, file) === null) {
                    issues.push({
                        severity: "error",
                        message: `Relationship target '${target}' resolves to '${this.relPath(resolved)}' which does not exist`,
                        path: this.relPath(xmlFile),
                        code: "rels-target-missing",
                    });
                }
            }
        }
        return finalize(issues);
    }

    // ----- Plan 02: Font table retention --------------------------------------

    /**
     * Check the font-table part the main document relates
     * (by name in its folder when none does, #690, #699) for empty
     * `<w:fonts>` (no `<w:font>` children).
     * An empty font table is XSD-valid but almost always indicates a repairer bug
     * that discards all font metadata.
     *
     * Reports `font-table-empty` at `"warning"` severity.
     */
    async validateFontTable(): Promise<ValidationResult> {
        const issues: ValidationIssue[] = [];
        const fontTablePath = await this.mainRelatedPath("fontTable", "fontTable.xml");
        if (!fontTablePath) return { valid: true, issues: [] };

        let dom: Document;
        try {
            dom = await this.readPartDom(fontTablePath);
        } catch {
            return { valid: true, issues: [] };
        }
        const fontList: Element[] = [];
        for (const ns of WORD_PARAGRAPH_NAMESPACES) {
            const list = getElementsByTagNameNSAll(dom, ns, "font");
            for (let i = 0; i < list.length; i += 1) {
                const elem = list[i];
                if (elem) fontList.push(elem);
            }
        }
        if (fontList.length === 0) {
            issues.push({
                severity: "warning",
                message:
                    `${this.relPath(fontTablePath)} contains zero <w:font> entries. ` +
                    "This causes Word to substitute all fonts with its default, " +
                    "silently changing document appearance.",
                path: this.relPath(fontTablePath),
                code: "font-table-empty",
            });
        }
        return finalize(issues);
    }

    // ----- Plan 03: Style passthrough -----------------------------------------

    /**
     * Check the styles part the main document relates (by name in its
     * folder when none does, #688, #699) for the presence of `<w:latentStyles>`.
     * Missing latent styles is XSD-valid but unusual for Word-generated files.
     *
     * Reports `latent-styles-missing` at `"info"` severity.
     */
    async validateLatentStyles(): Promise<ValidationResult> {
        const issues: ValidationIssue[] = [];
        const stylesPath = await this.mainStylesPath();
        if (!stylesPath) return { valid: true, issues: [] };

        let dom: Document;
        try {
            dom = await this.readPartDom(stylesPath);
        } catch {
            return { valid: true, issues: [] };
        }
        let hasLatentStyles = false;
        for (const ns of WORD_PARAGRAPH_NAMESPACES) {
            if (getElementsByTagNameNSAll(dom, ns, "latentStyles").length > 0) {
                hasLatentStyles = true;
                break;
            }
        }
        if (!hasLatentStyles) {
            issues.push({
                severity: "info",
                message:
                    `${this.relPath(stylesPath)} has no <w:latentStyles> block. ` +
                    "This is unusual for Word-generated files and indicates " +
                    "the repairer may have regenerated styles from scratch.",
                path: this.relPath(stylesPath),
                code: "latent-styles-missing",
            });
        }
        return finalize(issues);
    }

    // ----- Plan 05: tblLook preservation --------------------------------------

    /**
     * Check every `<w:tbl>` in `word/document.xml` that references a
     * `<w:tblStyle>` for the presence of `<w:tblLook>`. Missing `<w:tblLook>`
     * prevents conditional formatting bands from applying.
     *
     * Reports `tbl-look-missing` at `"info"` severity.
     */
    async validateTableLook(): Promise<ValidationResult> {
        const issues: ValidationIssue[] = [];
        for (const xmlFile of this.documentXmlFiles()) {
            let dom: Document;
            try {
                dom = await this.readPartDom(xmlFile);
            } catch {
                continue;
            }
            for (const ns of WORD_PARAGRAPH_NAMESPACES) {
                const tbls = getElementsByTagNameNSAll(dom, ns, "tbl");
                for (let i = 0; i < tbls.length; i += 1) {
                    const tbl = tbls[i];
                    if (!tbl) continue;
                    const tblPr = directChild(tbl, ns, "tblPr");
                    if (!tblPr) continue;
                    const hasTblStyle = directChild(tblPr, ns, "tblStyle") !== null;
                    if (!hasTblStyle) continue;
                    const hasTblLook = directChild(tblPr, ns, "tblLook") !== null;
                    if (!hasTblLook) {
                        issues.push({
                            severity: "info",
                            message:
                                "<w:tbl> has a <w:tblStyle> reference but no <w:tblLook> in <w:tblPr>. " +
                                "Conditional formatting bands from the table style will not apply.",
                            path: this.relPath(xmlFile),
                            code: "tbl-look-missing",
                        });
                    }
                }
            }
        }
        return finalize(issues);
    }

    // ----- Plan 06: Tracked-change ID stability --------------------------------

    /**
     * Check tracked-change elements (`<w:ins>`, `<w:del>`) for sequential `w:id`
     * values starting from 1. A sequential-from-1 pattern is a strong heuristic
     * signal that the repairer regenerated IDs rather than preserving originals.
     *
     * Reports `tracked-change-ids-regenerated` at `"info"` severity.
     */
    async validateTrackedChangeIds(): Promise<ValidationResult> {
        const issues: ValidationIssue[] = [];
        for (const xmlFile of this.documentXmlFiles()) {
            let dom: Document;
            try {
                dom = await this.readPartDom(xmlFile);
            } catch {
                continue;
            }
            const ids: number[] = [];
            let hasUnparseableId = false;
            for (const ns of WORD_PARAGRAPH_NAMESPACES) {
                for (const tag of TRACKED_CHANGE_ID_TAGS) {
                    const list = getElementsByTagNameNSAll(dom, ns, tag);
                    for (let j = 0; j < list.length; j += 1) {
                        const elem = list[j];
                        if (!elem) continue;
                        const id = elem.getAttributeNS(ns, "id");
                        if (!id) continue;
                        if (!/^\d+$/.test(id)) {
                            hasUnparseableId = true;
                        } else {
                            ids.push(Number.parseInt(id, 10));
                        }
                    }
                }
            }
            if (ids.length < 2 || hasUnparseableId) continue;
            ids.sort((a, b) => a - b);
            // Heuristic: if IDs are sequential from 1 with no gaps, likely regenerated.
            const isSequentialFrom1 = ids.every((id, idx) => id === idx + 1);
            if (isSequentialFrom1) {
                issues.push({
                    severity: "info",
                    message:
                        `Tracked-change w:id values (${ids.join(", ")}) appear to be regenerated ` +
                        "(sequential from 1). This breaks external references to tracked changes.",
                    path: this.relPath(xmlFile),
                    code: "tracked-change-ids-regenerated",
                });
            }
        }
        return finalize(issues);
    }

    // ----- Plan 07: Cell border normalization ---------------------------------

    /**
     * Check for cell borders (`<w:tcBorders>`) that duplicate the enclosing
     * table's `<w:tblBorders>` defaults. Redundant borders override style-level
     * conditional formatting (e.g. `<w:tblLook>` banded-row suppression).
     *
     * Reports `cell-borders-redundant` at `"info"` severity.
     */
    async validateRedundantCellBorders(): Promise<ValidationResult> {
        const issues: ValidationIssue[] = [];
        // Each story's table styles come from the styles part its root
        // relates (#698): a reached glossary's `document.xml` against the
        // glossary's own — or not judged when it relates none — and every
        // other `document.xml` against the main document's.
        const stylesPath = await this.mainStylesPath();
        const glossaryStyles = await this.glossaryStoryStylesFiles();
        const bordersByStyles = new Map<string | null, Map<string, BorderSignature>>();
        const bordersFor = async (styles: string | null): Promise<Map<string, BorderSignature>> => {
            if (!bordersByStyles.has(styles)) {
                bordersByStyles.set(styles, await this.tableStyleBordersByStyleId(styles));
            }
            return bordersByStyles.get(styles)!;
        };
        for (const xmlFile of this.documentXmlFiles()) {
            const inGlossary = glossaryStyles.has(xmlFile);
            const styles = inGlossary ? glossaryStyles.get(xmlFile) : stylesPath;
            if (inGlossary && styles === undefined) continue;
            const styleBordersById = await bordersFor(styles ?? null);
            let dom: Document;
            try {
                dom = await this.readPartDom(xmlFile);
            } catch {
                continue;
            }
            for (const ns of WORD_PARAGRAPH_NAMESPACES) {
                const tbls = getElementsByTagNameNSAll(dom, ns, "tbl");
                for (let i = 0; i < tbls.length; i += 1) {
                    const tbl = tbls[i];
                    if (!tbl) continue;
                    const tblPr = directChild(tbl, ns, "tblPr");
                    if (!tblPr) continue;
                    const tblStyleId = directChild(tblPr, ns, "tblStyle")?.getAttributeNS(ns, "val") ?? null;
                    const tableDefaults =
                        borderSignature(directChild(tblPr, ns, "tblBorders"), ns) ??
                        (tblStyleId ? styleBordersById.get(tblStyleId) : undefined);
                    if (!tableDefaults) continue;

                    for (const row of directChildren(tbl, ns, "tr")) {
                        for (const cell of directChildren(row, ns, "tc")) {
                            const tcPr = directChild(cell, ns, "tcPr");
                            if (!tcPr) continue;
                            const cellBorders = borderSignature(directChild(tcPr, ns, "tcBorders"), ns);
                            if (cellBorders && borderSignaturesEqual(cellBorders, tableDefaults)) {
                                issues.push({
                                    severity: "info",
                                    message:
                                        "<w:tcBorders> in table cell duplicates the table-level <w:tblBorders> " +
                                        "defaults. This overrides style-level conditional formatting.",
                                    path: this.relPath(xmlFile),
                                    code: "cell-borders-redundant",
                                });
                            }
                        }
                    }
                }
            }
        }
        return finalize(issues);
    }

    // ----- Plan 08: Relationship ID scheme ------------------------------------

    /**
     * Check relationship `.rels` files for the `rId1`-`rIdN` sequential pattern.
     * Sequential IDs are valid but indicate the repairer may have regenerated
     * relationship IDs rather than preserving the originals.
     *
     * Reports `rel-ids-sequential` at `"info"` severity.
     */
    async validateRelationshipIdStability(): Promise<ValidationResult> {
        const issues: ValidationIssue[] = [];
        for (const xmlFile of this.xmlFiles) {
            const part = this.entryName(xmlFile);
            if (!isRelationshipsPartName(part)) continue;
            let dom: Document;
            try {
                dom = await this.readPartDom(xmlFile);
            } catch {
                continue;
            }
            const relsList = this.relationshipElements(dom, part);
            const ids: number[] = [];
            let allIdsUseRidPattern = true;
            for (let i = 0; i < relsList.length; i += 1) {
                const elem = relsList[i];
                if (!elem) continue;
                const id = elem.getAttribute("Id");
                if (!id) continue;
                const match = /^rId(\d+)$/.exec(id);
                if (match) {
                    ids.push(parseInt(match[1], 10));
                } else {
                    allIdsUseRidPattern = false;
                }
            }
            if (!allIdsUseRidPattern || ids.length < 2) continue;
            ids.sort((a, b) => a - b);
            const isSequential = ids.every((id, idx) => id === idx + 1);
            if (isSequential) {
                issues.push({
                    severity: "info",
                    message:
                        `Relationship Id values in ${this.relPath(xmlFile)} are sequential from rId1 ` +
                        "(likely regenerated by the repairer).",
                    path: this.relPath(xmlFile),
                    code: "rel-ids-sequential",
                });
            }
        }
        return finalize(issues);
    }

    // ----- Plan 09: Redundant explicit properties -----------------------------

    /**
     * Check for explicit `<w:rFonts>` on runs that duplicate the document
     * defaults (`<w:docDefaults>/<w:rPrDefault>/<w:rPr>/<w:rFonts>`).
     * Redundant explicit properties make the document harder to re-style.
     *
     * Reports `run-props-redundant` at `"info"` severity.
     */
    async validateRedundantRunProperties(): Promise<ValidationResult> {
        const issues: ValidationIssue[] = [];

        // Each story's docDefaults come from the styles part its root
        // relates (#698): a reached glossary's runs against the glossary's
        // own styles part — or not judged when it relates none — and every
        // other part against the main document's.
        const stylesPath = await this.mainStylesPath();
        const glossaryStyles = await this.glossaryStoryStylesFiles();
        const defaultsByStyles = new Map<string, { ascii: string; hAnsi: string; cs: string } | undefined>();
        const defaultsFor = async (styles: string): Promise<{ ascii: string; hAnsi: string; cs: string } | undefined> => {
            if (!defaultsByStyles.has(styles)) {
                defaultsByStyles.set(styles, await this.docDefaultRunFonts(styles));
            }
            return defaultsByStyles.get(styles);
        };

        // Check every user-text part; redundant run properties commonly show
        // up in headers and footers as well as the main document body.
        for (const xmlFile of this.userTextXmlFiles()) {
            const styles = glossaryStyles.has(xmlFile) ? glossaryStyles.get(xmlFile) : stylesPath;
            if (!styles) continue;
            const defaults = await defaultsFor(styles);
            if (defaults === undefined || (!defaults.ascii && !defaults.hAnsi && !defaults.cs)) continue;
            let dom: Document;
            try {
                dom = await this.readPartDom(xmlFile);
            } catch {
                continue;
            }
            for (const ns of WORD_PARAGRAPH_NAMESPACES) {
                const runs = getElementsByTagNameNSAll(dom, ns, "r");
                for (let i = 0; i < runs.length; i += 1) {
                    const run = runs[i];
                    if (!run) continue;
                    const rPr = directChild(run, ns, "rPr");
                    if (!rPr) continue;
                    const rFonts = directChild(rPr, ns, "rFonts");
                    if (!rFonts) continue;
                    const ascii = rFonts.getAttributeNS(ns, "ascii") ?? "";
                    const hAnsi = rFonts.getAttributeNS(ns, "hAnsi") ?? "";
                    const cs = rFonts.getAttributeNS(ns, "cs") ?? "";
                    if (ascii && ascii === defaults.ascii && hAnsi === defaults.hAnsi && cs === defaults.cs) {
                        issues.push({
                            severity: "info",
                            message:
                                `<w:rFonts w:ascii="${ascii}" w:hAnsi="${hAnsi}" w:cs="${cs}" /> ` +
                                "on run duplicates <w:docDefaults>; redundant explicit property.",
                            path: this.relPath(xmlFile),
                            code: "run-props-redundant",
                        });
                    }
                }
            }
        }
        return finalize(issues);
    }

    /**
     * The document-default run fonts of a styles part —
     * `<w:docDefaults>/<w:rPrDefault>/<w:rPr>/<w:rFonts>` — or `undefined`
     * when the part cannot be parsed or declares none.
     */
    private async docDefaultRunFonts(stylesPath: string): Promise<{ ascii: string; hAnsi: string; cs: string } | undefined> {
        let stylesDom: Document;
        try {
            stylesDom = await this.readPartDom(stylesPath);
        } catch {
            return undefined;
        }
        for (const ns of WORD_PARAGRAPH_NAMESPACES) {
            const ddList = getElementsByTagNameNSAll(stylesDom, ns, "docDefaults");
            if (ddList.length === 0) continue;
            const dd = ddList[0];
            if (!dd) continue;
            const rPrDefaultList = getElementsByTagNameNSAll(dd, ns, "rPrDefault");
            if (rPrDefaultList.length === 0) continue;
            const rPrDefault = rPrDefaultList[0];
            if (!rPrDefault) continue;
            const rPrList = getElementsByTagNameNSAll(rPrDefault, ns, "rPr");
            if (rPrList.length === 0) continue;
            const rPr = rPrList[0];
            if (!rPr) continue;
            const rFontsList = getElementsByTagNameNSAll(rPr, ns, "rFonts");
            if (rFontsList.length === 0) continue;
            const rFonts = rFontsList[0];
            if (!rFonts) continue;
            return {
                ascii: rFonts.getAttributeNS(ns, "ascii") ?? "",
                hAnsi: rFonts.getAttributeNS(ns, "hAnsi") ?? "",
                cs: rFonts.getAttributeNS(ns, "cs") ?? "",
            };
        }
        return undefined;
    }

    // ----- repair -------------------------------------------------------------

    async repair(): Promise<number> {
        const baseRepairs = await super.repair();
        const durableRepairs = await this.repairDurableId();
        const paraIdRepairs = await this.repairParaId();
        const ignorableRepairs = await this.repairIgnorable();
        const stampRepairs = await this.repairMissingParaIds();
        const styleRepairs = await this.repairMissingStyleDefinitions();
        const commentThreadingRepairs = await this.repairCommentThreading();
        const drawingScalarRepairs = await this.repairDrawingScalarTextWhitespace();
        const inlinePictureRepairs = await this.repairInlinePictureScaffolds();
        const extendedPropertyRepairs = await this.repairExtendedPropertiesWhitespace();
        const corePropertyRepairs = await this.repairCorePropertiesWhitespace();
        return (
            baseRepairs +
            durableRepairs +
            paraIdRepairs +
            ignorableRepairs +
            stampRepairs +
            styleRepairs +
            commentThreadingRepairs +
            drawingScalarRepairs +
            inlinePictureRepairs +
            extendedPropertyRepairs +
            corePropertyRepairs
        );
    }

    async repairDrawingScalarTextWhitespace(): Promise<number> {
        let repairs = 0;
        for (const xmlFile of this.userTextXmlFiles()) {
            let dom: Document;
            try {
                dom = parseXmlBytes(await fs.readFile(xmlFile));
            } catch {
                continue;
            }

            let modified = false;
            // Counted only after the write: a part whose encoding
            // `writeXmlPart` refuses is left untouched (#880).
            let partRepairs = 0;
            for (const { namespace, localName } of WORD_DRAWING_SCALAR_TEXT_ELEMENTS) {
                const elems = getElementsByTagNameNSAll(dom, namespace, localName);
                for (let i = 0; i < elems.length; i += 1) {
                    const elem = elems[i];
                    if (!elem) continue;
                    for (let child = elem.firstChild; child; child = child.nextSibling) {
                        if (child.nodeType !== 3) continue;
                        const text = child as Text;
                        const trimmed = text.data.trim();
                        if (trimmed === text.data) continue;
                        text.data = trimmed;
                        partRepairs += 1;
                        modified = true;
                    }
                }
            }

            if (modified && (await writeXmlPart(xmlFile, dom))) {
                repairs += partRepairs;
            }
        }
        return repairs;
    }

    async repairInlinePictureScaffolds(): Promise<number> {
        let repairs = 0;
        for (const xmlFile of this.userTextXmlFiles()) {
            let dom: Document;
            try {
                dom = parseXmlBytes(await fs.readFile(xmlFile));
            } catch {
                continue;
            }

            let modified = false;
            // Counted only after the write: a part whose encoding
            // `writeXmlPart` refuses is left untouched (#880).
            let partRepairs = 0;
            const usedDocPrIds = collectNumericAttributeValues(dom, WP_NAMESPACE, "docPr", "id");
            const allocateDocPrId = (): string => {
                let candidate = 1;
                while (usedDocPrIds.has(candidate)) candidate += 1;
                usedDocPrIds.add(candidate);
                return String(candidate);
            };

            const inlines = getElementsByTagNameNSAll(dom, WP_NAMESPACE, "inline");
            for (let i = 0; i < inlines.length; i += 1) {
                const inline = inlines[i];
                if (!inline) continue;
                const graphic = directChild(inline, DRAWINGML_NAMESPACE, "graphic");
                if (!graphic) continue;

                if (!directChild(inline, WP_NAMESPACE, "extent")) {
                    const extent = dom.createElementNS(WP_NAMESPACE, "wp:extent");
                    extent.setAttribute("cx", DEFAULT_PICTURE_EXTENT_EMU);
                    extent.setAttribute("cy", DEFAULT_PICTURE_EXTENT_EMU);
                    inline.insertBefore(extent, graphic);
                    partRepairs += 1;
                    modified = true;
                }

                if (!directChild(inline, WP_NAMESPACE, "docPr")) {
                    const docPr = dom.createElementNS(WP_NAMESPACE, "wp:docPr");
                    docPr.setAttribute("id", allocateDocPrId());
                    docPr.setAttribute("name", "Picture 1");
                    inline.insertBefore(docPr, graphic);
                    partRepairs += 1;
                    modified = true;
                }

                const pictures = getElementsByTagNameNSAll(inline, PICTURE_NAMESPACE, "pic");
                for (let j = 0; j < pictures.length; j += 1) {
                    const pic = pictures[j];
                    if (!pic) continue;
                    const blipFill = directChild(pic, PICTURE_NAMESPACE, "blipFill");

                    if (!directChild(pic, PICTURE_NAMESPACE, "nvPicPr")) {
                        const nvPicPr = dom.createElementNS(PICTURE_NAMESPACE, "pic:nvPicPr");
                        const cNvPr = dom.createElementNS(PICTURE_NAMESPACE, "pic:cNvPr");
                        cNvPr.setAttribute("id", "0");
                        cNvPr.setAttribute("name", "image1.png");
                        const cNvPicPr = dom.createElementNS(PICTURE_NAMESPACE, "pic:cNvPicPr");
                        nvPicPr.appendChild(cNvPr);
                        nvPicPr.appendChild(cNvPicPr);
                        pic.insertBefore(nvPicPr, blipFill ?? pic.firstChild);
                        partRepairs += 1;
                        modified = true;
                    }

                    if (
                        blipFill &&
                        !directChild(blipFill, DRAWINGML_NAMESPACE, "stretch") &&
                        !directChild(blipFill, DRAWINGML_NAMESPACE, "tile")
                    ) {
                        const stretch = dom.createElementNS(DRAWINGML_NAMESPACE, "a:stretch");
                        stretch.appendChild(dom.createElementNS(DRAWINGML_NAMESPACE, "a:fillRect"));
                        blipFill.appendChild(stretch);
                        partRepairs += 1;
                        modified = true;
                    }

                    if (!directChild(pic, PICTURE_NAMESPACE, "spPr")) {
                        const spPr = dom.createElementNS(PICTURE_NAMESPACE, "pic:spPr");
                        const xfrm = dom.createElementNS(DRAWINGML_NAMESPACE, "a:xfrm");
                        const off = dom.createElementNS(DRAWINGML_NAMESPACE, "a:off");
                        off.setAttribute("x", "0");
                        off.setAttribute("y", "0");
                        const ext = dom.createElementNS(DRAWINGML_NAMESPACE, "a:ext");
                        ext.setAttribute("cx", DEFAULT_PICTURE_EXTENT_EMU);
                        ext.setAttribute("cy", DEFAULT_PICTURE_EXTENT_EMU);
                        xfrm.appendChild(off);
                        xfrm.appendChild(ext);
                        const prstGeom = dom.createElementNS(DRAWINGML_NAMESPACE, "a:prstGeom");
                        prstGeom.setAttribute("prst", "rect");
                        prstGeom.appendChild(dom.createElementNS(DRAWINGML_NAMESPACE, "a:avLst"));
                        spPr.appendChild(xfrm);
                        spPr.appendChild(prstGeom);
                        pic.appendChild(spPr);
                        partRepairs += 1;
                        modified = true;
                    }
                }
            }

            if (modified && (await writeXmlPart(xmlFile, dom))) {
                repairs += partRepairs;
            }
        }
        return repairs;
    }

    async repairExtendedPropertiesWhitespace(): Promise<number> {
        const appXml = path.join(this.unpackedDir, "docProps", "app.xml");
        let dom: Document;
        try {
            dom = parseXmlBytes(await fs.readFile(appXml));
        } catch {
            return 0;
        }

        let repairs = 0;
        const all = getElementsByTagNameAll(dom);
        for (let i = 0; i < all.length; i += 1) {
            const elem = all[i];
            if (!elem) continue;
            for (let child = elem.firstChild; child; child = child.nextSibling) {
                if (child.nodeType !== 3) continue;
                const text = child as Text;
                const trimmed = text.data.trim();
                if (trimmed === text.data) continue;
                text.data = trimmed;
                repairs += 1;
            }
        }
        if (repairs > 0 && !(await writeXmlPart(appXml, dom))) {
            // The part's encoding cannot be written back: nothing was repaired (#880).
            return 0;
        }
        return repairs;
    }

    async repairCorePropertiesWhitespace(): Promise<number> {
        const coreXml = path.join(this.unpackedDir, "docProps", "core.xml");
        let dom: Document;
        try {
            dom = parseXmlBytes(await fs.readFile(coreXml));
        } catch {
            return 0;
        }

        const trimFields = new Set(["lastModifiedBy", "revision", "created", "modified"]);
        let repairs = 0;
        const all = getElementsByTagNameAll(dom);
        for (let i = 0; i < all.length; i += 1) {
            const elem = all[i];
            if (!elem) continue;
            const localName = elem.localName || elem.tagName.split(":").pop() || elem.tagName;
            if (!trimFields.has(localName)) continue;
            for (let child = elem.firstChild; child; child = child.nextSibling) {
                if (child.nodeType !== 3) continue;
                const text = child as Text;
                const trimmed = text.data.trim();
                if (trimmed === text.data) continue;
                text.data = trimmed;
                repairs += 1;
            }
        }
        if (repairs > 0 && !(await writeXmlPart(coreXml, dom))) {
            // The part's encoding cannot be written back: nothing was repaired (#880).
            return 0;
        }
        return repairs;
    }

    /**
     * Align the threaded-comment parts, preflighting the encoding and
     * writability of every part to be written before changing any of them.
     * A failed or refused write returns only the repairs already written
     * (#909, #920).
     */
    async repairCommentThreading(): Promise<number> {
        const commentsXml = await this.mainRelatedPath("comments", "comments.xml");
        if (!commentsXml) return 0;
        const commentsIdsXml = await this.mainThreadPath(COMMENTS_IDS_RELATIONSHIP_TYPE, commentsXml, "commentsIds.xml");
        const commentsExtensibleXml = await this.mainThreadPath(
            COMMENTS_EXTENSIBLE_RELATIONSHIP_TYPE,
            commentsXml,
            "commentsExtensible.xml",
        );
        if (!commentsIdsXml) return 0;

        let commentsDom: Document;
        let idsDom: Document;
        try {
            commentsDom = parseXmlBytes(await fs.readFile(commentsXml));
            idsDom = parseXmlBytes(await fs.readFile(commentsIdsXml));
        } catch {
            return 0;
        }

        const commentParaIds = new Set<string>();
        for (const wNs of WORD_PARAGRAPH_NAMESPACES) {
            const comments = getElementsByTagNameNSAll(commentsDom, wNs, "comment");
            for (let i = 0; i < comments.length; i += 1) {
                const comment = comments[i];
                if (!comment) continue;
                const paragraphs = getElementsByTagNameNSAll(comment, wNs, "p");
                for (let j = 0; j < paragraphs.length; j += 1) {
                    const paragraph = paragraphs[j];
                    const paraId = paragraph?.getAttributeNS(W14_NAMESPACE, "paraId");
                    if (paraId) commentParaIds.add(paraId);
                }
            }
        }

        // Work out, without mutating, which parts this repair would write and
        // the durableIds commentsIds.xml would hold once its write lands: the
        // extensible part is aligned to that in-memory result, so a part
        // `writeXmlPart` would refuse must refuse the whole repair (#889).
        const commentIds = getElementsByTagNameNSAll(idsDom, W16CID_NAMESPACE, "commentId");
        const commentId = commentIds[0];
        let canonicalDurableId: string | null = null;
        let retargetParaId = false;
        let retargetDurableId = false;
        if (commentParaIds.size === 1 && commentIds.length === 1) {
            const targetParaId = [...commentParaIds][0];
            canonicalDurableId = targetParaId;
            const currentParaId = commentId?.getAttributeNS(W16CID_NAMESPACE, "paraId");
            retargetParaId =
                commentId !== undefined && !!currentParaId && currentParaId !== targetParaId && !commentParaIds.has(currentParaId);
            const currentDurableId = commentId?.getAttributeNS(W16CID_NAMESPACE, "durableId");
            retargetDurableId = commentId !== undefined && !!currentDurableId && currentDurableId !== canonicalDurableId;
        }

        const durableIds = new Set<string>();
        for (let i = 0; i < commentIds.length; i += 1) {
            const durableId = commentIds[i]?.getAttributeNS(W16CID_NAMESPACE, "durableId");
            if (durableId) durableIds.add(durableId);
        }
        // The set the extensible part aligns to once the ids write lands: the
        // same durableIds, with a retargeted single entry at the canonical
        // value the write would leave there.
        const writtenDurableIds = retargetDurableId && canonicalDurableId !== null ? new Set<string>([canonicalDurableId]) : durableIds;
        const alignedDurableId = writtenDurableIds.size === 1 ? [...writtenDurableIds][0] : null;

        let cexDom: Document | null = null;
        let cexEntry: Element | null = null;
        let retargetExtensible = false;
        if (commentsExtensibleXml && alignedDurableId !== null) {
            try {
                cexDom = parseXmlBytes(await fs.readFile(commentsExtensibleXml));
            } catch {
                cexDom = null;
            }
            if (cexDom) {
                const entries = getElementsByTagNameNSAll(cexDom, W16CEX_NAMESPACE, "commentExtensible");
                cexEntry = entries.length === 1 ? (entries[0] ?? null) : null;
                const currentDurableId = cexEntry?.getAttributeNS(W16CEX_NAMESPACE, "durableId");
                retargetExtensible =
                    cexEntry !== null &&
                    !!currentDurableId &&
                    currentDurableId !== alignedDurableId &&
                    !writtenDurableIds.has(currentDurableId);
            }
        }

        // Preflight: decode and check writability of every part before mutating
        // any DOM, so one refused write cannot split references that agreed
        // before it ran. A part the repair would leave alone does not block
        // it (#889).
        const writeSet: string[] = [];
        if (retargetParaId || retargetDurableId) writeSet.push(commentsIdsXml);
        if (retargetExtensible && commentsExtensibleXml) writeSet.push(commentsExtensibleXml);
        for (const part of writeSet) {
            if (decodeXmlPart(await fs.readFile(part)) === undefined) return 0;
            try {
                await fs.access(part, fsConstants.W_OK);
            } catch {
                return 0;
            }
        }

        // Counted only after the write: a part whose encoding `writeXmlPart`
        // refuses is left untouched (#880).
        let repairs = 0;
        if (commentId && canonicalDurableId !== null) {
            if (retargetParaId) {
                commentId.setAttributeNS(W16CID_NAMESPACE, "w16cid:paraId", canonicalDurableId);
            }
            if (retargetDurableId) {
                commentId.setAttributeNS(W16CID_NAMESPACE, "w16cid:durableId", canonicalDurableId);
            }
        }
        const idsRepairs = (retargetParaId ? 1 : 0) + (retargetDurableId ? 1 : 0);
        if (idsRepairs > 0) {
            let written: boolean;
            try {
                written = await writeXmlPart(commentsIdsXml, idsDom);
            } catch {
                return repairs;
            }
            if (!written) return repairs;
            repairs += idsRepairs;
        }
        if (retargetExtensible && cexDom && cexEntry && commentsExtensibleXml && alignedDurableId !== null) {
            cexEntry.setAttributeNS(W16CEX_NAMESPACE, "w16cex:durableId", alignedDurableId);
            let written: boolean;
            try {
                written = await writeXmlPart(commentsExtensibleXml, cexDom);
            } catch {
                return repairs;
            }
            if (written) repairs += 1;
        }
        return repairs;
    }

    /**
     * Inject canonical definitions into each story root's styles part —
     * the main document's (by name in its folder when none is related,
     * #699), and a reached glossary's own for the references its
     * story makes (#698) — for every well-known style ID referenced by
     * document XML but not defined.
     * Only handles the small set of styles writers commonly leak —
     * unknown style IDs are NOT auto-defined (we don't know what they
     * mean). Pairs with `validateStyleReferences`.
     */
    async repairMissingStyleDefinitions(): Promise<number> {
        const stylesPath = await this.mainStylesPath();
        const glossaryStyles = await this.glossaryStoryStylesFiles();

        // Collect each story's referenced IDs against the styles part its
        // root relates (#698): a reached glossary's references belong to
        // the glossary's own styles part — or are unrepaired when it
        // relates none — never the main document's. A part both roots
        // reach stays with the main document.
        const referencedByStyles = new Map<string, Set<string>>();
        if (stylesPath) referencedByStyles.set(stylesPath, new Set());
        const refTags = ["pStyle", "rStyle", "tblStyle", "numStyle", "linkedStyle"] as const;
        for (const xmlFile of this.userTextXmlFiles()) {
            const styles = glossaryStyles.has(xmlFile) ? glossaryStyles.get(xmlFile) : stylesPath;
            if (!styles) continue;
            try {
                const dom = parseXmlBytes(await fs.readFile(xmlFile));
                const referenced = referencedByStyles.get(styles) ?? new Set<string>();
                for (const tag of refTags) {
                    for (const ns of WORD_PARAGRAPH_NAMESPACES) {
                        const list = getElementsByTagNameNSAll(dom, ns, tag);
                        for (let i = 0; i < list.length; i += 1) {
                            const elem = list[i];
                            if (!elem) continue;
                            const val = elem.getAttributeNS(ns, "val");
                            if (val) referenced.add(val);
                        }
                    }
                }
                referencedByStyles.set(styles, referenced);
            } catch {
                // pass
            }
        }

        let repairs = 0;
        for (const [styles, referenced] of referencedByStyles) {
            // Two sources of "needs to be defined": styles referenced by
            // the story's XML, and the four ECMA-376 implied-defaults the
            // main document's styles part must always carry.
            const needed = new Set<string>(referenced);
            if (styles === stylesPath) {
                for (const def of REQUIRED_DEFAULT_STYLES) needed.add(def.styleId);
            }
            repairs += await this.writeMissingStyleDefinitions(styles, needed);
        }
        return repairs;
    }

    /**
     * Write the `WELL_KNOWN_STYLE_DEFINITIONS` fragment of every `needed`
     * styleId `stylesPath` does not define into it; returns the count.
     */
    private async writeMissingStyleDefinitions(stylesPath: string, needed: ReadonlySet<string>): Promise<number> {
        let stylesDom: Document;
        try {
            stylesDom = parseXmlBytes(await fs.readFile(stylesPath));
        } catch {
            return 0;
        }
        const stylesRoot = stylesDom.documentElement;
        if (!stylesRoot) return 0;

        const defined = new Set<string>();
        for (const ns of WORD_PARAGRAPH_NAMESPACES) {
            const list = getElementsByTagNameNSAll(stylesDom, ns, "style");
            for (let i = 0; i < list.length; i += 1) {
                const elem = list[i];
                if (!elem) continue;
                const id = elem.getAttributeNS(ns, "styleId");
                if (id) defined.add(id);
            }
        }

        const missing = [...needed].filter((id) => !defined.has(id));
        if (missing.length === 0) return 0;

        let repairs = 0;
        for (const id of missing) {
            const fragment = WELL_KNOWN_STYLE_DEFINITIONS[id];
            if (!fragment) continue;
            // Append the fragment as a child of <w:styles>.
            const actualNamespace = stylesRoot.namespaceURI ?? WORD_2006_NAMESPACE;
            const tmp = parseXml(`<w:styles xmlns:w="${actualNamespace}">${fragment}</w:styles>`);
            const tmpRoot = tmp.documentElement;
            if (!tmpRoot) continue;
            const child = tmpRoot.firstChild;
            if (!child) continue;
            const imported = stylesDom.importNode(child, true);
            stylesRoot.appendChild(imported);
            repairs += 1;
        }

        if (repairs > 0 && !(await writeXmlPart(stylesPath, stylesDom))) {
            // The part's encoding cannot be written back: nothing was repaired (#880).
            return 0;
        }
        return repairs;
    }

    /**
     * Stamp a fresh in-range `w14:paraId` (and the matching
     * `w14:textId`) on every `<w:p>` and `<w:tr>` that needs one. The
     * new paraIds are random and avoid colliding with any paraId
     * already present on either side (incl. the renumbered ones from
     * `repairParaId`); textIds are sampled from the same space and
     * likewise deduped within their own pool (paraId and textId
     * namespaces are independent in Word's model — Word's saves often
     * emit the placeholder `77777777` for textId, but the spec only
     * requires it be a hex digit string).
     *
     * Four cases:
     *   1. Element has neither → stamp both (Word's tracked-changes
     *      infrastructure expects them paired).
     *   2. Element has paraId but no textId → stamp textId. This is
     *      the asymmetric shape that triggers Word's "Document
     *      Recovery — Table Properties" dialog on open.
     *   3. Element has textId but no paraId → stamp paraId (the
     *      inverse asymmetric shape — unusual but handled for
     *      completeness so Word doesn't see a dangling textId).
     *   4. Element has both → leave alone.
     */
    async repairMissingParaIds(): Promise<number> {
        let repairs = 0;
        const usedParaIds = new Set<string>();
        const usedTextIds = new Set<string>();

        // Parse every user-text XML file once; collect existing IDs and
        // store parsed DOMs for the stamp pass so we avoid a re-parse.
        const domByFile = new Map<string, Document>();
        for (const xmlFile of this.userTextXmlFiles()) {
            try {
                const dom = parseXmlBytes(await fs.readFile(xmlFile));
                domByFile.set(xmlFile, dom);
                const all = getElementsByTagNameAll(dom);
                for (let i = 0; i < all.length; i += 1) {
                    const elem = all[i];
                    if (!elem) continue;
                    const paraId = elem.getAttributeNS(W14_NAMESPACE, "paraId");
                    if (paraId) usedParaIds.add(paraId.toUpperCase());
                    const textId = elem.getAttributeNS(W14_NAMESPACE, "textId");
                    if (textId) usedTextIds.add(textId.toUpperCase());
                }
            } catch {
                // pass
            }
        }

        let allocCounter = 1;
        const allocate = (pool: Set<string>): string => {
            for (;;) {
                const newId = allocCounter.toString(16).toUpperCase().padStart(8, "0");
                allocCounter += 1;
                if (!pool.has(newId)) {
                    pool.add(newId);
                    return newId;
                }
            }
        };

        // Stamp pass: reuse the DOMs we already parsed.
        for (const [xmlFile, dom] of domByFile) {
            try {
                let modified = false;
                // Counted only after the write: a part whose encoding
                // `writeXmlPart` refuses is left untouched (#880).
                let partRepairs = 0;
                for (const local of ["p", "tr"] as const) {
                    for (const ns of WORD_PARAGRAPH_NAMESPACES) {
                        const list = getElementsByTagNameNSAll(dom, ns, local);
                        for (let i = 0; i < list.length; i += 1) {
                            const elem = list[i];
                            if (!elem) continue;
                            const paraId = elem.getAttributeNS(W14_NAMESPACE, "paraId");
                            const textId = elem.getAttributeNS(W14_NAMESPACE, "textId");
                            if (!paraId && !textId) {
                                // Case 1: stamp both as a pair.
                                elem.setAttributeNS(W14_NAMESPACE, "w14:paraId", allocate(usedParaIds));
                                elem.setAttributeNS(W14_NAMESPACE, "w14:textId", allocate(usedTextIds));
                                partRepairs += 2;
                                modified = true;
                            } else if (paraId && !textId) {
                                // Case 2: asymmetric — stamp textId.
                                elem.setAttributeNS(W14_NAMESPACE, "w14:textId", allocate(usedTextIds));
                                partRepairs += 1;
                                modified = true;
                            } else if (!paraId && textId) {
                                // Case 3: has textId but no paraId — stamp paraId.
                                elem.setAttributeNS(W14_NAMESPACE, "w14:paraId", allocate(usedParaIds));
                                partRepairs += 1;
                                modified = true;
                            }
                            // Case 4: both present — nothing to do.
                        }
                    }
                }
                if (modified) {
                    // The stamps are w14 attributes: declare w14 on the root and
                    // list it in mc:Ignorable, as Word writes it, so the repair
                    // leaves no non-ignorable extension behind (#653).
                    if (dom.documentElement) declareIgnorableW14(dom.documentElement);
                    if (await writeXmlPart(xmlFile, dom)) {
                        repairs += partRepairs;
                    }
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
                const dom = parseXmlBytes(await fs.readFile(xmlFile));
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
                    const attr = attrs[i];
                    if (!attr) continue;
                    if (!attr.name.endsWith("Ignorable")) continue;
                    ignorableAttrs.push({ name: attr.name, value: attr.value });
                }
                if (ignorableAttrs.length === 0) continue;

                // Counted only after the write: a part whose encoding
                // `writeXmlPart` refuses is left untouched (#880).
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
                        } else {
                            // Truly unknown — drop from Ignorable rather than
                            // guessing a URI.
                            perFileMutations += 1;
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

                if (perFileMutations > 0 && (await writeXmlPart(xmlFile, dom))) {
                    repairs += perFileMutations;
                }
            } catch {
                // swallow — corrupted XML files surface elsewhere
            }
        }
        return repairs;
    }

    /**
     * Repair every over-cap `w16cid:durableId` (`>= 0x7FFFFFFF`) and
     * propagate the renumbering into the threaded-comments extension, as
     * `repairParaId` does for paraIds.
     *
     * Two-pass implementation:
     *
     *  Pass 1: scan all XML files, build a `Map<oldId, newId>` for every
     *          over-cap hex `w16cid:durableId` outside `numbering.xml`.
     *          Each old value gets one fresh in-range value that collides
     *          with no durableId in use — `w16cid` anywhere outside
     *          `numbering.xml`, and `w16cex` in `commentsExtensible.xml` —
     *          so two over-cap ids can't land on the same fresh value and
     *          none can land on a value an extensible entry already names.
     *  Pass 2: walk every XML file again and rewrite (a) every
     *          `w16cid:durableId` named in the map, AND (b) every
     *          `w16cex:durableId` in any part that names a remapped value.
     *          Without (b), `commentsExtensible.xml`'s references still
     *          name the old values and every extensible entry is left
     *          orphaned — surfaced by `validateCommentThreading`'s
     *          `comment-thread-durableid-orphan` check.
     *
     * The map is keyed by the upper-cased value: both attributes are
     * `xsd:hexBinary`, whose value ignores the case of its hex letters, so
     * `ffffffff` in one part and `FFFFFFFF` in the other name one id and get
     * one new value (#904).
     *
     * `numbering.xml` is unchanged by the remap: its decimal durableIds
     * are self-contained (no part references them), so each keeps its own
     * per-element renumbering, independent of the preflight below.
     * Every part the hex rewrite would write must be decodable and writable;
     * otherwise only the decimal renumbering runs. A failed or refused write
     * of a part carrying a hex rename stops the remaining hex renames and
     * counts only the repairs already written (#909, #920). A failed or
     * refused write of a part carrying only the decimal renumbering counts
     * nothing and stops nothing, and that renumbering runs whether the hex
     * rewrite is refused by the preflight or stopped by a failed write
     * (#925).
     */
    async repairDurableId(): Promise<number> {
        let repairs = 0;
        const remap = new Map<string, string>();
        const usedDurableIds = new Set<string>();
        const overCapIds = new Set<string>();

        // Pass 1: discover every over-cap hex durableId and every value
        // already in use, so a fresh value collides with neither.
        const parsedDoms = new Map<string, Document>();
        for (const xmlFile of this.xmlFiles) {
            try {
                const dom = parseXmlBytes(await fs.readFile(xmlFile));
                parsedDoms.set(xmlFile, dom);
                const base = baseName(xmlFile);
                const all = getElementsByTagNameAll(dom);
                for (let i = 0; i < all.length; i += 1) {
                    const elem = all[i];
                    if (!elem) continue;
                    if (base !== "numbering.xml") {
                        const durableId = elem.getAttributeNS(W16CID_NAMESPACE, "durableId");
                        if (durableId) {
                            usedDurableIds.add(durableId.toUpperCase());
                            const v = parseIdValue(durableId, 16);
                            if (Number.isNaN(v) || v >= MAX_DURABLE_ID) overCapIds.add(durableId.toUpperCase());
                        }
                    }
                    if (base === "commentsExtensible.xml") {
                        const durableId = elem.getAttributeNS(W16CEX_NAMESPACE, "durableId");
                        if (durableId) usedDurableIds.add(durableId.toUpperCase());
                    }
                }
            } catch {
                // pass — corrupted XML files surface elsewhere.
            }
        }

        // One fresh in-range value per distinct over-cap id; the same old
        // value always maps to the same new value.
        for (const oldId of overCapIds) {
            for (;;) {
                const newId = nextSecureLongHexNumber().toString(16).toUpperCase().padStart(8, "0");
                if (!usedDurableIds.has(newId)) {
                    usedDurableIds.add(newId);
                    remap.set(oldId, newId);
                    break;
                }
            }
        }

        // Preflight: the same rename must land in every part that
        // references it, so decode every part the hex rewrite would write
        // before mutating any DOM. A part the rewrite would leave alone
        // does not block it; one it cannot decode or write refuses
        // the whole hex rewrite — the numbering renumbering below still
        // runs — rather than splitting references that agreed before it
        // ran (#889).
        let hexRefused = false;
        if (remap.size > 0) {
            for (const [xmlFile, dom] of parsedDoms) {
                const base = baseName(xmlFile);
                let wouldWrite = false;
                const all = getElementsByTagNameAll(dom);
                for (let i = 0; i < all.length; i += 1) {
                    const elem = all[i];
                    if (!elem) continue;
                    const w16cid = elem.getAttributeNS(W16CID_NAMESPACE, "durableId");
                    if (w16cid && base !== "numbering.xml" && remap.has(w16cid.toUpperCase())) wouldWrite = true;
                    const w16cex = elem.getAttributeNS(W16CEX_NAMESPACE, "durableId");
                    if (w16cex && remap.has(w16cex.toUpperCase())) wouldWrite = true;
                    if (wouldWrite) break;
                }
                if (!wouldWrite) continue;
                if (decodeXmlPart(await fs.readFile(xmlFile)) === undefined) {
                    hexRefused = true;
                    break;
                }
                try {
                    await fs.access(xmlFile, fsConstants.W_OK);
                } catch {
                    hexRefused = true;
                    break;
                }
            }
        }

        // Pass 2: rewrite all references; numbering.xml keeps its own
        // self-contained decimal renumbering.
        let hexStopped = false;
        for (const [xmlFile, dom] of parsedDoms) {
            const base = baseName(xmlFile);
            let modified = false;
            // Counted only after the write: a part whose encoding
            // `writeXmlPart` refuses is left untouched (#880).
            let partRepairs = 0;
            // A part the hex rewrite renames: its failed write stops the
            // remaining hex renames, like the preflight refusal (#909, #920).
            let hexPart = false;
            const all = getElementsByTagNameAll(dom);
            for (let i = 0; i < all.length; i += 1) {
                const elem = all[i];
                if (!elem) continue;
                const cidDurableId = elem.getAttributeNS(W16CID_NAMESPACE, "durableId");
                if (cidDurableId && base === "numbering.xml") {
                    const v = parseIdValue(cidDurableId, 10);
                    if (Number.isNaN(v) || v >= MAX_DURABLE_ID) {
                        // setAttributeNS keeps the prefix binding intact.
                        elem.setAttributeNS(W16CID_NAMESPACE, "w16cid:durableId", String(nextSecureLongHexNumber()));
                        partRepairs += 1;
                        modified = true;
                    }
                } else if (cidDurableId && !hexRefused && !hexStopped && remap.has(cidDurableId.toUpperCase())) {
                    // setAttributeNS keeps the prefix binding intact.
                    elem.setAttributeNS(W16CID_NAMESPACE, "w16cid:durableId", remap.get(cidDurableId.toUpperCase())!);
                    partRepairs += 1;
                    modified = true;
                    hexPart = true;
                }
                const cexDurableId = elem.getAttributeNS(W16CEX_NAMESPACE, "durableId");
                if (cexDurableId && !hexRefused && !hexStopped && remap.has(cexDurableId.toUpperCase())) {
                    // setAttributeNS keeps the prefix binding intact.
                    elem.setAttributeNS(W16CEX_NAMESPACE, "w16cex:durableId", remap.get(cexDurableId.toUpperCase())!);
                    partRepairs += 1;
                    modified = true;
                    hexPart = true;
                }
            }
            if (modified) {
                let written: boolean;
                try {
                    written = await writeXmlPart(xmlFile, dom);
                } catch {
                    // A failed write of a hex part stops the remaining hex
                    // renames; a decimal-only write is skipped, as numbering's
                    // renumbering is independent of the hex rewrite (#925).
                    if (hexPart) hexStopped = true;
                    continue;
                }
                if (!written) {
                    if (hexPart) hexStopped = true;
                    continue;
                }
                repairs += partRepairs;
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
     *
     * Every part to be written must be decodable and writable before any
     * write starts. A failed or refused write returns only the repairs already
     * written (#909, #920).
     */
    async repairParaId(): Promise<number> {
        let repairs = 0;
        const remap = new Map<string, string>();
        // textId is not cross-referenced (unlike paraId, which appears in
        // w15:paraId / w15:paraIdParent in commentsExtended.xml), so it gets
        // its own remap to keep the value pools cleanly separated.
        const textIdRemap = new Map<string, string>();
        // Separate pools for paraId and textId so that the two value
        // namespaces remain independent (Word's model tolerates, and
        // routinely produces, the same hex value appearing as both a
        // paraId on one element and a textId on another).
        const usedNewParaIds = new Set<string>();
        const usedNewTextIds = new Set<string>();

        let paraIdCounter = 1;
        const allocateNewParaId = (): string => {
            for (;;) {
                const newId = paraIdCounter.toString(16).toUpperCase().padStart(8, "0");
                paraIdCounter += 1;
                if (!usedNewParaIds.has(newId)) {
                    usedNewParaIds.add(newId);
                    return newId;
                }
            }
        };

        let textIdCounter = 1;
        const allocateNewTextId = (): string => {
            for (;;) {
                const newId = textIdCounter.toString(16).toUpperCase().padStart(8, "0");
                textIdCounter += 1;
                if (!usedNewTextIds.has(newId)) {
                    usedNewTextIds.add(newId);
                    return newId;
                }
            }
        };

        // Pass 1: discover every over-cap w14:paraId / w14:textId and pick
        // replacements. Both share the < 0x80000000 cap (ST_LongHexNumber).
        // Existing in-range IDs are collected into the per-type pools so
        // that replacements can't collide with values already in use.
        const parsedDoms = new Map<string, Document>();
        for (const xmlFile of this.xmlFiles) {
            try {
                const dom = parseXmlBytes(await fs.readFile(xmlFile));
                parsedDoms.set(xmlFile, dom);
                const all = getElementsByTagNameAll(dom);
                for (let i = 0; i < all.length; i += 1) {
                    const elem = all[i];
                    if (!elem) continue;
                    const paraId = elem.getAttributeNS(W14_NAMESPACE, "paraId");
                    if (paraId && !remap.has(paraId)) {
                        const v = parseIdValue(paraId, 16);
                        if (Number.isNaN(v) || v >= MAX_PARA_ID) {
                            remap.set(paraId, allocateNewParaId());
                        } else {
                            usedNewParaIds.add(paraId.toUpperCase());
                        }
                    }
                    const textId = elem.getAttributeNS(W14_NAMESPACE, "textId");
                    if (textId && !textIdRemap.has(textId)) {
                        const v = parseIdValue(textId, 16);
                        if (Number.isNaN(v) || v >= MAX_PARA_ID) {
                            textIdRemap.set(textId, allocateNewTextId());
                        } else {
                            usedNewTextIds.add(textId.toUpperCase());
                        }
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
            const exts = getElementsByTagNameNSAll(dom, W15_NAMESPACE, "commentEx");
            for (let i = 0; i < exts.length; i += 1) {
                const elem = exts[i];
                if (!elem) continue;
                for (const attr of ["paraId", "paraIdParent"] as const) {
                    const val = elem.getAttributeNS(W15_NAMESPACE, attr);
                    if (!val) continue;
                    if (remap.has(val)) continue;
                    const v = parseIdValue(val, 16);
                    if (Number.isNaN(v) || v >= MAX_PARA_ID) {
                        remap.set(val, allocateNewParaId());
                    } else {
                        usedNewParaIds.add(val.toUpperCase());
                    }
                }
            }
        }

        if (remap.size === 0 && textIdRemap.size === 0) return 0;

        // Preflight: the same rename must land in every part that references
        // it, so decode every part pass 2 would rewrite before mutating any
        // DOM. A part the pass would leave alone does not block the repair;
        // one it cannot decode or write refuses all of it, rather than
        // splitting references that agreed before it ran (#889).
        for (const [xmlFile, dom] of parsedDoms) {
            let wouldWrite = false;
            const all = getElementsByTagNameAll(dom);
            for (let i = 0; i < all.length; i += 1) {
                const elem = all[i];
                if (!elem) continue;
                const w14Para = elem.getAttributeNS(W14_NAMESPACE, "paraId");
                if (w14Para && remap.has(w14Para)) wouldWrite = true;
                const w14Text = elem.getAttributeNS(W14_NAMESPACE, "textId");
                if (w14Text && textIdRemap.has(w14Text)) wouldWrite = true;
                const w15Para = elem.getAttributeNS(W15_NAMESPACE, "paraId");
                if (w15Para && remap.has(w15Para)) wouldWrite = true;
                const w15Parent = elem.getAttributeNS(W15_NAMESPACE, "paraIdParent");
                if (w15Parent && remap.has(w15Parent)) wouldWrite = true;
                if (wouldWrite) break;
            }
            if (!wouldWrite) continue;
            if (decodeXmlPart(await fs.readFile(xmlFile)) === undefined) return 0;
            try {
                await fs.access(xmlFile, fsConstants.W_OK);
            } catch {
                return 0;
            }
        }

        // Pass 2: rewrite all references.
        for (const [xmlFile, dom] of parsedDoms) {
            let modified = false;
            // Counted only after the write: a part whose encoding
            // `writeXmlPart` refuses is left untouched (#880).
            let partRepairs = 0;
            const all = getElementsByTagNameAll(dom);
            for (let i = 0; i < all.length; i += 1) {
                const elem = all[i];
                if (!elem) continue;
                const w14Para = elem.getAttributeNS(W14_NAMESPACE, "paraId");
                if (w14Para && remap.has(w14Para)) {
                    elem.setAttributeNS(W14_NAMESPACE, "w14:paraId", remap.get(w14Para)!);
                    partRepairs += 1;
                    modified = true;
                }
                const w14Text = elem.getAttributeNS(W14_NAMESPACE, "textId");
                if (w14Text && textIdRemap.has(w14Text)) {
                    elem.setAttributeNS(W14_NAMESPACE, "w14:textId", textIdRemap.get(w14Text)!);
                    partRepairs += 1;
                    modified = true;
                }
                const w15Para = elem.getAttributeNS(W15_NAMESPACE, "paraId");
                if (w15Para && remap.has(w15Para)) {
                    elem.setAttributeNS(W15_NAMESPACE, "w15:paraId", remap.get(w15Para)!);
                    partRepairs += 1;
                    modified = true;
                }
                const w15Parent = elem.getAttributeNS(W15_NAMESPACE, "paraIdParent");
                if (w15Parent && remap.has(w15Parent)) {
                    elem.setAttributeNS(W15_NAMESPACE, "w15:paraIdParent", remap.get(w15Parent)!);
                    partRepairs += 1;
                    modified = true;
                }
            }
            if (modified) {
                let written: boolean;
                try {
                    written = await writeXmlPart(xmlFile, dom);
                } catch {
                    return repairs;
                }
                if (!written) return repairs;
                repairs += partRepairs;
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

    private async tableStyleBordersByStyleId(stylesPath: string | null): Promise<Map<string, BorderSignature>> {
        const out = new Map<string, BorderSignature>();
        if (!stylesPath) return out;

        let dom: Document;
        try {
            dom = await this.readPartDom(stylesPath);
        } catch {
            return out;
        }

        for (const ns of WORD_PARAGRAPH_NAMESPACES) {
            const styles = getElementsByTagNameNSAll(dom, ns, "style");
            for (let i = 0; i < styles.length; i += 1) {
                const style = styles[i];
                if (!style) continue;
                if (style.getAttributeNS(ns, "type") !== "table") continue;
                const styleId = style.getAttributeNS(ns, "styleId");
                if (!styleId) continue;
                const tblPr = directChild(style, ns, "tblPr");
                const signature = borderSignature(tblPr ? directChild(tblPr, ns, "tblBorders") : null, ns);
                if (signature) out.set(styleId, signature);
            }
        }
        return out;
    }

    /**
     * The user-text parts each reached glossary's story is judged against its
     * own styles part for (#698): a reached glossary document — a
     * `glossaryDocument` target of the main document's relationships — and
     * every part only the glossary's own relationships reach
     * (`glossaryTypedRelationshipTargets`, `via: "glossary"`), never the main
     * document itself, wherever it is stored, nor a part the main document
     * also reaches, which stays the main document's (`wordMainReachedParts`:
     * every relationship's target, whatever its type, so a part the main
     * document reaches through a non-story type cannot be claimed through a
     * story one, #742). Each maps to
     * the styles part that root relates; `undefined` when the root relates none
     * — or relates one that is not XML — so the part is not judged. An empty
     * map — no reached glossary, or a package view that cannot be read — keeps
     * the by-name behaviour: every part judged against the main styles part.
     */
    private async glossaryStoryStylesFiles(): Promise<Map<string, string | undefined>> {
        const files = new Map<string, string | undefined>();
        try {
            const context = this.context();
            const roots = await wordTypedRelationshipTargets(context, new Set(["glossaryDocument"]), this.profile);
            if (roots.length === 0) return files;
            const stylesByRoot = new Map<string, string>();
            for (const target of await glossaryTypedRelationshipTargets(context, new Set(["styles"]), this.profile)) {
                if (target.root !== undefined && !stylesByRoot.has(canonicalPartName(target.root))) {
                    stylesByRoot.set(canonicalPartName(target.root), target.part);
                }
            }
            // Canonical part name to its root's canonical name, and to the
            // part spelling `xmlFileOf` resolves by OPC equivalence.
            const rootByPart = new Map<string, string>();
            const storedByPart = new Map<string, string>();
            for (const root of roots) {
                rootByPart.set(canonicalPartName(root.part), canonicalPartName(root.part));
                storedByPart.set(canonicalPartName(root.part), root.part);
            }
            // The main document is never a glossary's story part, wherever it
            // is stored, and neither is any part the main document's own
            // relationships reach, whatever their type (#742).
            const mainReached = await wordMainReachedParts(context, this.profile);
            for (const target of await glossaryTypedRelationshipTargets(context, STORY_TEXT_SEGMENTS, this.profile)) {
                const part = canonicalPartName(target.part);
                if (target.root === undefined || rootByPart.has(part) || mainReached.has(part)) continue;
                rootByPart.set(part, canonicalPartName(target.root));
                storedByPart.set(part, target.part);
            }
            for (const [part, root] of rootByPart) {
                const file = this.xmlFileOf([storedByPart.get(part)!]);
                if (file === undefined) continue;
                const styles = stylesByRoot.get(root);
                files.set(file, styles === undefined ? undefined : this.xmlFileOf([styles]));
            }
        } catch {
            // An unreadable package view is reported by its own checks; every
            // part keeps the by-name judgement against the main styles part.
        }
        return files;
    }

    /**
     * The `w:styleId` values a styles part defines: the empty set when
     * `stylesPath` is `null` — no styles part was found, so every reference
     * dangles — and `undefined` when the part cannot be parsed, which the
     * XML diagnostics already report.
     */
    private async definedStyleIds(stylesPath: string | null): Promise<ReadonlySet<string> | undefined> {
        const defined = new Set<string>();
        if (stylesPath === null) return defined;
        let dom: Document;
        try {
            dom = await this.readPartDom(stylesPath);
        } catch {
            // A malformed styles part is covered by the xml-syntax check.
            return undefined;
        }
        for (const ns of WORD_PARAGRAPH_NAMESPACES) {
            const list = getElementsByTagNameNSAll(dom, ns, "style");
            for (let i = 0; i < list.length; i += 1) {
                const elem = list[i];
                if (!elem) continue;
                const id = elem.getAttributeNS(ns, "styleId");
                if (id) defined.add(id);
            }
        }
        return defined;
    }
}

// ===== module-level helpers ===================================================

type BorderSide = (typeof TABLE_BORDER_SIDES)[number];
type BorderSignature = ReadonlyMap<BorderSide, ReadonlyMap<string, string>>;

function finalize(issues: ValidationIssue[]): ValidationResult {
    return { valid: issues.every((i) => i.severity !== "error"), issues };
}

/**
 * Strict applies markup compatibility twice: the Word rules judge Word's view
 * and the schema walk judges the schema's view. A finding both views make is
 * reported once, by the Word rule; the schema walk keeps what only its view
 * finds (a namespace Word understands but the schemas do not).
 */
export function withoutRepeatedCompatibilityFindings(schema: ValidationResult, rules: ValidationResult): ValidationResult {
    // The schema walk's path is path.relative output (backslashes on Windows);
    // the Word rule's is the normalized part name.
    const key = (issue: ValidationIssue): string => `${issue.code}\u0000${(issue.path ?? "").replace(/\\/g, "/")}\u0000${issue.message}`;
    const reported = new Map<string, number>();
    for (const issue of rules.issues) {
        if (!issue.code?.startsWith("mce-")) continue;
        reported.set(key(issue), (reported.get(key(issue)) ?? 0) + 1);
    }
    const issues = schema.issues.filter((issue) => {
        const remaining = reported.get(key(issue)) ?? 0;
        if (remaining === 0) return true;
        reported.set(key(issue), remaining - 1);
        return false;
    });
    return { valid: schema.valid, issues };
}

/**
 * A full non-lenient run reports each story `w:headerReference`/
 * `w:footerReference` the `word-header-footer-relationship` rule judged
 * once — through that rule — so the same run's `validateAllRelationshipIds`
 * result drops its `rels-id-mismatch` copy, and for an `r:id` naming no
 * Relationship its `rels-id-missing` copy, of each reported element (#612,
 * #619); a dangling Transitional `r:id` on a `c:chart` or an
 * `o:OLEObject` pairs the same way with `word-part-reference-relationship`
 * (#705, sessions 45 and 46), and a dangling `wne:toolbarData` `r:id` with
 * `word-attached-toolbar-relationship` (#716, session 48). The pairing is the structured identity both producers registered on
 * the issue object when they created it (#616): part path (normalized on
 * registration, as `withoutRepeatedCompatibilityFindings` compares paths),
 * element namespace URI, local name and the raw `r:id` all match. The
 * namespace URI the `r:id` was read in is deliberately not part of the
 * pairing (#727): the Word rule reads the markup-compatibility view and the
 * legacy walk the raw part, so an element with an `mc:Choice` and an
 * `mc:Fallback` twin yields two legacy copies of the one defect the Word
 * rule reports once, and a Strict-namespace `r:id` on the selected twin
 * (review BE F1) is the same element name's other spelling. The key is the
 * same element name — namespace URI and local name — with the same Id in
 * the same part, not the element instance: every legacy copy of it is
 * dropped, the unselected twin's copy with it, and no finding is dropped
 * for another part, another Id or another element name. The former
 * message-prose pairing let a quoted `r:id`
 * double the report and a crafted `r:id` consume a foreign element's only
 * mismatch. A finding without a registered identity — another producer, or
 * a copy made between producer and helper — is never dropped. Where the
 * Word rule stays silent — a foreign-namespace `x:headerReference`, a
 * glossary no main-document relationship reaches — the legacy findings stay,
 * and `validateAllRelationshipIds` called on its own reports every one.
 */
export function withoutRepeatedHeaderFooterFindings(relationshipIds: ValidationResult, rules: ValidationResult): ValidationResult {
    const key = (identity: StoryReferenceIdentity): string =>
        `${identity.part}\u0000${identity.namespaceURI}\u0000${identity.localName}\u0000${identity.rid}`;
    const reported = new Set<string>();
    for (const issue of rules.issues) {
        if (
            issue.code !== WORD_HEADER_FOOTER_RELATIONSHIP &&
            issue.code !== WORD_PART_REFERENCE_RELATIONSHIP &&
            issue.code !== WORD_ATTACHED_TOOLBAR_RELATIONSHIP
        )
            continue;
        const identity = storyReferenceIdentity(issue);
        if (!identity) continue;
        reported.add(key(identity));
    }
    const issues = relationshipIds.issues.filter((issue) => {
        if (issue.code !== "rels-id-mismatch" && issue.code !== "rels-id-missing") return true;
        const identity = storyReferenceIdentity(issue);
        if (!identity) return true;
        return !reported.has(key(identity));
    });
    return { valid: relationshipIds.valid, issues };
}

function directChild(parent: Element, namespaceURI: string, local: string): Element | null {
    for (let child = parent.firstChild; child; child = child.nextSibling) {
        if (child.nodeType !== 1) continue;
        const elem = child as Element;
        if (elem.namespaceURI === namespaceURI && (elem.localName || elem.tagName.split(":").pop()) === local) return elem;
    }
    return null;
}

function directChildren(parent: Element, namespaceURI: string, local: string): Element[] {
    const out: Element[] = [];
    for (let child = parent.firstChild; child; child = child.nextSibling) {
        if (child.nodeType !== 1) continue;
        const elem = child as Element;
        if (elem.namespaceURI === namespaceURI && (elem.localName || elem.tagName.split(":").pop()) === local) {
            out.push(elem);
        }
    }
    return out;
}

/**
 * Declare `w14` on a part root and list it in the root's `mc:Ignorable`,
 * reusing the prefix the root binds to the compatibility namespace, else
 * declaring `mc`. A root that binds `w14` or `mc` to another namespace is
 * left as it is (#653).
 */
function declareIgnorableW14(root: Element): void {
    const bindings = new Map<string, string>();
    for (let i = 0; i < root.attributes.length; i += 1) {
        const attr = root.attributes[i];
        if (attr?.name.startsWith("xmlns:")) bindings.set(attr.name.slice("xmlns:".length), attr.value);
    }
    const w14 = bindings.get("w14");
    if (w14 !== undefined && w14 !== W14_NAMESPACE) return;
    let mc = [...bindings].find(([, uri]) => uri === NS.MC)?.[0];
    if (mc === undefined) {
        if (bindings.has("mc")) return;
        mc = "mc";
        root.setAttributeNS(XMLNS_NAMESPACE, "xmlns:mc", NS.MC);
    }
    // A namespace-qualified declaration, so the serializer sees w14 in scope
    // and stops declaring it again on each stamped element.
    if (w14 === undefined) root.setAttributeNS(XMLNS_NAMESPACE, "xmlns:w14", W14_NAMESPACE);
    const ignorable = root.getAttributeNodeNS(NS.MC, "Ignorable");
    const tokens = (ignorable?.value ?? "").split(/\s+/).filter(Boolean);
    if (tokens.includes("w14")) return;
    root.setAttributeNS(NS.MC, ignorable?.name ?? `${mc}:Ignorable`, [...tokens, "w14"].join(" "));
}

function isExternalRelationship(rel: Element, target: string): boolean {
    if (rel.getAttribute("TargetMode") === "External") return true;
    return /^[a-z][a-z\d+.-]*:/i.test(target);
}

function resolveRelationshipTargetPath(unpackedDir: string, relsFile: string, target: string): string | null {
    const targetWithoutFragment = target.split("#", 1)[0];
    if (!targetWithoutFragment) return null;

    let resolved: string;
    if (targetWithoutFragment.startsWith("/")) {
        resolved = path.resolve(unpackedDir, targetWithoutFragment.replace(/^\/+/, ""));
    } else {
        const relsDir = path.dirname(relsFile);
        // `_RELS` is the `_rels` folder: part names compare case-insensitively
        // (#346), canonically in fact, so an escaped spelling of the folder
        // resolves its targets against the source part's folder too (#781).
        const baseDir = isRelationshipsFolderName(path.basename(relsDir)) ? path.dirname(relsDir) : relsDir;
        resolved = path.resolve(baseDir, targetWithoutFragment);
    }

    return containRelationshipTarget(unpackedDir, resolved);
}

function borderSignature(borders: Element | null, namespaceURI: string): BorderSignature | null {
    if (!borders) return null;
    const out = new Map<BorderSide, ReadonlyMap<string, string>>();
    for (const side of TABLE_BORDER_SIDES) {
        const sideElem = directChild(borders, namespaceURI, side);
        if (!sideElem) continue;
        const attrs = new Map<string, string>();
        for (const attr of BORDER_COMPARE_ATTRIBUTES) {
            const value = sideElem.getAttributeNS(namespaceURI, attr);
            if (value !== null) attrs.set(attr, value);
        }
        out.set(side, attrs);
    }
    return out.size > 0 ? out : null;
}

function borderSignaturesEqual(a: BorderSignature, b: BorderSignature): boolean {
    if (a.size !== b.size) return false;
    for (const side of TABLE_BORDER_SIDES) {
        const aAttrs = a.get(side);
        const bAttrs = b.get(side);
        if (!aAttrs && !bAttrs) continue;
        if (!aAttrs || !bAttrs) return false;
        if (aAttrs.size !== bAttrs.size) return false;
        for (const [attr, value] of aAttrs) {
            if (bAttrs.get(attr) !== value) return false;
        }
    }
    return true;
}

function collectNumericAttributeValues(dom: Document, namespaceURI: string, local: string, attr: string): Set<number> {
    const out = new Set<number>();
    const elems = getElementsByTagNameNSAll(dom, namespaceURI, local);
    for (let i = 0; i < elems.length; i += 1) {
        const elem = elems[i];
        if (!elem) continue;
        const value = Number.parseInt(elem.getAttribute(attr) ?? "", 10);
        if (Number.isFinite(value)) out.add(value);
    }
    return out;
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
    let count = 0;
    if (!doc.documentElement) return 0;

    const stack: Element[] = [doc.documentElement];
    while (stack.length > 0) {
        const node = stack.pop();
        if (!node) continue;

        const localName = node.localName;
        const ns = node.namespaceURI;

        if (localName === "p" && (ns === WORD_2006_NAMESPACE || ns === WORD_STRICT_NAMESPACE)) {
            count++;
        }

        if (
            (localName === "txbxContent" && (ns === WORD_2006_NAMESPACE || ns === WORD_STRICT_NAMESPACE)) ||
            (localName === "textbox" && ns === VML_NAMESPACE)
        ) {
            continue;
        }

        let child = node.lastChild;
        while (child) {
            if (child.nodeType === 1) {
                stack.push(child as Element);
            }
            child = child.previousSibling;
        }
    }

    return count;
}
