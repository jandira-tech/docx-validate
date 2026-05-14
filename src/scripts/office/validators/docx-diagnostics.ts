import { promises as fs } from "node:fs";
import path from "node:path";

import type { ValidationIssue } from "../../../lib/types";
import { parseXml } from "../../../lib/xml-helpers";

const WORD_NAMESPACES = new Set([
    "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "http://purl.oclc.org/ooxml/wordprocessingml/main",
]);
const W15_NAMESPACE = "http://schemas.microsoft.com/office/word/2012/wordml";
const W16CEX_NAMESPACE = "http://schemas.microsoft.com/office/word/2018/wordml/cex";
const W16CID_NAMESPACE = "http://schemas.microsoft.com/office/word/2016/wordml/cid";
const PACKAGE_RELATIONSHIPS_NAMESPACE = "http://schemas.openxmlformats.org/package/2006/relationships";
const CONTENT_TYPES_NAMESPACE = "http://schemas.openxmlformats.org/package/2006/content-types";
const MATH_NAMESPACE = "http://schemas.openxmlformats.org/officeDocument/2006/math";
const WP_NAMESPACE = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
const PICTURE_NAMESPACE = "http://schemas.openxmlformats.org/drawingml/2006/picture";

export interface DocxSemanticCounter {
    category: string;
    label: string;
    path: string;
    unit: string;
    count: number;
}

export interface DocxSemanticInventory {
    counters: ReadonlyMap<string, DocxSemanticCounter>;
}

interface MutableDocxSemanticInventory {
    counters: Map<string, DocxSemanticCounter>;
}

interface RepairPlanGroup {
    path?: string;
    code: string;
    action: string;
    supported: boolean;
    count: number;
}

export async function collectDocxSemanticInventory(unpackedDir: string): Promise<DocxSemanticInventory> {
    const inventory: MutableDocxSemanticInventory = { counters: new Map() };
    const files = await walkFiles(unpackedDir);
    for (const file of files) {
        const rel = path.relative(unpackedDir, file).split(path.sep).join("/");
        if (!rel.endsWith(".xml") && !rel.endsWith(".rels")) {
            await collectPackageAsset(file, rel, inventory);
            continue;
        }
        let dom: Document;
        try {
            dom = parseXml(await fs.readFile(file, "utf-8"));
        } catch {
            continue;
        }
        collectXmlPart(rel, dom, inventory);
    }
    return { counters: new Map([...inventory.counters].sort(([a], [b]) => a.localeCompare(b))) };
}

export function compareDocxSemanticInventories(before: DocxSemanticInventory, after: DocxSemanticInventory): ValidationIssue[] {
    const issues: ValidationIssue[] = [];
    const beforeCounters = [...before.counters.values()].sort(compareCounters);
    for (const counter of beforeCounters) {
        const afterCount = after.counters.get(counterKey(counter.path, counter.category, counter.label, counter.unit))?.count ?? 0;
        if (counter.count <= afterCount) continue;
        const lost = counter.count - afterCount;
        issues.push({
            severity: "error",
            path: counter.path,
            code: "repair-content-loss",
            message:
                `Repair lost ${counter.category} '${counter.label}': ` + `${counter.count} → ${afterCount} (-${lost} ${counter.unit}).`,
        });
    }
    if (issues.length > 0) return issues;
    return [
        {
            severity: "info",
            code: "repair-content-preserved",
            message: `Repair semantic inventory preserved: no tracked content counters decreased across ${before.counters.size} counter(s).`,
        },
    ];
}

export function buildRepairPlanIssues(issues: readonly ValidationIssue[]): ValidationIssue[] {
    const groups = new Map<string, RepairPlanGroup>();
    for (const issue of issues) {
        if (issue.severity === "info") continue;
        const code = issue.code ?? "uncoded-validation-issue";
        const action = repairActionForIssue(issue);
        const supported = action !== null;
        const groupAction = action ?? "no automatic repair is registered; validation will report it again after repair if it remains";
        const key = `${issue.path ?? ""}\u0000${code}\u0000${supported ? "1" : "0"}\u0000${groupAction}`;
        const existing = groups.get(key);
        if (existing) {
            existing.count += 1;
            continue;
        }
        groups.set(key, {
            path: issue.path,
            code,
            action: groupAction,
            supported,
            count: 1,
        });
    }
    return [...groups.values()].sort(compareRepairPlanGroups).map((group) => ({
        severity: "info" as const,
        path: group.path,
        code: group.supported ? "repair-plan" : "repair-plan-unavailable",
        message:
            `Before repair found ${group.count} [${group.code}] issue(s). ` +
            (group.supported ? `Auto-repair will ${group.action}.` : `Auto-repair will not change them because ${group.action}.`),
    }));
}

async function collectPackageAsset(file: string, rel: string, inventory: MutableDocxSemanticInventory): Promise<void> {
    const stat = await fs.stat(file);
    addCounter(inventory, rel, "package asset", "part exists", "part(s)", 1);
    addCounter(inventory, rel, "package asset", "part bytes", "byte(s)", stat.size);
}

function collectXmlPart(rel: string, dom: Document, inventory: MutableDocxSemanticInventory): void {
    if (rel.endsWith(".rels")) {
        collectRelationships(rel, dom, inventory);
        return;
    }
    if (rel === "[Content_Types].xml") collectContentTypes(rel, dom, inventory);
    collectDocumentStructure(rel, dom, inventory);
    collectText(rel, dom, inventory);
    collectFormatting(rel, dom, inventory);
    collectStyles(rel, dom, inventory);
    collectComments(rel, dom, inventory);
    collectTrackedChanges(rel, dom, inventory);
}

function collectRelationships(rel: string, dom: Document, inventory: MutableDocxSemanticInventory): void {
    const rels = dom.getElementsByTagNameNS(PACKAGE_RELATIONSHIPS_NAMESPACE, "Relationship");
    for (let i = 0; i < rels.length; i += 1) {
        const elem = rels.item(i);
        if (!elem) continue;
        const type = elem.getAttribute("Type") ?? "unknown";
        const targetMode = elem.getAttribute("TargetMode") ?? "Internal";
        addCounter(inventory, rel, "relationship", `type '${shortRelationshipType(type)}'`, "relationship(s)", 1);
        addCounter(inventory, rel, "relationship", `target mode '${targetMode}'`, "relationship(s)", 1);
    }
}

function collectContentTypes(rel: string, dom: Document, inventory: MutableDocxSemanticInventory): void {
    const overrides = dom.getElementsByTagNameNS(CONTENT_TYPES_NAMESPACE, "Override");
    for (let i = 0; i < overrides.length; i += 1) {
        const elem = overrides.item(i);
        if (!elem) continue;
        addCounter(inventory, rel, "content type", `override '${elem.getAttribute("ContentType") ?? "unknown"}'`, "declaration(s)", 1);
    }
    const defaults = dom.getElementsByTagNameNS(CONTENT_TYPES_NAMESPACE, "Default");
    for (let i = 0; i < defaults.length; i += 1) {
        const elem = defaults.item(i);
        if (!elem) continue;
        addCounter(inventory, rel, "content type", `default '${elem.getAttribute("ContentType") ?? "unknown"}'`, "declaration(s)", 1);
    }
}

function collectDocumentStructure(rel: string, dom: Document, inventory: MutableDocxSemanticInventory): void {
    const structure = new Map<string, string>([
        ["p", "paragraph"],
        ["r", "run"],
        ["tbl", "table"],
        ["tr", "table row"],
        ["tc", "table cell"],
        ["sdt", "content control"],
        ["hyperlink", "hyperlink"],
        ["fldSimple", "simple field"],
        ["footnote", "footnote"],
        ["endnote", "endnote"],
        ["bookmarkStart", "bookmark start"],
        ["bookmarkEnd", "bookmark end"],
        ["drawing", "drawing reference"],
        ["pict", "VML picture"],
    ]);
    for (const [local, label] of structure) {
        let count = 0;
        for (const ns of WORD_NAMESPACES) count += dom.getElementsByTagNameNS(ns, local).length;
        if (count > 0) addCounter(inventory, rel, "document structure", label, "element(s)", count);
    }
    const mathElements =
        dom.getElementsByTagNameNS(MATH_NAMESPACE, "oMath").length + dom.getElementsByTagNameNS(MATH_NAMESPACE, "oMathPara").length;
    if (mathElements > 0) addCounter(inventory, rel, "document structure", "math object", "element(s)", mathElements);
    const drawingElements =
        dom.getElementsByTagNameNS(WP_NAMESPACE, "inline").length + dom.getElementsByTagNameNS(WP_NAMESPACE, "anchor").length;
    if (drawingElements > 0) addCounter(inventory, rel, "document structure", "DrawingML placement", "element(s)", drawingElements);
    const pictures = dom.getElementsByTagNameNS(PICTURE_NAMESPACE, "pic").length;
    if (pictures > 0) addCounter(inventory, rel, "document structure", "DrawingML picture", "element(s)", pictures);
}

function collectText(rel: string, dom: Document, inventory: MutableDocxSemanticInventory): void {
    const all = dom.getElementsByTagName("*");
    for (let i = 0; i < all.length; i += 1) {
        const elem = all.item(i);
        if (!elem || !isWordElement(elem)) continue;
        const local = localName(elem);
        if (local === "fldSimple") {
            const instr = elem.getAttributeNS(wordNamespace(elem), "instr") ?? elem.getAttribute("w:instr") ?? elem.getAttribute("instr");
            if (instr) addCounter(inventory, rel, "text", "field instruction", "character(s)", instr.length);
            continue;
        }
        if (local !== "t" && local !== "delText" && local !== "instrText" && local !== "delInstrText") continue;
        const length = elem.textContent?.length ?? 0;
        if (length === 0) continue;
        addCounter(inventory, rel, "text", textLabel(elem, local), "character(s)", length);
    }
}

function collectFormatting(rel: string, dom: Document, inventory: MutableDocxSemanticInventory): void {
    for (const ns of WORD_NAMESPACES) {
        const runs = dom.getElementsByTagNameNS(ns, "r");
        for (let i = 0; i < runs.length; i += 1) {
            const run = runs.item(i);
            if (!run) continue;
            const chars = runTextLength(run);
            if (chars === 0) continue;
            const rPr = directWordChild(run, "rPr");
            if (!rPr) continue;
            if (enabledChild(rPr, "b") || enabledChild(rPr, "bCs"))
                addCounter(inventory, rel, "formatting", "bold", "formatted character(s)", chars);
            if (enabledChild(rPr, "i") || enabledChild(rPr, "iCs"))
                addCounter(inventory, rel, "formatting", "italic", "formatted character(s)", chars);
            if (enabledChild(rPr, "strike")) addCounter(inventory, rel, "formatting", "strikethrough", "formatted character(s)", chars);
            if (enabledChild(rPr, "dstrike"))
                addCounter(inventory, rel, "formatting", "double strikethrough", "formatted character(s)", chars);
            if (enabledChild(rPr, "caps")) addCounter(inventory, rel, "formatting", "caps", "formatted character(s)", chars);
            if (enabledChild(rPr, "smallCaps")) addCounter(inventory, rel, "formatting", "small caps", "formatted character(s)", chars);
            if (enabledChild(rPr, "vanish")) addCounter(inventory, rel, "formatting", "hidden", "formatted character(s)", chars);
            const underline = wordChildAttr(rPr, "u", "val");
            if (underline && underline !== "none")
                addCounter(inventory, rel, "formatting", `underline '${underline}'`, "formatted character(s)", chars);
            const color = wordChildAttr(rPr, "color", "val");
            if (color) addCounter(inventory, rel, "formatting", `color '${color}'`, "formatted character(s)", chars);
            const highlight = wordChildAttr(rPr, "highlight", "val");
            if (highlight) addCounter(inventory, rel, "formatting", `highlight '${highlight}'`, "formatted character(s)", chars);
            const size = wordChildAttr(rPr, "sz", "val");
            if (size) addCounter(inventory, rel, "formatting", `font size '${size}'`, "formatted character(s)", chars);
            const vertAlign = wordChildAttr(rPr, "vertAlign", "val");
            if (vertAlign) addCounter(inventory, rel, "formatting", `vertical align '${vertAlign}'`, "formatted character(s)", chars);
            const rStyle = wordChildAttr(rPr, "rStyle", "val");
            if (rStyle) addCounter(inventory, rel, "style reference", `run style '${rStyle}'`, "styled character(s)", chars);
        }

        const paragraphs = dom.getElementsByTagNameNS(ns, "p");
        for (let i = 0; i < paragraphs.length; i += 1) {
            const paragraph = paragraphs.item(i);
            if (!paragraph) continue;
            const pPr = directWordChild(paragraph, "pPr");
            if (!pPr) continue;
            const pStyle = wordChildAttr(pPr, "pStyle", "val");
            if (pStyle) addCounter(inventory, rel, "style reference", `paragraph style '${pStyle}'`, "paragraph(s)", 1);
            const numPr = directWordChild(pPr, "numPr");
            if (numPr) {
                const numId = wordChildAttr(numPr, "numId", "val");
                if (numId) addCounter(inventory, rel, "numbering", `numbering id '${numId}'`, "paragraph(s)", 1);
                const ilvl = wordChildAttr(numPr, "ilvl", "val");
                if (ilvl) addCounter(inventory, rel, "numbering", `numbering level '${ilvl}'`, "paragraph(s)", 1);
            }
        }

        const tables = dom.getElementsByTagNameNS(ns, "tbl");
        for (let i = 0; i < tables.length; i += 1) {
            const table = tables.item(i);
            if (!table) continue;
            const tblPr = directWordChild(table, "tblPr");
            if (!tblPr) continue;
            const tblStyle = wordChildAttr(tblPr, "tblStyle", "val");
            if (tblStyle) addCounter(inventory, rel, "style reference", `table style '${tblStyle}'`, "table(s)", 1);
        }
    }
}

function collectStyles(rel: string, dom: Document, inventory: MutableDocxSemanticInventory): void {
    if (!rel.endsWith("/styles.xml") && rel !== "word/styles.xml") return;
    for (const ns of WORD_NAMESPACES) {
        const styles = dom.getElementsByTagNameNS(ns, "style");
        for (let i = 0; i < styles.length; i += 1) {
            const style = styles.item(i);
            if (!style) continue;
            const styleId =
                style.getAttributeNS(ns, "styleId") ?? style.getAttribute("w:styleId") ?? style.getAttribute("styleId") ?? "unknown";
            const type = style.getAttributeNS(ns, "type") ?? style.getAttribute("w:type") ?? style.getAttribute("type") ?? "unknown";
            addCounter(inventory, rel, "style definition", `${type} '${styleId}'`, "definition(s)", 1);
            const rPr = directWordChild(style, "rPr");
            if (!rPr) continue;
            if (enabledChild(rPr, "b") || enabledChild(rPr, "bCs"))
                addCounter(inventory, rel, "style formatting", `'${styleId}' bold`, "property occurrence(s)", 1);
            if (enabledChild(rPr, "i") || enabledChild(rPr, "iCs"))
                addCounter(inventory, rel, "style formatting", `'${styleId}' italic`, "property occurrence(s)", 1);
            const color = wordChildAttr(rPr, "color", "val");
            if (color) addCounter(inventory, rel, "style formatting", `'${styleId}' color '${color}'`, "property occurrence(s)", 1);
        }
    }
}

function collectComments(rel: string, dom: Document, inventory: MutableDocxSemanticInventory): void {
    for (const ns of WORD_NAMESPACES) {
        const comments = dom.getElementsByTagNameNS(ns, "comment");
        for (let i = 0; i < comments.length; i += 1) {
            const comment = comments.item(i);
            if (!comment) continue;
            addCounter(inventory, rel, "comment", "comment entry", "comment(s)", 1);
            const author = comment.getAttributeNS(ns, "author") ?? comment.getAttribute("w:author") ?? comment.getAttribute("author");
            if (author) addCounter(inventory, rel, "comment", `comment by '${author}'`, "comment(s)", 1);
        }
        for (const local of ["commentRangeStart", "commentRangeEnd", "commentReference"] as const) {
            const count = dom.getElementsByTagNameNS(ns, local).length;
            if (count > 0) addCounter(inventory, rel, "comment marker", local, "marker(s)", count);
        }
    }
    const commentEx = dom.getElementsByTagNameNS(W15_NAMESPACE, "commentEx").length;
    if (commentEx > 0) addCounter(inventory, rel, "comment thread", "commentsExtended entry", "entry(s)", commentEx);
    const commentIds = dom.getElementsByTagNameNS(W16CID_NAMESPACE, "commentId").length;
    if (commentIds > 0) addCounter(inventory, rel, "comment thread", "commentsIds entry", "entry(s)", commentIds);
    const commentExtensible = dom.getElementsByTagNameNS(W16CEX_NAMESPACE, "commentExtensible").length;
    if (commentExtensible > 0) addCounter(inventory, rel, "comment thread", "commentsExtensible entry", "entry(s)", commentExtensible);
}

function collectTrackedChanges(rel: string, dom: Document, inventory: MutableDocxSemanticInventory): void {
    const tracked = [
        "ins",
        "del",
        "moveFrom",
        "moveTo",
        "pPrChange",
        "rPrChange",
        "tblPrChange",
        "trPrChange",
        "tcPrChange",
        "sectPrChange",
    ] as const;
    for (const ns of WORD_NAMESPACES) {
        for (const local of tracked) {
            const elems = dom.getElementsByTagNameNS(ns, local);
            if (elems.length > 0) addCounter(inventory, rel, "tracked change", local, "element(s)", elems.length);
            for (let i = 0; i < elems.length; i += 1) {
                const elem = elems.item(i);
                if (!elem) continue;
                const author = elem.getAttributeNS(ns, "author") ?? elem.getAttribute("w:author") ?? elem.getAttribute("author");
                if (author) addCounter(inventory, rel, "tracked change", `${local} by '${author}'`, "element(s)", 1);
            }
        }
    }
}

function textLabel(elem: Element, local: string): string {
    if (local === "delText") return "deleted text";
    if (local === "delInstrText") return "deleted field instruction";
    if (local === "instrText") return hasWordAncestor(elem, "del") ? "deleted field instruction" : "field instruction";
    if (hasWordAncestor(elem, "ins")) return "inserted text";
    if (hasWordAncestor(elem, "del")) return "deleted text";
    if (hasWordAncestor(elem, "moveFrom")) return "moved-from text";
    if (hasWordAncestor(elem, "moveTo")) return "moved-to text";
    if (hasWordAncestor(elem, "comment")) return "comment text";
    return "visible text";
}

function runTextLength(run: Element): number {
    let length = 0;
    const all = run.getElementsByTagName("*");
    for (let i = 0; i < all.length; i += 1) {
        const elem = all.item(i);
        if (!elem || !isWordElement(elem)) continue;
        const local = localName(elem);
        if (local === "t" || local === "delText" || local === "instrText" || local === "delInstrText") {
            length += elem.textContent?.length ?? 0;
        }
    }
    return length;
}

function enabledChild(parent: Element, local: string): boolean {
    const child = directWordChild(parent, local);
    if (!child) return false;
    const val = child.getAttributeNS(wordNamespace(child), "val") ?? child.getAttribute("w:val") ?? child.getAttribute("val");
    if (val === null || val === "") return true;
    return !new Set(["0", "false", "off", "none"]).has(val.toLowerCase());
}

function wordChildAttr(parent: Element, local: string, attr: string): string | null {
    const child = directWordChild(parent, local);
    if (!child) return null;
    return child.getAttributeNS(wordNamespace(child), attr) ?? child.getAttribute(`w:${attr}`) ?? child.getAttribute(attr);
}

function directWordChild(parent: Element, local: string): Element | null {
    for (let child = parent.firstChild; child; child = child.nextSibling) {
        if (child.nodeType !== 1) continue;
        const elem = child as Element;
        if (isWordElement(elem) && localName(elem) === local) return elem;
    }
    return null;
}

function hasWordAncestor(elem: Element, local: string): boolean {
    for (let node = elem.parentNode; node; node = node.parentNode) {
        if (node.nodeType !== 1) continue;
        const parent = node as Element;
        if (isWordElement(parent) && localName(parent) === local) return true;
    }
    return false;
}

function isWordElement(elem: Element): boolean {
    return WORD_NAMESPACES.has(elem.namespaceURI ?? "");
}

function wordNamespace(elem: Element): string {
    return elem.namespaceURI && WORD_NAMESPACES.has(elem.namespaceURI)
        ? elem.namespaceURI
        : "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
}

function localName(elem: Element): string {
    return elem.localName || elem.tagName.split(":").pop() || elem.tagName;
}

function addCounter(
    inventory: MutableDocxSemanticInventory,
    pathValue: string,
    category: string,
    label: string,
    unit: string,
    count: number,
): void {
    if (count <= 0) return;
    const key = counterKey(pathValue, category, label, unit);
    const existing = inventory.counters.get(key);
    if (existing) {
        existing.count += count;
        return;
    }
    inventory.counters.set(key, { path: pathValue, category, label, unit, count });
}

function counterKey(pathValue: string, category: string, label: string, unit: string): string {
    return `${pathValue}\u0000${category}\u0000${label}\u0000${unit}`;
}

function compareCounters(a: DocxSemanticCounter, b: DocxSemanticCounter): number {
    return (
        a.path.localeCompare(b.path) ||
        a.category.localeCompare(b.category) ||
        a.label.localeCompare(b.label) ||
        a.unit.localeCompare(b.unit)
    );
}

function compareRepairPlanGroups(a: RepairPlanGroup, b: RepairPlanGroup): number {
    return (a.path ?? "").localeCompare(b.path ?? "") || a.code.localeCompare(b.code) || a.action.localeCompare(b.action);
}

function shortRelationshipType(type: string): string {
    return type.split("/").filter(Boolean).pop() ?? type;
}

function repairActionForIssue(issue: ValidationIssue): string | null {
    switch (issue.code) {
        case "ws-missing-preserve":
            return "add xml:space='preserve' to affected Word text elements";
        case "rels-empty-part":
            return "delete empty relationship sidecar parts that contain no relationships";
        case "id-durable-overflow":
        case "id-durable-decimal":
            return "rewrite invalid durableId values to in-range values";
        case "id-paraid-overflow":
            return "rewrite over-cap paraId values and propagate the remap to threaded-comment references";
        case "id-textid-overflow":
            return "rewrite over-cap textId values to in-range values";
        case "ignorable-undeclared":
            return "declare known mc:Ignorable prefixes on the XML root or drop unknown ignorable tokens";
        case "paraid-missing-element":
            return "stamp missing w14:paraId and w14:textId anchors on affected paragraphs or table rows";
        case "textid-missing-element":
            return "stamp missing w14:textId anchors on affected paragraphs or table rows";
        case "style-reference-undefined":
            return "inject canonical definitions for known missing style IDs and leave unknown style IDs reported";
        case "style-default-missing":
            return "inject canonical implied-default Word style definitions";
        case "comment-thread-commentid-paraid-orphan":
        case "comment-thread-commentid-missing-paraid":
        case "comment-thread-commentid-missing-durableid":
        case "comment-thread-durableid-orphan":
        case "comment-thread-durableid-missing":
            return "align commentsIds/commentsExtensible IDs with comments.xml when the target comment thread is unambiguous";
        case "word-drawing-scalar-whitespace":
            return "trim leading and trailing whitespace around scalar DrawingML values";
        default:
            return xsdRepairAction(issue);
    }
}

function xsdRepairAction(issue: ValidationIssue): string | null {
    if (issue.code !== "xsd-error" && issue.code !== "xsd-summary") return null;
    const message = issue.message;
    if (message.includes("}graphic': This element is not expected") || message.includes("}blipFill': This element is not expected")) {
        return "insert missing inline-picture scaffolding required by WordprocessingML/DrawingML schemas";
    }
    if (message.includes("Expected is ( {http://schemas.openxmlformats.org/drawingml/2006/picture}spPr")) {
        return "insert missing picture shape properties required by DrawingML schemas";
    }
    if (message.includes("}filetime': '' is not a valid value")) {
        return "trim invalid whitespace from core document properties that Word expects as scalar values";
    }
    return null;
}

async function walkFiles(root: string): Promise<string[]> {
    const out: string[] = [];
    await walk(root, out);
    return out.sort((a, b) => a.localeCompare(b));
}

async function walk(dir: string, out: string[]): Promise<void> {
    let entries: string[];
    try {
        entries = await fs.readdir(dir);
    } catch {
        return;
    }
    for (const name of entries.sort((a, b) => a.localeCompare(b))) {
        const full = path.join(dir, name);
        const stat = await fs.lstat(full);
        if (stat.isSymbolicLink()) {
            continue;
        }
        if (stat.isDirectory()) {
            await walk(full, out);
        } else if (stat.isFile()) {
            out.push(full);
        }
    }
}
