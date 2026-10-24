/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

/**
 * Content-model placement of runs and run properties, as Word sees it.
 *
 * Replaces the `r` / `rPr` entries of `WORD_BLOCKING_MISPLACED_LOCALS` and
 * the `}rPr': This element is not expected.` message match in
 * `isWordBlockingXsdIssue` (validators/docx.ts) with namespace- and
 * parent-aware DOM checks that need no libxml2 diagnostic.
 *
 * Word evidence:
 * - Misplaced run: `tests/fixtures/broken/word-rejects-misplaced-run.docx`
 *   puts `w:r` directly inside `w:tc`; real Word answers OPEN_ERROR. The
 *   probed twins `word-invalid-fixtures/OPEN_ERROR/tests__fixtures__critic-
 *   variations__docx__13-table-with-changes.docx` and `20-kitchen-sink.docx`
 *   carry the same shape. The tolerated control is the same `w:r` under
 *   `w:p`, `w:ins`, `w:del`, and `w:hyperlink` in the very same documents.
 * - Misplaced run properties: `tests/fixtures/external/open-xml-sdk/5Errors.docx`
 *   (manifest `word: unreadable-content-warning`) carries a second `w:rPr`
 *   inside one `w:r` plus an `w:pPr` nested in `w:pPr`; Word offers recovery.
 *   The control is the single leading `w:rPr` of every well-formed run.
 * - Misplaced underline: `tests/fixtures/external/superdoc/super-editor/
 *   invalid-underline.docx` (manifest `word: unreadable-content-warning`)
 *   puts `w:u` directly inside `w:pPr`; the control is `w:u` inside `w:rPr`.
 */

import type { CheckDefinition } from "../../../lib/check-registry";
import type { ValidationIssue } from "../../../lib/types";
import { getElementsByTagNameNSAll } from "../../../lib/xml-helpers";
import { type WmlPartDocument, WORDPROCESSINGML_NAMESPACE as W, wmlPartDocuments } from "./wml-parts";

const MATH_NAMESPACE = "http://schemas.openxmlformats.org/officeDocument/2006/math";

/**
 * Local names of the parents that may hold a `w:r` directly, exactly as
 * wml.xsd admits them (EG_PContent members and the run-content wrappers).
 * Pinned against the bundled schema by
 * `tests/word-content-model.integration.test.ts`.
 */
export const RUN_PARENTS: ReadonlySet<string> = new Set([
    "bdo",
    "customXml",
    "del",
    "dir",
    "fldSimple",
    "hyperlink",
    "ins",
    "moveFrom",
    "moveTo",
    "p",
    "rt",
    "rubyBase",
    "sdtContent",
    "smartTag",
]);

/**
 * Parents that may hold a `w:rPr` directly, per namespace. The WordprocessingML
 * names come from wml.xsd (CT_R, CT_PPr, CT_Style, CT_Lvl, CT_TblStylePr,
 * CT_SdtPr, CT_SdtEndPr, CT_RPrDefault, and the math-control ins/del of
 * EG_RPrMath); the math names come from shared-math.xsd: CT_R and CT_CtrlPr
 * both reference `w:EG_RPr`, so `m:r` and `m:ctrlPr` legally hold a `w:rPr`
 * (math runs additionally carry their own `m:rPr`, which is not this check's
 * business). Pinned by `tests/word-content-model.integration.test.ts`.
 */
export const RUN_PROPERTIES_PARENTS: Readonly<Record<string, ReadonlySet<string>>> = Object.freeze({
    [W]: new Set(["del", "ins", "lvl", "pPr", "r", "rPrChange", "rPrDefault", "sdtEndPr", "sdtPr", "style", "tblStylePr"]),
    [MATH_NAMESPACE]: new Set(["ctrlPr", "r"]),
});

/** `w:u` is a member of EG_RPrBase: only a `w:rPr` may hold it directly. */
export const UNDERLINE_PARENTS: ReadonlySet<string> = new Set(["rPr"]);

const elementParent = (node: Element): Element | null => {
    const parent = node.parentNode;
    return parent?.nodeType === 1 ? (parent as Element) : null;
};

/** Whether `elem` has a preceding element sibling inside its parent. */
const hasPrecedingElementSibling = (elem: Element): boolean => {
    for (let sibling = elem.previousSibling; sibling; sibling = sibling.previousSibling) {
        if (sibling.nodeType === 1) {
            return true;
        }
    }
    return false;
};

const describeParents = (parents: ReadonlySet<string>): string =>
    [...parents]
        .sort()
        .map((local) => `w:${local}`)
        .join(", ");

const runIssues = ({ part, document: doc }: WmlPartDocument): ValidationIssue[] => {
    const issues: ValidationIssue[] = [];
    for (const run of getElementsByTagNameNSAll(doc, W, "r")) {
        const parent = elementParent(run);
        if (!parent || parent.namespaceURI !== W || RUN_PARENTS.has(parent.localName ?? "")) {
            continue;
        }
        issues.push({
            severity: "error",
            code: wordRunMisplacedCheck.id,
            path: part,
            message:
                `w:r sits directly inside w:${parent.localName}; wml.xsd admits a run only inside paragraph content ` +
                `(${describeParents(RUN_PARENTS)}). Microsoft Word refuses to open a run outside a paragraph.`,
        });
    }
    return issues;
};

const runPropertiesIssues = ({ part, document: doc }: WmlPartDocument): ValidationIssue[] => {
    const issues: ValidationIssue[] = [];
    for (const properties of getElementsByTagNameNSAll(doc, W, "rPr")) {
        const parent = elementParent(properties);
        if (!parent) {
            continue;
        }
        // Own keys only: a namespace URI such as "constructor" must not reach Object.prototype.
        const namespace = parent.namespaceURI ?? "";
        const allowed = Object.hasOwn(RUN_PROPERTIES_PARENTS, namespace) ? RUN_PROPERTIES_PARENTS[namespace] : undefined;
        if (allowed?.has(parent.localName ?? "")) {
            // CT_R admits w:rPr only as its first child, before run content; a
            // second or trailing w:rPr is the defect 5Errors.docx carries.
            if (parent.namespaceURI === W && parent.localName === "r" && hasPrecedingElementSibling(properties)) {
                issues.push({
                    severity: "error",
                    code: wordRunPropertiesMisplacedCheck.id,
                    path: part,
                    message:
                        "w:rPr follows another element inside w:r; CT_R requires the run properties first, before any run content. " +
                        "Microsoft Word offers recovery for a duplicated or trailing run-properties block.",
                });
            }
            continue;
        }
        issues.push({
            severity: "error",
            code: wordRunPropertiesMisplacedCheck.id,
            path: part,
            message:
                `w:rPr sits directly inside ${parent.namespaceURI === MATH_NAMESPACE ? "m" : "w"}:${parent.localName}; run properties belong inside ` +
                `${describeParents(RUN_PROPERTIES_PARENTS[W] ?? new Set())} (or m:ctrlPr). Microsoft Word offers recovery for run properties outside their containers.`,
        });
    }
    return issues;
};

const underlineIssues = ({ part, document: doc }: WmlPartDocument): ValidationIssue[] => {
    const issues: ValidationIssue[] = [];
    for (const underline of getElementsByTagNameNSAll(doc, W, "u")) {
        const parent = elementParent(underline);
        if (!parent || parent.namespaceURI !== W || UNDERLINE_PARENTS.has(parent.localName ?? "")) {
            continue;
        }
        issues.push({
            severity: "error",
            code: wordUnderlineMisplacedCheck.id,
            path: part,
            message:
                `w:u sits directly inside w:${parent.localName}; underline is a run property and belongs inside w:rPr. ` +
                "Microsoft Word offers recovery for a run property at paragraph-properties level.",
        });
    }
    return issues;
};

export const wordRunMisplacedCheck: CheckDefinition = {
    id: "word-run-misplaced",
    description: "Runs sit inside paragraph content, not directly in cells, rows, or the body.",
    groups: ["story"],
    formats: ["docx"],
    requires: ["xml-syntax"],
    run: (ctx): ValidationIssue[] => wmlPartDocuments(ctx).flatMap(runIssues),
};

export const wordRunPropertiesMisplacedCheck: CheckDefinition = {
    id: "word-run-properties-misplaced",
    description: "Run-properties blocks sit in their schema parents and lead their runs.",
    groups: ["story"],
    formats: ["docx"],
    requires: ["xml-syntax"],
    run: (ctx): ValidationIssue[] => wmlPartDocuments(ctx).flatMap(runPropertiesIssues),
};

export const wordUnderlineMisplacedCheck: CheckDefinition = {
    id: "word-underline-misplaced",
    description: "Underline run properties sit inside a run-properties block.",
    groups: ["story"],
    formats: ["docx"],
    requires: ["xml-syntax"],
    run: (ctx): ValidationIssue[] => wmlPartDocuments(ctx).flatMap(underlineIssues),
};
