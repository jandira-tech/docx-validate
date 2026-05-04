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
 * Port target: src/docx-validate/scripts/office/helpers/simplify_redlines.py — task #5.
 *
 * Simplify tracked changes by merging adjacent w:ins or w:del elements.
 *
 * Merges adjacent <w:ins> elements from the same author into a single element.
 * Same for <w:del> elements. This makes heavily-redlined documents easier to
 * work with by reducing the number of tracked change wrappers.
 *
 * Rules:
 * - Only merges w:ins with w:ins, w:del with w:del (same element type)
 * - Only merges if same author (ignores timestamp differences)
 * - Only merges if truly adjacent (only whitespace between them)
 */

import { promises as fs } from "node:fs";
import path from "node:path";

import JSZip from "jszip";

import { NS } from "../../../lib/types";
import { getElementsByTagNameNSAll, parseXml, serializeXml } from "../../../lib/xml-helpers";

const ELEMENT_NODE = 1;
const TEXT_NODE = 3;

export interface SimplifyRedlinesResult {
    count: number;
    message: string;
}

/**
 * Public entry point — mirrors `simplify_redlines(input_dir)` in
 * `src/docx-validate/scripts/office/helpers/simplify_redlines.py`.
 */
export async function simplifyRedlines(inputDir: string): Promise<SimplifyRedlinesResult> {
    const docXml = path.join(inputDir, "word", "document.xml");

    let text: string;
    try {
        text = await fs.readFile(docXml, "utf8");
    } catch {
        return { count: 0, message: `Error: ${docXml} not found` };
    }

    try {
        const dom = parseXml(text);
        const root = dom.documentElement;
        if (!root) {
            return { count: 0, message: "Error: empty document" };
        }

        let mergeCount = 0;

        const containers: Element[] = [...findElements(root, "p"), ...findElements(root, "tc")];

        for (const container of containers) {
            mergeCount += mergeTrackedChangesIn(container, "ins");
            mergeCount += mergeTrackedChangesIn(container, "del");
        }

        await fs.writeFile(docXml, serializeXml(dom, "UTF-8"), "utf8");
        return {
            count: mergeCount,
            message: `Simplified ${mergeCount} tracked changes`,
        };
    } catch (e) {
        const msg = e instanceof Error ? e.message : String(e);
        return { count: 0, message: `Error: ${msg}` };
    }
}

/**
 * Infer the author of new tracked changes between `modifiedDir` and
 * `originalDocx`. Returns the single new author when exactly one is found,
 * `defaultAuthor` when no new authors are detected, and throws when more
 * than one author has added changes (ambiguous; caller must pick).
 *
 * `defaultAuthor` is required — there is no built-in fallback string.
 */
export async function inferAuthor(modifiedDir: string, originalDocx: string, defaultAuthor: string): Promise<string> {
    const modifiedXml = path.join(modifiedDir, "word", "document.xml");
    const modifiedAuthors = await getTrackedChangeAuthorsFromPath(modifiedXml);

    if (modifiedAuthors.size === 0) {
        return defaultAuthor;
    }

    const originalAuthors = await getAuthorsFromDocx(originalDocx);

    const newChanges = new Map<string, number>();
    for (const [author, count] of modifiedAuthors) {
        const originalCount = originalAuthors.get(author) ?? 0;
        const diff = count - originalCount;
        if (diff > 0) {
            newChanges.set(author, diff);
        }
    }

    if (newChanges.size === 0) {
        return defaultAuthor;
    }

    if (newChanges.size === 1) {
        return newChanges.keys().next().value as string;
    }

    const summary = Array.from(newChanges.entries())
        .map(([k, v]) => `${k}: ${v}`)
        .join(", ");
    throw new Error(`Multiple authors added new changes: {${summary}}. Cannot infer which author to validate.`);
}

/**
 * Count tracked-change authors in a `word/document.xml` file. Mirrors the
 * Python `get_tracked_change_authors(doc_xml_path)` helper. Returns an empty
 * map for missing or unparseable files.
 */
export async function getTrackedChangeAuthorsFromPath(docXmlPath: string): Promise<Map<string, number>> {
    let text: string;
    try {
        text = await fs.readFile(docXmlPath, "utf8");
    } catch {
        return new Map();
    }
    return countAuthors(text);
}

function countAuthors(xmlText: string): Map<string, number> {
    const out = new Map<string, number>();
    let dom: Document;
    try {
        dom = parseXml(xmlText);
    } catch {
        return out;
    }
    const root = dom.documentElement;
    if (!root) {
        return out;
    }

    for (const tag of ["ins", "del"] as const) {
        for (const elem of getElementsByTagNameNSAll(root, NS.W, tag)) {
            const author = elem.getAttributeNS(NS.W, "author") || elem.getAttribute("w:author");
            if (author) {
                out.set(author, (out.get(author) ?? 0) + 1);
            }
        }
    }

    return out;
}

async function getAuthorsFromDocx(docxPath: string): Promise<Map<string, number>> {
    let buf: Buffer;
    try {
        buf = await fs.readFile(docxPath);
    } catch {
        return new Map();
    }

    let zip: JSZip;
    try {
        zip = await JSZip.loadAsync(buf);
    } catch {
        return new Map();
    }

    const file = zip.file("word/document.xml");
    if (!file) {
        return new Map();
    }

    const text = await file.async("string");
    return countAuthors(text);
}

function findElements(root: Node, tag: string): Element[] {
    const results: Element[] = [];
    const traverse = (node: Node): void => {
        if (node.nodeType === ELEMENT_NODE) {
            const el = node as Element;
            if (matchesTag(el, tag)) {
                results.push(el);
            }
            for (let child = node.firstChild; child; child = child.nextSibling) {
                traverse(child);
            }
        }
    };
    traverse(root);
    return results;
}

function matchesTag(node: Element, tag: string): boolean {
    const name = node.localName ?? node.tagName;
    return name === tag || name.endsWith(`:${tag}`);
}

function isElement(node: Node, tag: string): boolean {
    return node.nodeType === ELEMENT_NODE && matchesTag(node as Element, tag);
}

function getAuthor(elem: Element): string {
    const direct = elem.getAttribute("w:author");
    if (direct) {
        return direct;
    }
    const attrs = elem.attributes;
    if (attrs) {
        for (let i = 0; i < attrs.length; i += 1) {
            const attr = attrs.item(i);
            if (!attr) continue;
            if (attr.localName === "author" || attr.name.endsWith(":author")) {
                return attr.value;
            }
        }
    }
    return "";
}

function canMergeTracked(elem1: Element, elem2: Element): boolean {
    if (getAuthor(elem1) !== getAuthor(elem2)) {
        return false;
    }

    let node = elem1.nextSibling;
    while (node && node !== elem2) {
        if (node.nodeType === ELEMENT_NODE) {
            return false;
        }
        if (node.nodeType === TEXT_NODE && (node.nodeValue ?? "").trim() !== "") {
            return false;
        }
        node = node.nextSibling;
    }

    return true;
}

function mergeTrackedContent(target: Element, source: Element): void {
    while (source.firstChild) {
        const child = source.firstChild;
        source.removeChild(child);
        target.appendChild(child);
    }
}

function mergeTrackedChangesIn(container: Element, tag: string): number {
    const tracked: Element[] = [];
    for (let child = container.firstChild; child; child = child.nextSibling) {
        if (isElement(child, tag)) {
            tracked.push(child as Element);
        }
    }

    if (tracked.length < 2) {
        return 0;
    }

    let mergeCount = 0;
    let i = 0;
    while (i < tracked.length - 1) {
        const curr = tracked[i];
        const nextElem = tracked[i + 1];

        if (canMergeTracked(curr, nextElem)) {
            mergeTrackedContent(curr, nextElem);
            container.removeChild(nextElem);
            tracked.splice(i + 1, 1);
            mergeCount += 1;
        } else {
            i += 1;
        }
    }

    return mergeCount;
}
