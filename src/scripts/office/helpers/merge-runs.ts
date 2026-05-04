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
 * Port target: src/docx-validate/scripts/office/helpers/merge_runs.py — task #4.
 *
 * Merge adjacent runs with identical formatting in DOCX.
 *
 * Merges adjacent <w:r> elements that have identical <w:rPr> properties.
 * Works on runs in paragraphs and inside tracked changes (<w:ins>, <w:del>).
 *
 * Also:
 * - Removes rsid attributes from runs (revision metadata that doesn't affect rendering)
 * - Removes proofErr elements (spell/grammar markers that block merging)
 */

import { promises as fs } from "node:fs";
import path from "node:path";

import { parseXml, serializeXml } from "../../../lib/xml-helpers.ts";

export interface MergeRunsResult {
    count: number;
    message: string;
}

/**
 * Public entry point — mirrors `merge_runs(input_dir)` in
 * `src/docx-validate/scripts/office/helpers/merge_runs.py`.
 */
export async function mergeRuns(inputDir: string): Promise<MergeRunsResult> {
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

        removeElements(root, "proofErr");
        stripRunRsidAttrs(root);

        const containers = new Set<Element>();
        for (const run of findElements(root, "r")) {
            const parent = run.parentNode;
            if (parent && parent.nodeType === ELEMENT_NODE) {
                containers.add(parent as Element);
            }
        }

        let mergeCount = 0;
        for (const container of containers) {
            mergeCount += mergeRunsIn(container);
        }

        await fs.writeFile(docXml, serializeXml(dom, "UTF-8"), "utf8");
        return { count: mergeCount, message: `Merged ${mergeCount} runs` };
    } catch (e) {
        const msg = e instanceof Error ? e.message : String(e);
        return { count: 0, message: `Error: ${msg}` };
    }
}

const ELEMENT_NODE = 1;
const TEXT_NODE = 3;

function localTagName(node: Element): string {
    return node.localName ?? node.tagName;
}

function matchesTag(node: Element, tag: string): boolean {
    const name = localTagName(node);
    return name === tag || name.endsWith(`:${tag}`);
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

function getChild(parent: Element, tag: string): Element | null {
    for (let child = parent.firstChild; child; child = child.nextSibling) {
        if (child.nodeType === ELEMENT_NODE && matchesTag(child as Element, tag)) {
            return child as Element;
        }
    }
    return null;
}

function getChildren(parent: Element, tag: string): Element[] {
    const out: Element[] = [];
    for (let child = parent.firstChild; child; child = child.nextSibling) {
        if (child.nodeType === ELEMENT_NODE && matchesTag(child as Element, tag)) {
            out.push(child as Element);
        }
    }
    return out;
}

function isAdjacent(elem1: Element, elem2: Element): boolean {
    let node = elem1.nextSibling;
    while (node) {
        if (node === elem2) {
            return true;
        }
        if (node.nodeType === ELEMENT_NODE) {
            return false;
        }
        if (node.nodeType === TEXT_NODE && (node.nodeValue ?? "").trim() !== "") {
            return false;
        }
        node = node.nextSibling;
    }
    return false;
}

function removeElements(root: Node, tag: string): void {
    for (const elem of findElements(root, tag)) {
        elem.parentNode?.removeChild(elem);
    }
}

function stripRunRsidAttrs(root: Node): void {
    for (const run of findElements(root, "r")) {
        const attrs = run.attributes;
        if (!attrs) {
            continue;
        }
        const toRemove: string[] = [];
        for (let i = 0; i < attrs.length; i += 1) {
            const attr = attrs.item(i);
            if (attr && attr.name.toLowerCase().includes("rsid")) {
                toRemove.push(attr.name);
            }
        }
        for (const name of toRemove) {
            run.removeAttribute(name);
        }
    }
}

function isRun(node: Node): boolean {
    if (node.nodeType !== ELEMENT_NODE) {
        return false;
    }
    return matchesTag(node as Element, "r");
}

function firstChildRun(container: Element): Element | null {
    for (let child = container.firstChild; child; child = child.nextSibling) {
        if (isRun(child)) {
            return child as Element;
        }
    }
    return null;
}

function nextElementSibling(node: Node): Element | null {
    let sibling = node.nextSibling;
    while (sibling) {
        if (sibling.nodeType === ELEMENT_NODE) {
            return sibling as Element;
        }
        sibling = sibling.nextSibling;
    }
    return null;
}

function nextSiblingRun(node: Node): Element | null {
    let sibling = node.nextSibling;
    while (sibling) {
        if (sibling.nodeType === ELEMENT_NODE) {
            if (isRun(sibling)) {
                return sibling as Element;
            }
        }
        sibling = sibling.nextSibling;
    }
    return null;
}

function canMerge(run1: Element, run2: Element): boolean {
    const rpr1 = getChild(run1, "rPr");
    const rpr2 = getChild(run2, "rPr");

    if ((rpr1 === null) !== (rpr2 === null)) {
        return false;
    }
    if (rpr1 === null) {
        return true;
    }
    return serializeXml(rpr1) === serializeXml(rpr2 as Element);
}

function mergeRunContent(target: Element, source: Element): void {
    const children: Node[] = [];
    for (let child = source.firstChild; child; child = child.nextSibling) {
        children.push(child);
    }
    for (const child of children) {
        if (child.nodeType === ELEMENT_NODE) {
            const el = child as Element;
            if (!matchesTag(el, "rPr")) {
                target.appendChild(child);
            }
        }
    }
}

function consolidateText(run: Element): void {
    const tElements = getChildren(run, "t");

    for (let i = tElements.length - 1; i > 0; i -= 1) {
        const curr = tElements[i];
        const prev = tElements[i - 1];

        if (isAdjacent(prev, curr)) {
            const prevFirst = prev.firstChild as Text | null;
            const currFirst = curr.firstChild as Text | null;
            const prevText = prevFirst ? (prevFirst.data ?? "") : "";
            const currText = currFirst ? (currFirst.data ?? "") : "";
            const merged = prevText + currText;

            if (prevFirst) {
                prevFirst.data = merged;
            } else {
                const ownerDoc = run.ownerDocument;
                if (ownerDoc) {
                    prev.appendChild(ownerDoc.createTextNode(merged));
                }
            }

            if (merged.startsWith(" ") || merged.endsWith(" ")) {
                prev.setAttribute("xml:space", "preserve");
            } else if (prev.hasAttribute("xml:space")) {
                prev.removeAttribute("xml:space");
            }

            run.removeChild(curr);
        }
    }
}

function mergeRunsIn(container: Element): number {
    let mergeCount = 0;
    let run = firstChildRun(container);

    while (run) {
        while (true) {
            const nextElem = nextElementSibling(run);
            if (nextElem && isRun(nextElem) && canMerge(run, nextElem)) {
                mergeRunContent(run, nextElem);
                container.removeChild(nextElem);
                mergeCount += 1;
            } else {
                break;
            }
        }

        consolidateText(run);
        run = nextSiblingRun(run);
    }

    return mergeCount;
}
