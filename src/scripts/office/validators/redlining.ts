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
 * Validator for tracked changes in Word documents.
 *
 * Port of `src/docx-validate/scripts/office/validators/redlining.py` (task #9).
 *
 * Contract: strip every `w:ins` / `w:del` authored by `author` from BOTH the
 * modified `word/document.xml` and the original .docx, extract paragraph text
 * from each, and require the two text streams to match. If they don't, render
 * a word-level diff via `git diff --word-diff=plain` so the message points at
 * exactly which characters drifted.
 */

import { spawnSync } from "node:child_process";
import { promises as fs } from "node:fs";
import path from "node:path";

import JSZip from "jszip";

import { withTempDir } from "../../../lib/run-cli.ts";
import { NS } from "../../../lib/types.ts";
import type { ValidationResult } from "../../../lib/types.ts";
import { getElementsByTagNameNSAll, parseXml, serializeXml } from "../../../lib/xml-helpers.ts";

const ELEMENT_NODE = 1;

export interface RedliningOptions {
    unpackedDir: string;
    originalDocx: string;
    verbose?: boolean;
    /** Author whose tracked changes are being validated. Required. */
    author: string;
}

/**
 * Validate tracked changes in `unpackedDir/word/document.xml` against the
 * untouched copy inside `originalDocx`. Returns a `ValidationResult` whose
 * `valid` is true when (a) there are no tracked changes by `author` at all or
 * (b) every change by `author` is properly nested so the post-strip text
 * matches the original.
 *
 * Mirrors `RedliningValidator.validate()` from the Python source — including
 * its diagnostic messages, which downstream tooling greps for verbatim.
 */
export async function validateRedlining(options: RedliningOptions): Promise<ValidationResult> {
    const author = options.author;
    const verbose = options.verbose ?? false;
    const unpackedDir = options.unpackedDir;
    const originalDocx = options.originalDocx;

    const modifiedFile = path.join(unpackedDir, "word", "document.xml");

    let modifiedText: string;
    try {
        modifiedText = await fs.readFile(modifiedFile, "utf8");
    } catch {
        return failure(`FAILED - Modified document.xml not found at ${modifiedFile}`);
    }

    let modifiedRoot: Element;
    try {
        const parsed = parseXml(modifiedText);
        if (!parsed.documentElement) {
            throw new Error("missing document element");
        }
        modifiedRoot = parsed.documentElement;
    } catch {
        modifiedRoot = null as unknown as Element;
    }

    if (modifiedRoot) {
        const delElements = getElementsByTagNameNSAll(modifiedRoot, NS.W, "del");
        const insElements = getElementsByTagNameNSAll(modifiedRoot, NS.W, "ins");

        const authorDel = delElements.filter((el) => el.getAttributeNS(NS.W, "author") === author);
        const authorIns = insElements.filter((el) => el.getAttributeNS(NS.W, "author") === author);

        if (authorDel.length === 0 && authorIns.length === 0) {
            if (verbose) {
                console.log(`PASSED - No tracked changes by ${author} found.`);
            }
            return { valid: true, issues: [] };
        }
    }

    return withTempDir(async (tempDir) => {
        try {
            const data = await fs.readFile(originalDocx);
            const zip = await JSZip.loadAsync(data);
            await Promise.all(
                Object.values(zip.files).map(async (entry) => {
                    const target = path.join(tempDir, entry.name);
                    if (entry.dir) {
                        await fs.mkdir(target, { recursive: true });
                        return;
                    }
                    await fs.mkdir(path.dirname(target), { recursive: true });
                    const buf = await entry.async("nodebuffer");
                    await fs.writeFile(target, buf);
                }),
            );
        } catch (err) {
            const message = err instanceof Error ? err.message : String(err);
            return failure(`FAILED - Error unpacking original docx: ${message}`);
        }

        const originalFile = path.join(tempDir, "word", "document.xml");
        try {
            await fs.access(originalFile);
        } catch {
            return failure(`FAILED - Original document.xml not found in ${originalDocx}`);
        }

        let modifiedDoc: Document;
        let originalDoc: Document;
        try {
            modifiedDoc = parseXml(await fs.readFile(modifiedFile, "utf8"));
            originalDoc = parseXml(await fs.readFile(originalFile, "utf8"));
        } catch (err) {
            const message = err instanceof Error ? err.message : String(err);
            return failure(`FAILED - Error parsing XML files: ${message}`);
        }

        const modRoot = modifiedDoc.documentElement;
        const origRoot = originalDoc.documentElement;
        if (!modRoot || !origRoot) {
            return failure("FAILED - Error parsing XML files: missing document element");
        }

        removeAuthorTrackedChanges(origRoot, author);
        removeAuthorTrackedChanges(modRoot, author);

        const modifiedTxt = extractTextContent(modRoot);
        const originalTxt = extractTextContent(origRoot);

        if (modifiedTxt !== originalTxt) {
            const message = await generateDetailedDiff(originalTxt, modifiedTxt, author);
            console.log(message);
            return {
                valid: false,
                issues: [{ severity: "error", message, code: "redlining/mismatch" }],
            };
        }

        if (verbose) {
            console.log(`PASSED - All changes by ${author} are properly tracked`);
        }
        return { valid: true, issues: [] };
    });
}

function failure(message: string): ValidationResult {
    console.log(message);
    return { valid: false, issues: [{ severity: "error", message }] };
}

/**
 * Strip every `w:ins` authored by `author` (delete subtree) and unwrap every
 * `w:del` authored by `author` (re-insert children at the deletion's old
 * position, converting `w:delText` → `w:t` so the surviving text reads as
 * plain run text). Mirrors `_remove_author_tracked_changes` in Python — the
 * two-pass order matters: the ins pass must run first so a nested
 * `w:del` inside an author `w:ins` is dropped with the wrapper instead of
 * being unwrapped into the surrounding paragraph.
 */
export function removeAuthorTrackedChanges(root: Element, author: string): void {
    const allElements = collectAllElements(root);
    for (const parent of allElements) {
        const toRemove: Element[] = [];
        for (let child = parent.firstChild; child; child = child.nextSibling) {
            if (
                child.nodeType === ELEMENT_NODE &&
                (child as Element).namespaceURI === NS.W &&
                (child as Element).localName === "ins" &&
                (child as Element).getAttributeNS(NS.W, "author") === author
            ) {
                toRemove.push(child as Element);
            }
        }
        for (const elem of toRemove) {
            parent.removeChild(elem);
        }
    }

    const remaining = collectAllElements(root);
    for (const parent of remaining) {
        const children: Element[] = [];
        for (let child = parent.firstChild; child; child = child.nextSibling) {
            if (child.nodeType === ELEMENT_NODE) {
                children.push(child as Element);
            }
        }
        const targets: Element[] = children.filter(
            (el) => el.namespaceURI === NS.W && el.localName === "del" && el.getAttributeNS(NS.W, "author") === author,
        );
        for (const delElem of targets.slice().reverse()) {
            const delTexts = getElementsByTagNameNSAll(delElem, NS.W, "delText");
            for (const dt of delTexts) {
                renameElement(dt, "w:t");
            }
            const owned: Element[] = [];
            for (let child = delElem.firstChild; child; child = child.nextSibling) {
                if (child.nodeType === ELEMENT_NODE) {
                    owned.push(child as Element);
                }
            }
            // Iterate forward; insertBefore() preserves natural order (A, B, C)
            // when the reference node is fixed on `delElem`. Reversed iteration
            // would emit C, B, A and corrupt the stripped text stream.
            for (const child of owned) {
                parent.insertBefore(child, delElem);
            }
            parent.removeChild(delElem);
        }
    }
}

/**
 * Concatenate paragraph text from `root`. Mirrors `_extract_text_content`:
 * each `w:p` becomes the join of its descendant `w:t` text values, and only
 * non-empty paragraphs are emitted (joined by '\n').
 */
export function extractTextContent(root: Element): string {
    const paragraphs: string[] = [];
    for (const p of getElementsByTagNameNSAll(root, NS.W, "p")) {
        const parts: string[] = [];
        for (const t of getElementsByTagNameNSAll(p, NS.W, "t")) {
            if (t.textContent) {
                parts.push(t.textContent);
            }
        }
        const joined = parts.join("");
        if (joined) {
            paragraphs.push(joined);
        }
    }
    return paragraphs.join("\n");
}

async function generateDetailedDiff(originalText: string, modifiedText: string, author: string): Promise<string> {
    const errorParts: string[] = [
        `FAILED - Document text doesn't match after removing ${author}'s tracked changes`,
        "",
        "Likely causes:",
        "  1. Modified text inside another author's <w:ins> or <w:del> tags",
        "  2. Made edits without proper tracked changes",
        "  3. Didn't nest <w:del> inside <w:ins> when deleting another's insertion",
        "",
        "For pre-redlined documents, use correct patterns:",
        "  - To reject another's INSERTION: Nest <w:del> inside their <w:ins>",
        "  - To restore another's DELETION: Add new <w:ins> AFTER their <w:del>",
        "",
    ];

    const gitDiff = await getGitWordDiff(originalText, modifiedText);
    if (gitDiff) {
        errorParts.push("Differences:", "============", gitDiff);
    } else {
        errorParts.push("Unable to generate word diff (git not available)");
    }

    return errorParts.join("\n");
}

/**
 * Render `git diff --word-diff=plain` between two strings. Returns the
 * post-`@@` content lines on success, `null` on git failure. Mirrors
 * `_get_git_word_diff` — first attempt uses `--word-diff-regex=.` (per-char
 * granularity); if that yields nothing, retries without the regex (line
 * granularity).
 */
export async function getGitWordDiff(originalText: string, modifiedText: string): Promise<string | null> {
    return withTempDir((tempDir) => {
        const origFile = path.join(tempDir, "original.txt");
        const modFile = path.join(tempDir, "modified.txt");
        return Promise.all([fs.writeFile(origFile, originalText, "utf8"), fs.writeFile(modFile, modifiedText, "utf8")]).then(() => {
            const granular = runGitDiff(["diff", "--word-diff=plain", "--word-diff-regex=.", "-U0", "--no-index", origFile, modFile]);
            if (granular !== null && granular.trim() !== "") {
                const filtered = filterDiffContent(granular);
                if (filtered) {
                    return filtered;
                }
            }

            const wordLevel = runGitDiff(["diff", "--word-diff=plain", "-U0", "--no-index", origFile, modFile]);
            if (wordLevel !== null && wordLevel.trim() !== "") {
                return filterDiffContent(wordLevel);
            }

            return null;
        });
    });
}

function runGitDiff(args: string[]): string | null {
    try {
        const result = spawnSync("git", args, { encoding: "utf8" });
        if (result.error) {
            return null;
        }
        return result.stdout ?? "";
    } catch {
        return null;
    }
}

function filterDiffContent(stdout: string): string {
    const lines = stdout.split("\n");
    const contentLines: string[] = [];
    let inContent = false;
    for (const line of lines) {
        if (line.startsWith("@@")) {
            inContent = true;
            continue;
        }
        if (inContent && line.trim() !== "") {
            contentLines.push(line);
        }
    }
    return contentLines.join("\n");
}

function collectAllElements(root: Element): Element[] {
    const out: Element[] = [root];
    const list = root.getElementsByTagName("*");
    for (let i = 0; i < list.length; i += 1) {
        const item = list.item(i);
        if (item) {
            out.push(item as Element);
        }
    }
    return out;
}

/**
 * `@xmldom` does not allow live re-tagging of an element, so swap in a fresh
 * one carrying the same attributes and children. Used to convert
 * `w:delText` → `w:t` after stripping a tracked deletion.
 */
function renameElement(elem: Element, newQName: string): Element {
    const doc = elem.ownerDocument;
    if (!doc) {
        return elem;
    }
    const replacement = doc.createElementNS(NS.W, newQName);
    const attrs = elem.attributes;
    for (let i = 0; i < attrs.length; i += 1) {
        const attr = attrs.item(i);
        if (!attr) continue;
        if (attr.namespaceURI) {
            replacement.setAttributeNS(attr.namespaceURI, attr.name, attr.value);
        } else {
            replacement.setAttribute(attr.name, attr.value);
        }
    }
    while (elem.firstChild) {
        replacement.appendChild(elem.firstChild);
    }
    const parent = elem.parentNode;
    if (parent) {
        parent.replaceChild(replacement, elem);
    }
    return replacement;
}

/**
 * Class-based wrapper for API parity with Python's `RedliningValidator`.
 * The free functions above are the implementation; this class is a thin
 * adapter so callers that use `new RedliningValidator(opts)` work identically
 * to the Python source.
 */
export class RedliningValidator {
    constructor(private readonly options: RedliningOptions) {}

    validate(): Promise<ValidationResult> {
        return validateRedlining(this.options);
    }

    async repair(): Promise<number> {
        return 0;
    }
}

/** Re-export `serializeXml` so callers needing the post-strip XML can get it without re-importing helpers. */
export { serializeXml };
