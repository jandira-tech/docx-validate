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
 * Add comments to DOCX documents.
 *
 * 1:1 TypeScript port of `src/docx-validate/scripts/comment.py` (task #8).
 *
 * NOTE: `src/html-to-docx/src/tracking.ts` already implements a higher-level
 * comment / suggestion token system that knows about ranges, references, and
 * threading. This file is intentionally a verbatim port of the Python helper
 * — it operates directly on the unpacked OOXML parts (comments.xml,
 * commentsExtended.xml, commentsIds.xml, commentsExtensible.xml, plus
 * `_rels/document.xml.rels` and `[Content_Types].xml`) and is meant for
 * scripts / agents that need to mutate the on-disk DOCX bundle. Reach for
 * `tracking.ts` if you want a structured API instead.
 *
 * Usage:
 *   bunx tsx scripts/comment.ts unpacked/ 0 "Comment text"
 *   bunx tsx scripts/comment.ts unpacked/ 1 "Reply text" --parent 0
 *
 * Text should be pre-escaped XML (e.g., &amp; for &, &#x2019; for smart
 * quotes).
 *
 * After running, add markers to document.xml:
 *   <w:commentRangeStart w:id="0"/>
 *   ... commented content ...
 *   <w:commentRangeEnd w:id="0"/>
 *   <w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr><w:commentReference w:id="0"/></w:r>
 */

import { promises as fs } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

import { Command } from "commander";
import { commanderExitCode, runCli } from "../lib/run-cli";
import { parseXml, serializeXml } from "../lib/xml-helpers";

const TEMPLATE_DIR = path.join(path.dirname(fileURLToPath(import.meta.url)), "templates");

const NS: Record<string, string> = {
    w: "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    w14: "http://schemas.microsoft.com/office/word/2010/wordml",
    w15: "http://schemas.microsoft.com/office/word/2012/wordml",
    w16cid: "http://schemas.microsoft.com/office/word/2016/wordml/cid",
    w16cex: "http://schemas.microsoft.com/office/word/2018/wordml/cex",
};

const COMMENT_XML = `<w:comment w:id="{id}" w:author="{author}" w:date="{date}" w:initials="{initials}">
  <w:p w14:paraId="{para_id}" w14:textId="77777777">
    <w:r>
      <w:rPr><w:rStyle w:val="CommentReference"/></w:rPr>
      <w:annotationRef/>
    </w:r>
    <w:r>
      <w:rPr>
        <w:color w:val="000000"/>
        <w:sz w:val="20"/>
        <w:szCs w:val="20"/>
      </w:rPr>
      <w:t>{text}</w:t>
    </w:r>
  </w:p>
</w:comment>`;

const COMMENT_MARKER_TEMPLATE = `
Add to document.xml (markers must be direct children of w:p, never inside w:r):
  <w:commentRangeStart w:id="{cid}"/>
  <w:r>...</w:r>
  <w:commentRangeEnd w:id="{cid}"/>
  <w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr><w:commentReference w:id="{cid}"/></w:r>`;

const REPLY_MARKER_TEMPLATE = `
Nest markers inside parent {pid}'s markers (markers must be direct children of w:p, never inside w:r):
  <w:commentRangeStart w:id="{pid}"/><w:commentRangeStart w:id="{cid}"/>
  <w:r>...</w:r>
  <w:commentRangeEnd w:id="{cid}"/><w:commentRangeEnd w:id="{pid}"/>
  <w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr><w:commentReference w:id="{pid}"/></w:r>
  <w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr><w:commentReference w:id="{cid}"/></w:r>`;

function generateHexId(): string {
    const n = Math.floor(Math.random() * 0x7ffffffe);
    return n.toString(16).toUpperCase().padStart(8, "0");
}

const SMART_QUOTE_ENTITIES: Record<string, string> = {
    "“": "&#x201C;",
    "”": "&#x201D;",
    "‘": "&#x2018;",
    "’": "&#x2019;",
};

function encodeSmartQuotes(text: string): string {
    let out = text;
    for (const [char, entity] of Object.entries(SMART_QUOTE_ENTITIES)) {
        out = out.split(char).join(entity);
    }
    return out;
}

function escapeXml(text: string): string {
    return text
        .replace(/&/g, "&amp;")
        .replace(/</g, "&lt;")
        .replace(/>/g, "&gt;")
        .replace(/"/g, "&quot;")
        .replace(/'/g, "&apos;");
}

function formatTemplate(tpl: string, vars: Record<string, string | number>): string {
    return tpl.replace(/\{(\w+)\}/g, (_, key: string) => {
        if (key in vars) {
            return String(vars[key]);
        }
        return `{${key}}`;
    });
}

async function pathExists(p: string): Promise<boolean> {
    try {
        await fs.access(p);
        return true;
    } catch {
        return false;
    }
}

async function appendXml(xmlPath: string, rootTag: string, content: string): Promise<void> {
    const text = await fs.readFile(xmlPath, "utf-8");
    const dom = parseXml(text);
    const roots = dom.getElementsByTagName(rootTag);
    if (roots.length === 0) {
        throw new Error(`Root tag <${rootTag}> not found in ${xmlPath}`);
    }
    const root = roots.item(0)!;
    const nsAttrs = Object.entries(NS)
        .map(([k, v]) => `xmlns:${k}="${v}"`)
        .join(" ");
    const wrapperDom = parseXml(`<root ${nsAttrs}>${content}</root>`);
    const children = wrapperDom.documentElement.childNodes;
    for (let i = 0; i < children.length; i += 1) {
        const child = children.item(i);
        if (child && child.nodeType === 1 /* ELEMENT_NODE */) {
            root.appendChild(dom.importNode(child, true));
        }
    }
    const output = encodeSmartQuotes(serializeXml(dom, "UTF-8"));
    await fs.writeFile(xmlPath, output, "utf-8");
}

async function findParaId(commentsPath: string, commentId: number): Promise<string | null> {
    const text = await fs.readFile(commentsPath, "utf-8");
    const dom = parseXml(text);
    const comments = dom.getElementsByTagName("w:comment");
    for (let i = 0; i < comments.length; i += 1) {
        const c = comments.item(i)!;
        if (c.getAttribute("w:id") === String(commentId)) {
            const ps = c.getElementsByTagName("w:p");
            for (let j = 0; j < ps.length; j += 1) {
                const p = ps.item(j)!;
                const pid = p.getAttribute("w14:paraId");
                if (pid) {
                    return pid;
                }
            }
        }
    }
    return null;
}

/** Return true if `commentsPath` already contains a `<w:comment>` with the
 * given `w:id`. Used to preflight `addComment()` before mutating any part. */
async function commentIdExists(commentsPath: string, commentId: number): Promise<boolean> {
    const text = await fs.readFile(commentsPath, "utf-8");
    const dom = parseXml(text);
    const comments = dom.getElementsByTagName("w:comment");
    for (let i = 0; i < comments.length; i += 1) {
        const c = comments.item(i)!;
        if (c.getAttribute("w:id") === String(commentId)) {
            return true;
        }
    }
    return false;
}

async function getNextRid(relsPath: string): Promise<number> {
    const text = await fs.readFile(relsPath, "utf-8");
    const dom = parseXml(text);
    let maxRid = 0;
    const rels = dom.getElementsByTagName("Relationship");
    for (let i = 0; i < rels.length; i += 1) {
        const rel = rels.item(i)!;
        const rid = rel.getAttribute("Id");
        if (rid && rid.startsWith("rId")) {
            const n = Number.parseInt(rid.slice(3), 10);
            if (Number.isFinite(n)) {
                maxRid = Math.max(maxRid, n);
            }
        }
    }
    return maxRid + 1;
}

async function hasRelationship(relsPath: string, target: string): Promise<boolean> {
    const text = await fs.readFile(relsPath, "utf-8");
    const dom = parseXml(text);
    const rels = dom.getElementsByTagName("Relationship");
    for (let i = 0; i < rels.length; i += 1) {
        const rel = rels.item(i)!;
        if (rel.getAttribute("Target") === target) {
            return true;
        }
    }
    return false;
}

async function hasContentType(ctPath: string, partName: string): Promise<boolean> {
    const text = await fs.readFile(ctPath, "utf-8");
    const dom = parseXml(text);
    const overrides = dom.getElementsByTagName("Override");
    for (let i = 0; i < overrides.length; i += 1) {
        const override = overrides.item(i)!;
        if (override.getAttribute("PartName") === partName) {
            return true;
        }
    }
    return false;
}

async function ensureCommentRelationships(unpackedDir: string): Promise<void> {
    const relsPath = path.join(unpackedDir, "word", "_rels", "document.xml.rels");
    if (!(await pathExists(relsPath))) {
        return;
    }

    if (await hasRelationship(relsPath, "comments.xml")) {
        return;
    }

    const text = await fs.readFile(relsPath, "utf-8");
    const dom = parseXml(text);
    const root = dom.documentElement;
    let nextRid = await getNextRid(relsPath);

    const rels: Array<[string, string]> = [
        ["http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments", "comments.xml"],
        ["http://schemas.microsoft.com/office/2011/relationships/commentsExtended", "commentsExtended.xml"],
        ["http://schemas.microsoft.com/office/2016/09/relationships/commentsIds", "commentsIds.xml"],
        ["http://schemas.microsoft.com/office/2018/08/relationships/commentsExtensible", "commentsExtensible.xml"],
    ];

    for (const [relType, target] of rels) {
        const rel = dom.createElement("Relationship");
        rel.setAttribute("Id", `rId${nextRid}`);
        rel.setAttribute("Type", relType);
        rel.setAttribute("Target", target);
        root.appendChild(rel);
        nextRid += 1;
    }

    await fs.writeFile(relsPath, serializeXml(dom, "UTF-8"), "utf-8");
}

async function ensureCommentContentTypes(unpackedDir: string): Promise<void> {
    const ctPath = path.join(unpackedDir, "[Content_Types].xml");
    if (!(await pathExists(ctPath))) {
        return;
    }

    if (await hasContentType(ctPath, "/word/comments.xml")) {
        return;
    }

    const text = await fs.readFile(ctPath, "utf-8");
    const dom = parseXml(text);
    const root = dom.documentElement;

    const overrides: Array<[string, string]> = [
        ["/word/comments.xml", "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"],
        ["/word/commentsExtended.xml", "application/vnd.openxmlformats-officedocument.wordprocessingml.commentsExtended+xml"],
        ["/word/commentsIds.xml", "application/vnd.openxmlformats-officedocument.wordprocessingml.commentsIds+xml"],
        ["/word/commentsExtensible.xml", "application/vnd.openxmlformats-officedocument.wordprocessingml.commentsExtensible+xml"],
    ];

    for (const [partName, contentType] of overrides) {
        const override = dom.createElement("Override");
        override.setAttribute("PartName", partName);
        override.setAttribute("ContentType", contentType);
        root.appendChild(override);
    }

    await fs.writeFile(ctPath, serializeXml(dom, "UTF-8"), "utf-8");
}

export interface AddCommentOptions {
    unpackedDir: string;
    commentId: number;
    text: string;
    parent?: number;
    /** Comment author. Required. */
    author: string;
    /** Author initials. Required (Word renders this in the comment thumbnail). */
    initials: string;
    /** Override the timestamp; defaults to `new Date()`. Format: `YYYY-MM-DDTHH:MM:SSZ`. */
    date?: Date;
}

export interface AddCommentResult {
    paraId: string;
    message: string;
}

/**
 * Add a comment (or reply) to an unpacked DOCX.
 *
 * On success, returns the generated `paraId` and a human-readable status
 * message. On failure (missing `word/` directory, missing parent comment),
 * returns `paraId = ""` and a message starting with `Error:`.
 */
export async function addComment(opts: AddCommentOptions): Promise<AddCommentResult> {
    const { unpackedDir, commentId, text, parent, author, initials, date } = opts;

    const word = path.join(unpackedDir, "word");
    if (!(await pathExists(word))) {
        return { paraId: "", message: `Error: ${word} not found` };
    }

    const paraId = generateHexId();
    const durableId = generateHexId();
    const ts = formatUtcTimestamp(date ?? new Date());

    const comments = path.join(word, "comments.xml");
    const firstComment = !(await pathExists(comments));

    // Preflight validation: check duplicate commentId and resolve `parent`
    // BEFORE we mutate any part. Without this, an existing comments.xml could
    // be mutated and then the parent-not-found branch leaves an orphan comment
    // behind on a failed call.
    if (!firstComment) {
        if (await commentIdExists(comments, commentId)) {
            return {
                paraId: "",
                message: `Error: Comment id ${commentId} already exists`,
            };
        }
        if (parent !== undefined) {
            const parentPara = await findParaId(comments, parent);
            if (!parentPara) {
                return {
                    paraId: "",
                    message: `Error: Parent comment ${parent} not found`,
                };
            }
        }
    } else if (parent !== undefined) {
        // First comment in the document but a parent was requested — the
        // parent cannot exist yet, so reject before creating template files.
        return { paraId: "", message: `Error: Parent comment ${parent} not found` };
    }

    if (firstComment) {
        await fs.copyFile(path.join(TEMPLATE_DIR, "comments.xml"), comments);
        await ensureCommentRelationships(unpackedDir);
        await ensureCommentContentTypes(unpackedDir);
    }
    await appendXml(
        comments,
        "w:comments",
        formatTemplate(COMMENT_XML, {
            id: commentId,
            author: escapeXml(author),
            date: ts,
            initials: escapeXml(initials),
            para_id: paraId,
            text,
        }),
    );

    const ext = path.join(word, "commentsExtended.xml");
    if (!(await pathExists(ext))) {
        await fs.copyFile(path.join(TEMPLATE_DIR, "commentsExtended.xml"), ext);
    }
    if (parent !== undefined) {
        // Already preflighted above; resolve the actual paraId here for the link.
        const parentPara = await findParaId(comments, parent);
        if (!parentPara) {
            // Defensive: should be unreachable given preflight, but keep a
            // structured failure rather than a silent bad ref.
            return {
                paraId: "",
                message: `Error: Parent comment ${parent} not found`,
            };
        }
        await appendXml(ext, "w15:commentsEx", `<w15:commentEx w15:paraId="${paraId}" w15:paraIdParent="${parentPara}" w15:done="0"/>`);
    } else {
        await appendXml(ext, "w15:commentsEx", `<w15:commentEx w15:paraId="${paraId}" w15:done="0"/>`);
    }

    const ids = path.join(word, "commentsIds.xml");
    if (!(await pathExists(ids))) {
        await fs.copyFile(path.join(TEMPLATE_DIR, "commentsIds.xml"), ids);
    }
    await appendXml(ids, "w16cid:commentsIds", `<w16cid:commentId w16cid:paraId="${paraId}" w16cid:durableId="${durableId}"/>`);

    const extensible = path.join(word, "commentsExtensible.xml");
    if (!(await pathExists(extensible))) {
        await fs.copyFile(path.join(TEMPLATE_DIR, "commentsExtensible.xml"), extensible);
    }
    await appendXml(
        extensible,
        "w16cex:commentsExtensible",
        `<w16cex:commentExtensible w16cex:durableId="${durableId}" w16cex:dateUtc="${ts}"/>`,
    );

    const action = parent !== undefined ? "reply" : "comment";
    return {
        paraId,
        message: `Added ${action} ${commentId} (para_id=${paraId})`,
    };
}

function formatUtcTimestamp(d: Date): string {
    const pad = (n: number) => String(n).padStart(2, "0");
    return (
        `${d.getUTCFullYear()}-${pad(d.getUTCMonth() + 1)}-${pad(d.getUTCDate())}` +
        `T${pad(d.getUTCHours())}:${pad(d.getUTCMinutes())}:${pad(d.getUTCSeconds())}Z`
    );
}

export function buildCommentCommand(): Command {
    const cmd = new Command();
    cmd.name("comment")
        .description("Add comments to DOCX documents")
        .argument("<unpackedDir>", "Unpacked DOCX directory")
        .argument("<commentId>", "Comment ID (must be unique)")
        .argument("<text>", "Comment text")
        .requiredOption("--author <author>", "Author name (required)")
        .requiredOption("--initials <initials>", "Author initials (required; Word renders these in the comment thumbnail)")
        .option("--parent <parent>", "Parent comment ID (for replies)");
    return cmd;
}

interface CliOptions {
    author: string;
    initials: string;
    parent?: string;
}

export async function runCommentFromArgv(argv: readonly string[]): Promise<number> {
    const cmd = buildCommentCommand();
    cmd.exitOverride();
    try {
        cmd.parse(argv as string[], { from: "user" });
    } catch (err) {
        return commanderExitCode(err);
    }
    const opts = cmd.opts<CliOptions>();
    const [unpackedDir, commentIdRaw, text] = cmd.args;

    const commentId = Number.parseInt(commentIdRaw, 10);
    if (!Number.isFinite(commentId)) {
        process.stderr.write(`Error: commentId must be an integer (got ${commentIdRaw})\n`);
        return 1;
    }
    const parent = opts.parent !== undefined ? Number.parseInt(opts.parent, 10) : undefined;
    if (parent !== undefined && !Number.isFinite(parent)) {
        process.stderr.write(`Error: --parent must be an integer (got ${opts.parent})\n`);
        return 1;
    }

    const result = await addComment({
        unpackedDir,
        commentId,
        text,
        author: opts.author,
        initials: opts.initials,
        parent,
    });
    process.stdout.write(`${result.message}\n`);
    if (result.message.includes("Error")) {
        return 1;
    }
    if (parent !== undefined) {
        process.stdout.write(`${formatTemplate(REPLY_MARKER_TEMPLATE, { pid: parent, cid: commentId })}\n`);
    } else {
        process.stdout.write(`${formatTemplate(COMMENT_MARKER_TEMPLATE, { cid: commentId })}\n`);
    }
    return 0;
}

runCli(import.meta.url, () => runCommentFromArgv(process.argv.slice(2)));
