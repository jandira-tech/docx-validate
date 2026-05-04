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

import { promises as fs } from "node:fs";
import path from "node:path";

import { describe, expect, it } from "vitest";
import { withTempDir } from "../src/lib/run-cli";
import { parseXml } from "../src/lib/xml-helpers";
import { addComment } from "../src/scripts/comment";

const RELS_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>`;

const CT_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>`;

async function makeFixture(dir: string): Promise<void> {
    await fs.mkdir(path.join(dir, "word", "_rels"), { recursive: true });
    await fs.writeFile(path.join(dir, "word", "_rels", "document.xml.rels"), RELS_XML, "utf-8");
    await fs.writeFile(path.join(dir, "[Content_Types].xml"), CT_XML, "utf-8");
}

describe("addComment", () => {
    it("adds a top-level comment, seeds parts, and registers rels + content types", async () => {
        await withTempDir(async (dir) => {
            await makeFixture(dir);

            const result = await addComment({
                unpackedDir: dir,
                commentId: 0,
                text: "Hello world",
                author: "Test",
                initials: "C",
                date: new Date(Date.UTC(2026, 4, 3, 12, 0, 0)),
            });
            expect(result.message).toMatch(/^Added comment 0 \(para_id=[0-9A-F]{8}\)$/);
            expect(result.paraId).toMatch(/^[0-9A-F]{8}$/);

            for (const name of ["comments.xml", "commentsExtended.xml", "commentsIds.xml", "commentsExtensible.xml"]) {
                await fs.access(path.join(dir, "word", name));
            }

            const commentsXml = await fs.readFile(path.join(dir, "word", "comments.xml"), "utf-8");
            const commentsDom = parseXml(commentsXml);
            const comments = commentsDom.getElementsByTagName("w:comment");
            expect(comments.length).toBe(1);
            const c0 = comments.item(0)!;
            expect(c0.getAttribute("w:id")).toBe("0");
            expect(c0.getAttribute("w:author")).toBe("Test");
            expect(c0.getAttribute("w:initials")).toBe("C");
            expect(c0.getAttribute("w:date")).toBe("2026-05-03T12:00:00Z");
            expect(commentsXml).toContain("Hello world");

            const relsXml = await fs.readFile(path.join(dir, "word", "_rels", "document.xml.rels"), "utf-8");
            expect(relsXml).toContain('Target="comments.xml"');
            expect(relsXml).toContain('Target="commentsExtended.xml"');
            expect(relsXml).toContain('Target="commentsIds.xml"');
            expect(relsXml).toContain('Target="commentsExtensible.xml"');

            const ctXml = await fs.readFile(path.join(dir, "[Content_Types].xml"), "utf-8");
            expect(ctXml).toContain('PartName="/word/comments.xml"');
            expect(ctXml).toContain('PartName="/word/commentsExtended.xml"');

            const reply = await addComment({
                unpackedDir: dir,
                commentId: 1,
                text: "Reply text",
                parent: 0,
                author: "Reviewer",
                initials: "R",
                date: new Date(Date.UTC(2026, 4, 3, 12, 5, 0)),
            });
            expect(reply.message).toMatch(/^Added reply 1 \(para_id=[0-9A-F]{8}\)$/);

            const extXml = await fs.readFile(path.join(dir, "word", "commentsExtended.xml"), "utf-8");
            const extDom = parseXml(extXml);
            const exItems = extDom.getElementsByTagName("w15:commentEx");
            expect(exItems.length).toBe(2);
            const replyEx = exItems.item(1)!;
            expect(replyEx.getAttribute("w15:paraId")).toBe(reply.paraId);
            expect(replyEx.getAttribute("w15:paraIdParent")).toBe(result.paraId);
        });
    });

    it("reports an error when the unpacked DOCX has no word/ directory", async () => {
        await withTempDir(async (dir) => {
            const result = await addComment({
                unpackedDir: dir,
                commentId: 0,
                text: "x",
                author: "TestAuthor",
                initials: "TA",
            });
            expect(result.paraId).toBe("");
            expect(result.message).toMatch(/^Error: .*word not found$/);
        });
    });

    it("reports an error when the parent comment id is unknown", async () => {
        await withTempDir(async (dir) => {
            await makeFixture(dir);
            await addComment({
                unpackedDir: dir,
                commentId: 0,
                text: "first",
                author: "TestAuthor",
                initials: "TA",
            });
            const reply = await addComment({
                unpackedDir: dir,
                commentId: 1,
                text: "orphan reply",
                parent: 99,
                author: "TestAuthor",
                initials: "TA",
            });
            expect(reply.paraId).toBe("");
            expect(reply.message).toBe("Error: Parent comment 99 not found");
        });
    });
});
