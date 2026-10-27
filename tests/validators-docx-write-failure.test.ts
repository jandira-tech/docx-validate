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

import { describe, expect, it, vi } from "vitest";
import { withTempDir } from "../src/lib/run-cli";
import * as xmlPartIo from "../src/lib/xml-part-io";
import { DOCXSchemaValidator } from "../src/scripts/office/validators/docx";

const W_NS = `xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"`;
const W14_NS = `xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"`;
const W16CID_NS = `xmlns:w16cid="http://schemas.microsoft.com/office/word/2016/wordml/cid"`;
const W16CEX_NS = `xmlns:w16cex="http://schemas.microsoft.com/office/word/2018/wordml/cex"`;

describe("DOCX repair write failures after preflight (#909)", () => {
    it.each([
        ["repairCommentThreading", "direct", 1],
        ["repairCommentThreading", "direct", 2],
        ["repairCommentThreading", "repair", 1],
        ["repairCommentThreading", "repair", 2],
        ["repairParaId", "direct", 1],
        ["repairParaId", "direct", 2],
        ["repairParaId", "repair", 1],
        ["repairParaId", "repair", 2],
        ["repairDurableId", "direct", 1],
        ["repairDurableId", "direct", 2],
        ["repairDurableId", "repair", 1],
        ["repairDurableId", "repair", 2],
    ] as const)("%s through %s returns only written repairs when write %i throws", async (method, entry, failedWrite) => {
        await withTempDir(async (dir) => {
            const wordDir = path.join(dir, "word");
            await fs.mkdir(wordDir);
            let parts: string[];
            if (method === "repairCommentThreading") {
                await fs.writeFile(
                    path.join(wordDir, "comments.xml"),
                    `<?xml version="1.0"?><w:comments ${W_NS} ${W14_NS}>` +
                        `<w:comment w:id="0" w:author="A" w:date="2026-01-01T00:00:00Z" w:initials="A">` +
                        `<w:p w14:paraId="11111111" w14:textId="11111111"/></w:comment></w:comments>`,
                );
                await fs.writeFile(
                    path.join(wordDir, "commentsIds.xml"),
                    `<?xml version="1.0"?><w16cid:commentsIds ${W16CID_NS}>` +
                        `<w16cid:commentId w16cid:paraId="11111111" w16cid:durableId="22222222"/></w16cid:commentsIds>`,
                );
                await fs.writeFile(
                    path.join(wordDir, "commentsExtensible.xml"),
                    `<?xml version="1.0"?><w16cex:commentsExtensible ${W16CEX_NS}>` +
                        `<w16cex:commentExtensible w16cex:durableId="22222222"/></w16cex:commentsExtensible>`,
                );
                parts = ["commentsIds.xml", "commentsExtensible.xml"].map((name) => path.join(wordDir, name));
            } else {
                parts = ["a.xml", "b.xml", "c.xml"].map((name) => path.join(wordDir, name));
                const attribute = method === "repairParaId" ? 'w14:paraId="FFFFFFFF" w14:textId="11111111"' : 'w16cid:durableId="FFFFFFFF"';
                for (const part of parts) {
                    await fs.writeFile(part, `<?xml version="1.0"?><root ${W14_NS} ${W16CID_NS} ${attribute}/>`);
                }
            }
            const before = await Promise.all(parts.map((part) => fs.readFile(part)));
            const v = new DOCXSchemaValidator({ unpackedDir: dir });
            if (method !== "repairCommentThreading") expect(v.xmlFiles).toEqual(parts);
            const writeXmlPart = xmlPartIo.writeXmlPart;
            let writes = 0;
            const spy = vi.spyOn(xmlPartIo, "writeXmlPart").mockImplementation(async (...args) => {
                writes += 1;
                if (writes === failedWrite) throw Object.assign(new Error("permission denied"), { code: "EACCES" });
                return writeXmlPart(...args);
            });
            try {
                await expect(v[entry === "repair" ? "repair" : method]()).resolves.toBe(failedWrite - 1);
                expect(spy).toHaveBeenCalledTimes(failedWrite);
                const after = await Promise.all(parts.map((part) => fs.readFile(part)));
                expect(after.slice(failedWrite - 1)).toEqual(before.slice(failedWrite - 1));
                if (failedWrite === 2) expect(after[0]).not.toEqual(before[0]);
            } finally {
                spy.mockRestore();
            }
        });
    });
});
