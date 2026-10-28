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
import { MINIMAL_PARTS } from "./support/minimal-docx";

const W_NS = `xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"`;
const W14_NS = `xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"`;
const W16CID_NS = `xmlns:w16cid="http://schemas.microsoft.com/office/word/2016/wordml/cid"`;
const W16CEX_NS = `xmlns:w16cex="http://schemas.microsoft.com/office/word/2018/wordml/cex"`;

const writeFailureCases = [
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
] as const;

const writeRepairParts = async (dir: string, method: (typeof writeFailureCases)[number][0]): Promise<string[]> => {
    const wordDir = path.join(dir, "word");
    await fs.mkdir(wordDir);
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
        return ["commentsIds.xml", "commentsExtensible.xml"].map((name) => path.join(wordDir, name));
    } else {
        const parts = ["a.xml", "b.xml", "c.xml"].map((name) => path.join(wordDir, name));
        const attribute = method === "repairParaId" ? 'w14:paraId="FFFFFFFF" w14:textId="11111111"' : 'w16cid:durableId="FFFFFFFF"';
        for (const part of parts) {
            await fs.writeFile(part, `<?xml version="1.0"?><root ${W14_NS} ${W16CID_NS} ${attribute}/>`);
        }
        return parts;
    }
};

describe("DOCX repair write failures after preflight (#909)", () => {
    it.each(writeFailureCases)("%s through %s returns only written repairs when write %i throws", async (method, entry, failedWrite) => {
        await withTempDir(async (dir) => {
            const parts = await writeRepairParts(dir, method);
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

    it.each(writeFailureCases)(
        "%s through %s returns only written repairs when write %i is refused (#920)",
        async (method, entry, failedWrite) => {
            await withTempDir(async (dir) => {
                const parts = await writeRepairParts(dir, method);
                const before = await Promise.all(parts.map((part) => fs.readFile(part)));
                const v = new DOCXSchemaValidator({ unpackedDir: dir });
                if (method !== "repairCommentThreading") expect(v.xmlFiles).toEqual(parts);
                const writeXmlPart = xmlPartIo.writeXmlPart;
                let writes = 0;
                const spy = vi.spyOn(xmlPartIo, "writeXmlPart").mockImplementation(async (...args) => {
                    writes += 1;
                    if (writes === failedWrite) return false;
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
        },
    );
});

const writeDurableIdParts = async (dir: string, extensibleFirst = false): Promise<string[]> => {
    for (const [name, text] of Object.entries(MINIMAL_PARTS)) {
        await fs.mkdir(path.dirname(path.join(dir, name)), { recursive: true });
        await fs.writeFile(path.join(dir, name), text);
    }
    const wordDir = path.join(dir, "word");
    await fs.writeFile(
        path.join(wordDir, "comments.xml"),
        `<?xml version="1.0"?><w:comments ${W_NS} ${W14_NS}>` +
            `<w:comment w:id="0" w:author="A"><w:p w14:paraId="00000010" w14:textId="00000010"/></w:comment>` +
            `<w:comment w:id="1" w:author="A"><w:p w14:paraId="00000011" w14:textId="00000011"/></w:comment>` +
            `</w:comments>`,
    );
    // The `commentsIds` and `commentsExtensible` roots stored under unrelated
    // names, related by the main document's sidecar: the hex rewrite renames
    // `FFFFFFFF` in both, so they must stay paired. `extensibleFirst` swaps
    // the names, so the `commentsExtensible` root is written first.
    const [idsName, extensibleName] = extensibleFirst ? ["z.xml", "a.xml"] : ["a.xml", "z.xml"];
    await fs.writeFile(
        path.join(wordDir, idsName),
        `<?xml version="1.0"?><w16cid:commentsIds ${W16CID_NS}>` +
            `<w16cid:commentId w16cid:paraId="00000010" w16cid:durableId="FFFFFFFF"/>` +
            `<w16cid:commentId w16cid:paraId="00000011" w16cid:durableId="00000002"/>` +
            `</w16cid:commentsIds>`,
    );
    await fs.writeFile(
        path.join(wordDir, extensibleName),
        `<?xml version="1.0"?><w16cex:commentsExtensible ${W16CEX_NS}>` +
            `<w16cex:commentExtensible w16cex:durableId="FFFFFFFF"/>` +
            `<w16cex:commentExtensible w16cex:durableId="00000002"/>` +
            `</w16cex:commentsExtensible>`,
    );
    await fs.writeFile(
        path.join(wordDir, "numbering.xml"),
        `<?xml version="1.0"?><w:numbering ${W_NS} ${W16CID_NS}>` +
            `<w:abstractNum w:abstractNumId="0" w16cid:durableId="4294967295"/>` +
            `</w:numbering>`,
    );
    await fs.writeFile(
        path.join(wordDir, "_rels", "document.xml.rels"),
        MINIMAL_PARTS["word/_rels/document.xml.rels"].replace(
            "</Relationships>",
            `<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="comments.xml"/>` +
                `<Relationship Id="rId3" Type="http://schemas.microsoft.com/office/word/2016/09/relationships/commentsIds" Target="${idsName}"/>` +
                `<Relationship Id="rId4" Type="http://schemas.microsoft.com/office/word/2018/08/relationships/commentsExtensible" Target="${extensibleName}"/>` +
                "</Relationships>",
        ),
    );
    // In write order: the hex part, the decimal part, the second hex part.
    return ["a.xml", "numbering.xml", "z.xml"].map((name) => path.join(wordDir, name));
};

const hexDurableIds = async (part: string, attribute: string): Promise<string[]> =>
    [...(await fs.readFile(part, "utf8")).matchAll(new RegExp(`${attribute}="([^"]+)"`, "g"))].map((match) => match[1]!);

const decimalDurableId = async (part: string): Promise<string> =>
    (await fs.readFile(part, "utf8")).match(/w16cid:durableId="([^"]+)"/)?.[1] ?? "";

// [entry, failedWrite (1 = a.xml, 2 = numbering.xml, 3 = z.xml), failure, count]
const numberingSplitCases = [
    ["direct", 2, "refused", 2],
    ["direct", 2, "thrown", 2],
    ["direct", 1, "refused", 1],
    ["direct", 1, "thrown", 1],
    ["direct", 3, "refused", 2],
    ["repair", 2, "refused", 2],
    ["repair", 2, "thrown", 2],
    ["repair", 1, "refused", 1],
    ["repair", 1, "thrown", 1],
    ["repair", 3, "refused", 2],
] as const;

describe("DOCX repair durableId hex rename and decimal renumbering (#925)", () => {
    it.each(numberingSplitCases)(
        "%s: write %i %s keeps the hex rename paired and the decimal renumbering running",
        async (entry, failedWrite, failure, count) => {
            await withTempDir(async (dir) => {
                const parts = await writeDurableIdParts(dir);
                const [aXml, numberingXml, zXml] = parts;
                const before = await Promise.all(parts.map((part) => fs.readFile(part)));
                const v = new DOCXSchemaValidator({ unpackedDir: dir });
                // The write order the loop sees: hex part, decimal part, hex part.
                expect(v.xmlFiles.filter((file) => parts.includes(file))).toEqual(parts);
                const writeXmlPart = xmlPartIo.writeXmlPart;
                let writes = 0;
                const spy = vi.spyOn(xmlPartIo, "writeXmlPart").mockImplementation(async (...args) => {
                    writes += 1;
                    if (writes === failedWrite) {
                        if (failure === "thrown") throw Object.assign(new Error("permission denied"), { code: "EACCES" });
                        return false;
                    }
                    return writeXmlPart(...args);
                });
                try {
                    await expect(entry === "repair" ? v.repair() : v.repairDurableId()).resolves.toBe(count);
                    const after = await Promise.all(parts.map((part) => fs.readFile(part)));
                    const [aIds, zIds] = await Promise.all([
                        hexDurableIds(aXml, "w16cid:durableId"),
                        hexDurableIds(zXml, "w16cex:durableId"),
                    ]);
                    const numberingId = await decimalDurableId(numberingXml);
                    if (failedWrite === 2) {
                        // A failed decimal-only write neither stops the hex
                        // rewrite nor lands: numbering keeps its bytes.
                        expect(after[1]).toEqual(before[1]);
                        expect(aIds[0]).toMatch(/^[0-9A-F]{8}$/);
                        expect(aIds[0]).not.toBe("FFFFFFFF");
                        expect(zIds[0]).toBe(aIds[0]);
                    } else if (failedWrite === 1) {
                        // A failed hex write stops the remaining hex writes, so
                        // both hex parts keep the old value; the decimal
                        // renumbering still runs.
                        expect(after[0]).toEqual(before[0]);
                        expect(after[2]).toEqual(before[2]);
                        expect(numberingId).toMatch(/^[1-9][0-9]*$/);
                        expect(Number(numberingId)).toBeLessThan(4294967295);
                    } else {
                        // The #920 stop on the last hex part, with the decimal
                        // renumbering done before it.
                        expect(after[0]).not.toEqual(before[0]);
                        expect(aIds[0]).toMatch(/^[0-9A-F]{8}$/);
                        expect(aIds[0]).not.toBe("FFFFFFFF");
                        expect(after[2]).toEqual(before[2]);
                        expect(numberingId).toMatch(/^[1-9][0-9]*$/);
                        expect(Number(numberingId)).toBeLessThan(4294967295);
                    }
                    // The in-range sibling keeps its value in every hex part written.
                    if (failedWrite !== 1) {
                        expect(aIds[1]).toBe("00000002");
                        expect(zIds[1]).toBe("00000002");
                    }
                } finally {
                    spy.mockRestore();
                }
            });
        },
    );

    it.each([
        ["direct", "refused"],
        ["direct", "thrown"],
        ["repair", "refused"],
        ["repair", "thrown"],
    ] as const)("%s: a %s write of the commentsExtensible part, written first, stops the commentsIds rename", async (entry, failure) => {
        await withTempDir(async (dir) => {
            const parts = await writeDurableIdParts(dir, true);
            const before = await Promise.all(parts.map((part) => fs.readFile(part)));
            const v = new DOCXSchemaValidator({ unpackedDir: dir });
            expect(v.xmlFiles.filter((file) => parts.includes(file))).toEqual(parts);
            const writeXmlPart = xmlPartIo.writeXmlPart;
            let writes = 0;
            const spy = vi.spyOn(xmlPartIo, "writeXmlPart").mockImplementation(async (...args) => {
                writes += 1;
                if (writes === 1) {
                    if (failure === "thrown") throw Object.assign(new Error("permission denied"), { code: "EACCES" });
                    return false;
                }
                return writeXmlPart(...args);
            });
            try {
                // Only numbering's decimal renumbering lands; the
                // `commentsIds` part keeps `FFFFFFFF`, still paired with the
                // unwritten `commentsExtensible` part.
                await expect(entry === "repair" ? v.repair() : v.repairDurableId()).resolves.toBe(1);
                const after = await Promise.all(parts.map((part) => fs.readFile(part)));
                expect(after[0]).toEqual(before[0]);
                expect(after[2]).toEqual(before[2]);
                const numberingId = await decimalDurableId(parts[1]!);
                expect(numberingId).toMatch(/^[1-9][0-9]*$/);
                expect(Number(numberingId)).toBeLessThan(4294967295);
            } finally {
                spy.mockRestore();
            }
        });
    });
});
