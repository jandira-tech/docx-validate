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
import JSZip from "jszip";

import { describe, expect, it } from "vitest";
import { withTempDir } from "../src/lib/run-cli";
import {
    DOCXSchemaValidator,
    WORD_2006_NAMESPACE,
    WORD_PARAGRAPH_NAMESPACES,
    WORD_STRICT_NAMESPACE,
} from "../src/scripts/office/validators/docx";

const W_NS = `xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"`;
const W14_NS = `xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"`;
const W16CID_NS = `xmlns:w16cid="http://schemas.microsoft.com/office/word/2016/wordml/cid"`;

async function writeFile(p: string, content: string): Promise<void> {
    await fs.mkdir(path.dirname(p), { recursive: true });
    await fs.writeFile(p, content, "utf-8");
}

function wrapDocument(body: string, extraNs = ""): string {
    return `<?xml version="1.0"?><w:document ${W_NS} ${extraNs}><w:body>${body}</w:body></w:document>`;
}

describe("DOCXSchemaValidator", () => {
    describe("validateWhitespacePreservation", () => {
        it("flags <w:t> with leading whitespace and no xml:space=preserve", async () => {
            await withTempDir(async (dir) => {
                await writeFile(path.join(dir, "word", "document.xml"), wrapDocument(`<w:p><w:r><w:t> hello</w:t></w:r></w:p>`));
                const v = new DOCXSchemaValidator({ unpackedDir: dir });
                const result = await v.validateWhitespacePreservation();
                expect(result.valid).toBe(false);
                expect(result.issues[0].code).toBe("ws-missing-preserve");
            });
        });

        it("passes when xml:space=preserve is set", async () => {
            await withTempDir(async (dir) => {
                const xmlSpace = `xmlns:xml="http://www.w3.org/XML/1998/namespace"`;
                await writeFile(
                    path.join(dir, "word", "document.xml"),
                    `<?xml version="1.0"?><w:document ${W_NS} ${xmlSpace}><w:body>` +
                        `<w:p><w:r><w:t xml:space="preserve"> hello</w:t></w:r></w:p>` +
                        `</w:body></w:document>`,
                );
                const v = new DOCXSchemaValidator({ unpackedDir: dir });
                const result = await v.validateWhitespacePreservation();
                expect(result.valid).toBe(true);
            });
        });
    });

    describe("validateDeletions", () => {
        it("flags <w:t> inside <w:del>", async () => {
            await withTempDir(async (dir) => {
                await writeFile(
                    path.join(dir, "word", "document.xml"),
                    wrapDocument(`<w:p><w:del w:id="1"><w:r><w:t>bad</w:t></w:r></w:del></w:p>`),
                );
                const v = new DOCXSchemaValidator({ unpackedDir: dir });
                const result = await v.validateDeletions();
                expect(result.valid).toBe(false);
                expect(result.issues.some((i) => i.code === "del-contains-t")).toBe(true);
            });
        });

        it("passes when <w:del> wraps <w:delText>", async () => {
            await withTempDir(async (dir) => {
                await writeFile(
                    path.join(dir, "word", "document.xml"),
                    wrapDocument(`<w:p><w:del w:id="1"><w:r><w:delText>ok</w:delText></w:r></w:del></w:p>`),
                );
                const v = new DOCXSchemaValidator({ unpackedDir: dir });
                const result = await v.validateDeletions();
                expect(result.valid).toBe(true);
            });
        });
    });

    describe("validateInsertions", () => {
        it("flags <w:delText> inside <w:ins> (no enclosing <w:del>)", async () => {
            await withTempDir(async (dir) => {
                await writeFile(
                    path.join(dir, "word", "document.xml"),
                    wrapDocument(`<w:p><w:ins w:id="2"><w:r><w:delText>nope</w:delText></w:r></w:ins></w:p>`),
                );
                const v = new DOCXSchemaValidator({ unpackedDir: dir });
                const result = await v.validateInsertions();
                expect(result.valid).toBe(false);
                expect(result.issues.some((i) => i.code === "ins-contains-deltext")).toBe(true);
            });
        });
    });

    describe("validateCommentMarkers", () => {
        it("detects orphaned commentRangeEnd", async () => {
            await withTempDir(async (dir) => {
                await writeFile(path.join(dir, "word", "document.xml"), wrapDocument(`<w:p><w:commentRangeEnd w:id="9"/></w:p>`));
                const v = new DOCXSchemaValidator({ unpackedDir: dir });
                const result = await v.validateCommentMarkers();
                expect(result.valid).toBe(false);
                expect(result.issues[0].code).toBe("comment-orphan-end");
                expect(result.issues[0].message).toContain('id="9"');
            });
        });

        it("detects markers referencing a non-existent comment", async () => {
            await withTempDir(async (dir) => {
                await writeFile(
                    path.join(dir, "word", "document.xml"),
                    wrapDocument(
                        `<w:p><w:commentRangeStart w:id="0"/><w:r><w:t>x</w:t></w:r>` +
                            `<w:commentRangeEnd w:id="0"/>` +
                            `<w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr><w:commentReference w:id="0"/></w:r></w:p>`,
                    ),
                );
                await writeFile(
                    path.join(dir, "word", "comments.xml"),
                    `<?xml version="1.0"?><w:comments ${W_NS}>` +
                        `<w:comment w:id="42" w:author="A" w:date="2026-01-01T00:00:00Z" w:initials="A"/>` +
                        `</w:comments>`,
                );
                const v = new DOCXSchemaValidator({ unpackedDir: dir });
                const result = await v.validateCommentMarkers();
                expect(result.valid).toBe(false);
                expect(result.issues.some((i) => i.code === "comment-marker-missing" && i.message.includes('"0"'))).toBe(true);
            });
        });

        it("passes when markers are paired and reference an existing comment", async () => {
            await withTempDir(async (dir) => {
                await writeFile(
                    path.join(dir, "word", "document.xml"),
                    wrapDocument(
                        `<w:p><w:commentRangeStart w:id="0"/><w:r><w:t>x</w:t></w:r>` +
                            `<w:commentRangeEnd w:id="0"/>` +
                            `<w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr><w:commentReference w:id="0"/></w:r></w:p>`,
                    ),
                );
                await writeFile(
                    path.join(dir, "word", "comments.xml"),
                    `<?xml version="1.0"?><w:comments ${W_NS}>` +
                        `<w:comment w:id="0" w:author="A" w:date="2026-01-01T00:00:00Z" w:initials="A"/>` +
                        `</w:comments>`,
                );
                const v = new DOCXSchemaValidator({ unpackedDir: dir });
                const result = await v.validateCommentMarkers();
                expect(result.valid).toBe(true);
            });
        });
    });

    describe("validateIdConstraints", () => {
        it("flags paraId >= 0x80000000", async () => {
            await withTempDir(async (dir) => {
                await writeFile(path.join(dir, "word", "document.xml"), wrapDocument(`<w:p w14:paraId="80000000"/>`, W14_NS));
                const v = new DOCXSchemaValidator({ unpackedDir: dir });
                const result = await v.validateIdConstraints();
                expect(result.valid).toBe(false);
                expect(result.issues[0].code).toBe("id-paraid-overflow");
            });
        });

        it("flags durableId >= 0x7FFFFFFF (hex, non-numbering file)", async () => {
            await withTempDir(async (dir) => {
                await writeFile(
                    path.join(dir, "word", "commentsIds.xml"),
                    `<?xml version="1.0"?><w16cid:commentsIds ${W16CID_NS}>` +
                        `<w16cid:commentId w16cid:paraId="00000001" w16cid:durableId="7FFFFFFF"/>` +
                        `</w16cid:commentsIds>`,
                );
                const v = new DOCXSchemaValidator({ unpackedDir: dir });
                const result = await v.validateIdConstraints();
                expect(result.valid).toBe(false);
                expect(result.issues[0].code).toBe("id-durable-overflow");
            });
        });

        it("flags non-decimal durableId in numbering.xml", async () => {
            await withTempDir(async (dir) => {
                await writeFile(
                    path.join(dir, "word", "numbering.xml"),
                    `<?xml version="1.0"?><w:numbering ${W_NS} ${W16CID_NS}>` +
                        `<w:abstractNum w16cid:durableId="ABCDEF"/>` +
                        `</w:numbering>`,
                );
                const v = new DOCXSchemaValidator({ unpackedDir: dir });
                const result = await v.validateIdConstraints();
                expect(result.valid).toBe(false);
                expect(result.issues[0].code).toBe("id-durable-decimal");
            });
        });

        it("passes when paraId / durableId are within bounds", async () => {
            await withTempDir(async (dir) => {
                await writeFile(path.join(dir, "word", "document.xml"), wrapDocument(`<w:p w14:paraId="11111111"/>`, W14_NS));
                const v = new DOCXSchemaValidator({ unpackedDir: dir });
                const result = await v.validateIdConstraints();
                expect(result.valid).toBe(true);
            });
        });
    });

    describe("repairParaId", () => {
        it("rewrites an over-cap paraId in document.xml", async () => {
            await withTempDir(async (dir) => {
                const filePath = path.join(dir, "word", "document.xml");
                await writeFile(
                    filePath,
                    `<?xml version="1.0"?><w:document ${W_NS} ${W14_NS}><w:body>` +
                        `<w:p w14:paraId="FFFFFFFF"/>` +
                        `</w:body></w:document>`,
                );
                const v = new DOCXSchemaValidator({ unpackedDir: dir });
                const before = await v.validateIdConstraints();
                expect(before.valid).toBe(false);
                expect(before.issues.some((i) => i.code === "id-paraid-overflow")).toBe(true);

                const repairs = await v.repairParaId();
                expect(repairs).toBe(1);

                const after = await v.validateIdConstraints();
                expect(after.valid).toBe(true);

                const xml = await fs.readFile(filePath, "utf-8");
                expect(xml).not.toContain('w14:paraId="FFFFFFFF"');
                expect(xml).toMatch(/w14:paraId="[0-9A-F]{8}"/);
            });
        });

        it("repairs the endnotes.paraid-overflow fixture end-to-end", async () => {
            const fixturePath = path.join(__dirname, "fixtures/broken/endnotes.paraid-overflow.docx");
            await withTempDir(async (dir) => {
                const buf = await fs.readFile(fixturePath);
                const zip = await JSZip.loadAsync(buf);
                for (const [entryName, entry] of Object.entries(zip.files)) {
                    if (entry.dir) continue;
                    const dest = path.join(dir, entryName);
                    await fs.mkdir(path.dirname(dest), { recursive: true });
                    await fs.writeFile(dest, await entry.async("nodebuffer"));
                }

                const v = new DOCXSchemaValidator({ unpackedDir: dir });
                const before = await v.validateIdConstraints();
                expect(before.valid).toBe(false);
                expect(before.issues.some((i) => i.code === "id-paraid-overflow")).toBe(true);

                const repairs = await v.repairParaId();
                expect(repairs).toBeGreaterThan(0);

                const after = await v.validateIdConstraints();
                expect(after.valid).toBe(true);
                expect(after.issues.filter((i) => i.code === "id-paraid-overflow")).toHaveLength(0);
            });
        });
    });

    describe("repairDurableId", () => {
        it("rewrites an over-cap hex durableId in commentsIds.xml", async () => {
            await withTempDir(async (dir) => {
                const filePath = path.join(dir, "word", "commentsIds.xml");
                await writeFile(
                    filePath,
                    `<?xml version="1.0"?><w16cid:commentsIds ${W16CID_NS}>` +
                        `<w16cid:commentId w16cid:paraId="00000001" w16cid:durableId="FFFFFFFF"/>` +
                        `</w16cid:commentsIds>`,
                );
                const v = new DOCXSchemaValidator({ unpackedDir: dir });
                const repairs = await v.repairDurableId();
                expect(repairs).toBe(1);
                const after = await fs.readFile(filePath, "utf-8");
                expect(after).not.toContain('w16cid:durableId="FFFFFFFF"');
                expect(after).toMatch(/w16cid:durableId="[0-9A-F]{8}"/);
            });
        });
    });

    describe("repairParaId — cross-file consistency (regression: #commentsExtended)", () => {
        const W15_NS = `xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml"`;

        it("renumbers an over-cap paraId in document.xml AND comments.xml AND commentsExtended.xml to the same value", async () => {
            await withTempDir(async (dir) => {
                // The same over-cap paraId B68569E0 appears in all three
                // files. After repair, all three must agree on the new
                // (in-range) value, otherwise threading breaks.
                const docPath = path.join(dir, "word", "document.xml");
                const commentsPath = path.join(dir, "word", "comments.xml");
                const extPath = path.join(dir, "word", "commentsExtended.xml");
                await writeFile(
                    docPath,
                    `<?xml version="1.0"?><w:document ${W_NS} ${W14_NS}><w:body>` +
                        `<w:p w14:paraId="B68569E0"/>` +
                        `</w:body></w:document>`,
                );
                await writeFile(
                    commentsPath,
                    `<?xml version="1.0"?><w:comments ${W_NS} ${W14_NS}>` +
                        `<w:comment w:id="0" w:author="A" w:date="2026-01-01T00:00:00Z" w:initials="A">` +
                        `<w:p w14:paraId="B68569E0"/></w:comment></w:comments>`,
                );
                await writeFile(
                    extPath,
                    `<?xml version="1.0"?><w15:commentsEx ${W15_NS}>` +
                        `<w15:commentEx w15:paraId="B68569E0" w15:done="0"/>` +
                        `</w15:commentsEx>`,
                );

                const v = new DOCXSchemaValidator({ unpackedDir: dir });
                const repairs = await v.repairParaId();

                // Three rewrites: document.xml/w14:paraId, comments.xml/w14:paraId, commentsExtended.xml/w15:paraId.
                expect(repairs).toBe(3);

                const docXml = await fs.readFile(docPath, "utf-8");
                const commentsXml = await fs.readFile(commentsPath, "utf-8");
                const extXml = await fs.readFile(extPath, "utf-8");
                expect(docXml).not.toContain('paraId="B68569E0"');
                expect(commentsXml).not.toContain('paraId="B68569E0"');
                expect(extXml).not.toContain('paraId="B68569E0"');

                // All three files must end up with the SAME new paraId so
                // commentRangeStart/Reference still resolve to a real
                // <w:comment> and threading still works.
                const docMatch = /w14:paraId="([0-9A-F]{8})"/.exec(docXml);
                const commentsMatch = /w14:paraId="([0-9A-F]{8})"/.exec(commentsXml);
                const extMatch = /w15:paraId="([0-9A-F]{8})"/.exec(extXml);
                expect(docMatch?.[1]).toBeDefined();
                expect(commentsMatch?.[1]).toBe(docMatch?.[1]);
                expect(extMatch?.[1]).toBe(docMatch?.[1]);
            });
        });

        it("renumbers w15:paraIdParent in commentsExtended.xml when the parent paraId is also over-cap", async () => {
            await withTempDir(async (dir) => {
                const docPath = path.join(dir, "word", "document.xml");
                const commentsPath = path.join(dir, "word", "comments.xml");
                const extPath = path.join(dir, "word", "commentsExtended.xml");
                await writeFile(
                    docPath,
                    `<?xml version="1.0"?><w:document ${W_NS} ${W14_NS}><w:body>` +
                        `<w:p w14:paraId="AAAA0000"/>` +
                        `</w:body></w:document>`,
                );
                // Parent is over-cap (E0000000), reply is in-range (BBBBBBBB).
                await writeFile(
                    commentsPath,
                    `<?xml version="1.0"?><w:comments ${W_NS} ${W14_NS}>` +
                        `<w:comment w:id="0" w:author="A" w:date="2026-01-01T00:00:00Z" w:initials="A">` +
                        `<w:p w14:paraId="E0000000"/></w:comment>` +
                        `<w:comment w:id="1" w:author="B" w:date="2026-01-01T00:00:00Z" w:initials="B">` +
                        `<w:p w14:paraId="BBBBBBBB"/></w:comment></w:comments>`,
                );
                await writeFile(
                    extPath,
                    `<?xml version="1.0"?><w15:commentsEx ${W15_NS}>` +
                        `<w15:commentEx w15:paraId="E0000000" w15:done="0"/>` +
                        `<w15:commentEx w15:paraId="BBBBBBBB" w15:paraIdParent="E0000000" w15:done="0"/>` +
                        `</w15:commentsEx>`,
                );

                const v = new DOCXSchemaValidator({ unpackedDir: dir });
                await v.repairParaId();

                const commentsXml = await fs.readFile(commentsPath, "utf-8");
                const extXml = await fs.readFile(extPath, "utf-8");

                // The new in-range value the parent got rewritten to.
                const newParent = /w14:paraId="([0-9A-F]{8})"\s*[^>]*\/>\s*<\/w:comment>\s*<w:comment[^>]*>\s*<w:p w14:paraId="BBBBBBBB"/.exec(commentsXml)?.[1] ??
                    /<w:comment w:id="0"[^>]*>\s*<w:p w14:paraId="([0-9A-F]{8})"/.exec(commentsXml)?.[1];
                expect(newParent).toBeDefined();
                if (!newParent) return;
                expect(parseInt(newParent, 16)).toBeLessThan(0x80000000);

                // Both the paraId entry AND the paraIdParent reference in
                // commentsExtended.xml must point at the same new value.
                expect(extXml).toContain(`w15:paraId="${newParent}"`);
                expect(extXml).toContain(`w15:paraIdParent="${newParent}"`);
                expect(extXml).not.toContain("E0000000");
            });
        });

        it("repairs the sample-document.id-overflow fixture cleanly (no orphan paraIds afterward)", async () => {
            const fixturePath = path.join(__dirname, "fixtures/broken/sample-document.id-overflow.docx");
            await withTempDir(async (dir) => {
                const buf = await fs.readFile(fixturePath);
                const zip = await JSZip.loadAsync(buf);
                for (const [entryName, entry] of Object.entries(zip.files)) {
                    if (entry.dir) continue;
                    const dest = path.join(dir, entryName);
                    await fs.mkdir(path.dirname(dest), { recursive: true });
                    await fs.writeFile(dest, await entry.async("nodebuffer"));
                }

                const v = new DOCXSchemaValidator({ unpackedDir: dir });
                const repairs = await v.repairParaId();
                expect(repairs).toBeGreaterThan(0);

                const idCheck = await v.validateIdConstraints();
                expect(idCheck.issues.filter((i) => i.code === "id-paraid-overflow")).toHaveLength(0);

                // The threading cross-reference must survive the rename:
                // commentsExtended.xml's paraIds must still resolve to a
                // real <w:comment> first paragraph.
                const threadCheck = await v.validateCommentThreading();
                expect(threadCheck.issues.filter((i) => i.code === "comment-thread-paraid-orphan")).toHaveLength(0);
            });
        });
    });

    describe("validateAllParagraphsHaveParaId / repairMissingParaIds", () => {
        const TBL_NS = `xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"`;

        it("flags missing paraId on <w:tr> as ERROR under strict, WARNING under lenient", async () => {
            await withTempDir(async (dir) => {
                await writeFile(
                    path.join(dir, "word", "document.xml"),
                    `<?xml version="1.0"?><w:document ${TBL_NS} ${W14_NS}><w:body>` +
                        `<w:tbl><w:tr><w:tc><w:p w14:paraId="11111111"/></w:tc></w:tr></w:tbl>` +
                        `</w:body></w:document>`,
                );
                const strict = new DOCXSchemaValidator({ unpackedDir: dir, profile: "strict" });
                const r1 = await strict.validateAllParagraphsHaveParaId();
                expect(r1.valid).toBe(false);
                expect(r1.issues.find((i) => i.code === "paraid-missing-element" && i.message.includes("<w:tr>"))).toBeDefined();
                expect(r1.issues.find((i) => i.code === "paraid-missing-element")?.severity).toBe("error");

                const lenient = new DOCXSchemaValidator({ unpackedDir: dir, profile: "lenient" });
                const r2 = await lenient.validateAllParagraphsHaveParaId();
                expect(r2.valid).toBe(true);
                expect(r2.issues.find((i) => i.code === "paraid-missing-element")?.severity).toBe("warning");
            });
        });

        it("repairMissingParaIds stamps a w14:paraId on every <w:p> and <w:tr> that lacks one", async () => {
            await withTempDir(async (dir) => {
                const filePath = path.join(dir, "word", "document.xml");
                await writeFile(
                    filePath,
                    `<?xml version="1.0"?><w:document ${TBL_NS} ${W14_NS}><w:body>` +
                        // First paragraph already has a paraId — should be left alone.
                        `<w:p w14:paraId="11111111"/>` +
                        // Second paragraph has none — should get one.
                        `<w:p/>` +
                        // Table with two rows, neither stamped.
                        `<w:tbl><w:tr><w:tc><w:p/></w:tc></w:tr><w:tr><w:tc><w:p/></w:tc></w:tr></w:tbl>` +
                        `</w:body></w:document>`,
                );
                const v = new DOCXSchemaValidator({ unpackedDir: dir });
                const repairs = await v.repairMissingParaIds();
                // 1 missing paragraph + 2 missing rows + 2 missing inner-cell paragraphs = 5
                expect(repairs).toBe(5);

                const after = await fs.readFile(filePath, "utf-8");
                // Original paraId preserved.
                expect(after).toContain('w14:paraId="11111111"');
                // Every <w:p> and <w:tr> now has a paraId.
                const v2 = new DOCXSchemaValidator({ unpackedDir: dir, profile: "strict" });
                const post = await v2.validateAllParagraphsHaveParaId();
                expect(post.valid).toBe(true);
                expect(post.issues).toEqual([]);
            });
        });

        it("repairMissingParaIds picks values that don't collide with existing paraIds", async () => {
            await withTempDir(async (dir) => {
                const filePath = path.join(dir, "word", "document.xml");
                await writeFile(
                    filePath,
                    `<?xml version="1.0"?><w:document ${TBL_NS} ${W14_NS}><w:body>` +
                        `<w:p w14:paraId="11111111"/>` +
                        `<w:p w14:paraId="22222222"/>` +
                        `<w:p/>` +
                        `</w:body></w:document>`,
                );
                const v = new DOCXSchemaValidator({ unpackedDir: dir });
                await v.repairMissingParaIds();
                const after = await fs.readFile(filePath, "utf-8");
                const paraIds = Array.from(after.matchAll(/w14:paraId="([0-9A-F]{8})"/g)).map((m) => m[1]);
                expect(paraIds).toHaveLength(3);
                // All three must be unique.
                expect(new Set(paraIds).size).toBe(3);
                // The pre-existing ones survived.
                expect(paraIds).toContain("11111111");
                expect(paraIds).toContain("22222222");
            });
        });
    });

    describe("repairIgnorable", () => {
        it("declares known OOXML prefixes (w15, wp14, …) on the document root rather than dropping them from mc:Ignorable", async () => {
            await withTempDir(async (dir) => {
                const filePath = path.join(dir, "word", "document.xml");
                const mcNs = `xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"`;
                await writeFile(
                    filePath,
                    `<?xml version="1.0"?><w:document ${W_NS} ${mcNs} mc:Ignorable="w15 wp14"><w:body/></w:document>`,
                );
                const v = new DOCXSchemaValidator({ unpackedDir: dir });
                const repairs = await v.repairIgnorable();
                expect(repairs).toBe(2);

                const after = await fs.readFile(filePath, "utf-8");
                // Both prefixes were declared with their canonical URIs.
                expect(after).toContain('xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml"');
                expect(after).toContain('xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing"');
                // The mc:Ignorable list survives intact.
                expect(after).toMatch(/mc:Ignorable="(w15 wp14|wp14 w15)"/);

                // And `validateNamespaces` no longer flags ignorable-undeclared.
                const after2 = await v.validateNamespaces();
                expect(after2.issues.filter((i) => i.code === "ignorable-undeclared")).toHaveLength(0);
            });
        });

        it("drops a truly-unknown prefix from mc:Ignorable (not in the well-known table)", async () => {
            await withTempDir(async (dir) => {
                const filePath = path.join(dir, "word", "document.xml");
                const mcNs = `xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"`;
                await writeFile(
                    filePath,
                    `<?xml version="1.0"?><w:document ${W_NS} ${mcNs} mc:Ignorable="weirdthirdparty"><w:body/></w:document>`,
                );
                const v = new DOCXSchemaValidator({ unpackedDir: dir });
                const repairs = await v.repairIgnorable();
                expect(repairs).toBe(1);

                const after = await fs.readFile(filePath, "utf-8");
                // Nothing was declared — "weirdthirdparty" isn't a known
                // prefix — and the Ignorable attribute was removed entirely
                // because it had no surviving tokens.
                expect(after).not.toContain("xmlns:weirdthirdparty");
                expect(after).not.toContain("mc:Ignorable=");
            });
        });

        it("leaves an already-correct mc:Ignorable untouched", async () => {
            await withTempDir(async (dir) => {
                const filePath = path.join(dir, "word", "document.xml");
                const w15Ns = `xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml"`;
                const mcNs = `xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"`;
                const before =
                    `<?xml version="1.0"?><w:document ${W_NS} ${w15Ns} ${mcNs} mc:Ignorable="w15"><w:body/></w:document>`;
                await writeFile(filePath, before);
                const v = new DOCXSchemaValidator({ unpackedDir: dir });
                const repairs = await v.repairIgnorable();
                expect(repairs).toBe(0);
                expect(await fs.readFile(filePath, "utf-8")).toBe(before);
            });
        });
    });

    describe("WORD_PARAGRAPH_NAMESPACES", () => {
        it("exports the two expected namespace URIs", () => {
            expect(WORD_PARAGRAPH_NAMESPACES).toHaveLength(2);
            expect(WORD_PARAGRAPH_NAMESPACES[0]).toBe(WORD_2006_NAMESPACE);
            expect(WORD_PARAGRAPH_NAMESPACES[1]).toBe(WORD_STRICT_NAMESPACE);
            expect(WORD_2006_NAMESPACE).toBe("http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            expect(WORD_STRICT_NAMESPACE).toBe("http://purl.oclc.org/ooxml/wordprocessingml/main");
        });
    });

    describe("paragraph counts", () => {
        it("counts <w:p> in unpacked document, excluding text-box overlays", async () => {
            await withTempDir(async (dir) => {
                const vmlNs = `xmlns:v="urn:schemas-microsoft-com:vml"`;
                await writeFile(
                    path.join(dir, "word", "document.xml"),
                    `<?xml version="1.0"?><w:document ${W_NS} ${vmlNs}><w:body>` +
                        `<w:p><w:r><w:t>one</w:t></w:r></w:p>` +
                        `<w:p><w:r><w:t>two</w:t></w:r></w:p>` +
                        `<w:p><w:pict><v:shape><v:textbox><w:txbxContent><w:p><w:r><w:t>boxed</w:t></w:r></w:p></w:txbxContent></v:textbox></v:shape></w:pict></w:p>` +
                        `</w:body></w:document>`,
                );
                const v = new DOCXSchemaValidator({ unpackedDir: dir });
                // Three w:p elements at body, one nested inside v:textbox/w:txbxContent.
                // The Python xpath excludes the nested one.
                expect(v.countParagraphsInUnpacked()).toBe(3);
            });
        });

        it("compareParagraphCounts handles missing original gracefully", async () => {
            await withTempDir(async (dir) => {
                await writeFile(path.join(dir, "word", "document.xml"), wrapDocument(`<w:p/>`));
                const v = new DOCXSchemaValidator({ unpackedDir: dir });
                const counts = await v.compareParagraphCounts();
                expect(counts.original).toBe(0);
                expect(counts.modified).toBe(1);
                expect(counts.delta).toBe(1);
                expect(counts.originalUsesStrictNamespace).toBe(false);
            });
        });

        it("loads paragraph count from a real .docx zip when originalFile is set", async () => {
            await withTempDir(async (dir) => {
                const docXml = wrapDocument(`<w:p/><w:p/><w:p/>`);
                const zip = new JSZip();
                zip.file("word/document.xml", docXml);
                const buf = await zip.generateAsync({ type: "nodebuffer" });
                const origPath = path.join(dir, "orig.docx");
                await fs.writeFile(origPath, buf);

                await writeFile(path.join(dir, "unpacked", "word", "document.xml"), wrapDocument(`<w:p/>`));
                const v = new DOCXSchemaValidator({
                    unpackedDir: path.join(dir, "unpacked"),
                    originalFile: origPath,
                });
                const counts = await v.compareParagraphCounts();
                expect(counts.original).toBe(3);
                expect(counts.modified).toBe(1);
                expect(counts.delta).toBe(-2);
            });
        });

        it("compareParagraphCounts prints a summary line when verbose=true (Python parity)", async () => {
            const lines: string[] = [];
            const origWrite = process.stdout.write.bind(process.stdout);
            process.stdout.write = (chunk: unknown) => {
                if (typeof chunk === "string") lines.push(chunk);
                return true;
            };
            try {
                await withTempDir(async (dir) => {
                    await writeFile(path.join(dir, "word", "document.xml"), wrapDocument(`<w:p/><w:p/>`));
                    const v = new DOCXSchemaValidator({
                        unpackedDir: dir,
                        verbose: true,
                    });
                    await v.compareParagraphCounts();
                });
            } finally {
                process.stdout.write = origWrite;
            }
            const combined = lines.join("");
            expect(combined).toMatch(/Paragraphs: \d+ → \d+ \([+-]?\d+\)/);
        });
    });

    describe("validateNoTrackingTokens", () => {
        it("flags a leaked [[DOCX_INS_START:...]] token", async () => {
            await withTempDir(async (dir) => {
                await writeFile(
                    path.join(dir, "word", "document.xml"),
                    wrapDocument(`<w:p><w:r><w:t>[[DOCX_INS_START:%7B%22id%22%3A%22a%22%7D]]hi</w:t></w:r></w:p>`),
                );
                const v = new DOCXSchemaValidator({ unpackedDir: dir });
                const result = await v.validateNoTrackingTokens();
                expect(result.valid).toBe(false);
                expect(result.issues[0].code).toBe("tracking-token-leak");
                expect(result.issues[0].message).toContain("DOCX_INS_START");
            });
        });

        it("flags every distinct token type once each", async () => {
            await withTempDir(async (dir) => {
                const tokens =
                    "[[DOCX_INS_START:foo]]a[[DOCX_INS_END:foo]]" +
                    "[[DOCX_DEL_START:bar]]b[[DOCX_DEL_END:bar]]" +
                    "[[DOCX_CMT_START:baz]]c[[DOCX_CMT_END:baz]]" +
                    "[[DOCX_PMARK_INS:p1]][[DOCX_PMARK_DEL:p2]]";
                await writeFile(
                    path.join(dir, "word", "document.xml"),
                    wrapDocument(`<w:p><w:r><w:t>${tokens}</w:t></w:r></w:p>`),
                );
                const v = new DOCXSchemaValidator({ unpackedDir: dir });
                const result = await v.validateNoTrackingTokens();
                expect(result.valid).toBe(false);
                // 8 distinct tokens.
                expect(result.issues).toHaveLength(8);
                expect(result.issues.every((i) => i.code === "tracking-token-leak")).toBe(true);
            });
        });

        it("passes when no tracking tokens are present", async () => {
            await withTempDir(async (dir) => {
                await writeFile(
                    path.join(dir, "word", "document.xml"),
                    wrapDocument(`<w:p><w:r><w:t>regular [[brackets]] text</w:t></w:r></w:p>`),
                );
                const v = new DOCXSchemaValidator({ unpackedDir: dir });
                const result = await v.validateNoTrackingTokens();
                expect(result.valid).toBe(true);
                expect(result.issues).toHaveLength(0);
            });
        });

        it("scans header/footer/footnote XML, not just document.xml", async () => {
            await withTempDir(async (dir) => {
                await writeFile(path.join(dir, "word", "document.xml"), wrapDocument(`<w:p/>`));
                await writeFile(
                    path.join(dir, "word", "header1.xml"),
                    `<?xml version="1.0"?><w:hdr ${W_NS}><w:p><w:r><w:t>[[DOCX_CMT_START:x]]hi[[DOCX_CMT_END:x]]</w:t></w:r></w:p></w:hdr>`,
                );
                const v = new DOCXSchemaValidator({ unpackedDir: dir });
                const result = await v.validateNoTrackingTokens();
                expect(result.valid).toBe(false);
                expect(result.issues.some((i) => i.path?.endsWith("header1.xml"))).toBe(true);
            });
        });
    });

    describe("validateCommentThreading", () => {
        const W15_NS = `xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml"`;

        it("passes when comments.xml is absent", async () => {
            await withTempDir(async (dir) => {
                await writeFile(path.join(dir, "word", "document.xml"), wrapDocument(`<w:p/>`));
                const v = new DOCXSchemaValidator({ unpackedDir: dir });
                const result = await v.validateCommentThreading();
                expect(result.valid).toBe(true);
            });
        });

        it("passes when comments.xml exists but commentsExtended.xml is absent and counts match", async () => {
            await withTempDir(async (dir) => {
                await writeFile(
                    path.join(dir, "word", "document.xml"),
                    wrapDocument(
                        `<w:p><w:commentRangeStart w:id="0"/>x<w:commentRangeEnd w:id="0"/>` +
                            `<w:r><w:commentReference w:id="0"/></w:r></w:p>`,
                    ),
                );
                await writeFile(
                    path.join(dir, "word", "comments.xml"),
                    `<?xml version="1.0"?><w:comments ${W_NS} ${W14_NS}>` +
                        `<w:comment w:id="0" w:author="A" w:date="2026-01-01T00:00:00Z" w:initials="A">` +
                        `<w:p w14:paraId="11111111"/></w:comment></w:comments>`,
                );
                const v = new DOCXSchemaValidator({ unpackedDir: dir });
                const result = await v.validateCommentThreading();
                expect(result.valid).toBe(true);
            });
        });

        it("flags <w15:commentEx> with a paraId that no <w:comment> has (orphan extension entry)", async () => {
            await withTempDir(async (dir) => {
                await writeFile(
                    path.join(dir, "word", "document.xml"),
                    wrapDocument(
                        `<w:p><w:commentRangeStart w:id="0"/>x<w:commentRangeEnd w:id="0"/>` +
                            `<w:r><w:commentReference w:id="0"/></w:r></w:p>`,
                    ),
                );
                await writeFile(
                    path.join(dir, "word", "comments.xml"),
                    `<?xml version="1.0"?><w:comments ${W_NS} ${W14_NS}>` +
                        `<w:comment w:id="0" w:author="A" w:date="2026-01-01T00:00:00Z" w:initials="A">` +
                        `<w:p w14:paraId="11111111"/></w:comment></w:comments>`,
                );
                await writeFile(
                    path.join(dir, "word", "commentsExtended.xml"),
                    `<?xml version="1.0"?><w15:commentsEx ${W15_NS}>` +
                        `<w15:commentEx w15:paraId="11111111" w15:done="0"/>` +
                        `<w15:commentEx w15:paraId="DEADBEEF" w15:done="0"/>` +
                        `</w15:commentsEx>`,
                );
                const v = new DOCXSchemaValidator({ unpackedDir: dir });
                const result = await v.validateCommentThreading();
                expect(result.valid).toBe(false);
                expect(
                    result.issues.some((i) => i.code === "comment-thread-paraid-orphan" && i.message.includes("DEADBEEF")),
                ).toBe(true);
            });
        });

        it("flags missing <w15:commentEx> as ERROR in strict profile", async () => {
            await withTempDir(async (dir) => {
                await writeFile(
                    path.join(dir, "word", "document.xml"),
                    wrapDocument(
                        `<w:p>` +
                            `<w:commentRangeStart w:id="0"/><w:commentRangeStart w:id="1"/>x` +
                            `<w:commentRangeEnd w:id="0"/><w:commentRangeEnd w:id="1"/>` +
                            `<w:r><w:commentReference w:id="0"/></w:r>` +
                            `<w:r><w:commentReference w:id="1"/></w:r>` +
                            `</w:p>`,
                    ),
                );
                await writeFile(
                    path.join(dir, "word", "comments.xml"),
                    `<?xml version="1.0"?><w:comments ${W_NS} ${W14_NS}>` +
                        `<w:comment w:id="0" w:author="A" w:date="2026-01-01T00:00:00Z" w:initials="A">` +
                        `<w:p w14:paraId="11111111"/></w:comment>` +
                        `<w:comment w:id="1" w:author="B" w:date="2026-01-01T00:00:00Z" w:initials="B">` +
                        `<w:p w14:paraId="22222222"/></w:comment></w:comments>`,
                );
                await writeFile(
                    path.join(dir, "word", "commentsExtended.xml"),
                    `<?xml version="1.0"?><w15:commentsEx ${W15_NS}>` +
                        `<w15:commentEx w15:paraId="11111111" w15:done="0"/>` +
                        `</w15:commentsEx>`,
                );
                const v = new DOCXSchemaValidator({ unpackedDir: dir, profile: "strict" });
                const result = await v.validateCommentThreading();
                expect(result.valid).toBe(false);
                const missing = result.issues.find(
                    (i) => i.code === "comment-thread-paraid-missing" && i.message.includes("22222222"),
                );
                expect(missing).toBeDefined();
                expect(missing?.severity).toBe("error");
            });
        });

        it("downgrades missing <w15:commentEx> to WARNING in lenient profile (Word may legitimately omit it)", async () => {
            await withTempDir(async (dir) => {
                await writeFile(
                    path.join(dir, "word", "document.xml"),
                    wrapDocument(
                        `<w:p>` +
                            `<w:commentRangeStart w:id="0"/><w:commentRangeStart w:id="1"/>x` +
                            `<w:commentRangeEnd w:id="0"/><w:commentRangeEnd w:id="1"/>` +
                            `<w:r><w:commentReference w:id="0"/></w:r>` +
                            `<w:r><w:commentReference w:id="1"/></w:r>` +
                            `</w:p>`,
                    ),
                );
                await writeFile(
                    path.join(dir, "word", "comments.xml"),
                    `<?xml version="1.0"?><w:comments ${W_NS} ${W14_NS}>` +
                        `<w:comment w:id="0" w:author="A" w:date="2026-01-01T00:00:00Z" w:initials="A">` +
                        `<w:p w14:paraId="11111111"/></w:comment>` +
                        `<w:comment w:id="1" w:author="B" w:date="2026-01-01T00:00:00Z" w:initials="B">` +
                        `<w:p w14:paraId="22222222"/></w:comment></w:comments>`,
                );
                await writeFile(
                    path.join(dir, "word", "commentsExtended.xml"),
                    `<?xml version="1.0"?><w15:commentsEx ${W15_NS}>` +
                        `<w15:commentEx w15:paraId="11111111" w15:done="0"/>` +
                        `</w15:commentsEx>`,
                );
                const v = new DOCXSchemaValidator({ unpackedDir: dir, profile: "lenient" });
                const result = await v.validateCommentThreading();
                // Lenient profile: still reports the issue, but as a warning,
                // so the document remains "valid" overall.
                expect(result.valid).toBe(true);
                const missing = result.issues.find(
                    (i) => i.code === "comment-thread-paraid-missing" && i.message.includes("22222222"),
                );
                expect(missing).toBeDefined();
                expect(missing?.severity).toBe("warning");
            });
        });

        it("flags duplicate paraId entries in commentsExtended.xml (regression test for #153)", async () => {
            await withTempDir(async (dir) => {
                await writeFile(
                    path.join(dir, "word", "document.xml"),
                    wrapDocument(
                        `<w:p><w:commentRangeStart w:id="0"/>x<w:commentRangeEnd w:id="0"/>` +
                            `<w:r><w:commentReference w:id="0"/></w:r></w:p>`,
                    ),
                );
                await writeFile(
                    path.join(dir, "word", "comments.xml"),
                    `<?xml version="1.0"?><w:comments ${W_NS} ${W14_NS}>` +
                        `<w:comment w:id="0" w:author="A" w:date="2026-01-01T00:00:00Z" w:initials="A">` +
                        `<w:p w14:paraId="11111111"/></w:comment></w:comments>`,
                );
                await writeFile(
                    path.join(dir, "word", "commentsExtended.xml"),
                    `<?xml version="1.0"?><w15:commentsEx ${W15_NS}>` +
                        `<w15:commentEx w15:paraId="11111111" w15:done="0"/>` +
                        `<w15:commentEx w15:paraId="11111111" w15:done="0"/>` +
                        `</w15:commentsEx>`,
                );
                const v = new DOCXSchemaValidator({ unpackedDir: dir });
                const result = await v.validateCommentThreading();
                expect(result.valid).toBe(false);
                expect(result.issues.some((i) => i.code === "comment-thread-duplicate-paraid")).toBe(true);
            });
        });

        it("flags <w15:paraIdParent> that does not resolve", async () => {
            await withTempDir(async (dir) => {
                await writeFile(
                    path.join(dir, "word", "document.xml"),
                    wrapDocument(
                        `<w:p><w:commentRangeStart w:id="0"/>x<w:commentRangeEnd w:id="0"/>` +
                            `<w:r><w:commentReference w:id="0"/></w:r></w:p>`,
                    ),
                );
                await writeFile(
                    path.join(dir, "word", "comments.xml"),
                    `<?xml version="1.0"?><w:comments ${W_NS} ${W14_NS}>` +
                        `<w:comment w:id="0" w:author="A" w:date="2026-01-01T00:00:00Z" w:initials="A">` +
                        `<w:p w14:paraId="11111111"/></w:comment></w:comments>`,
                );
                await writeFile(
                    path.join(dir, "word", "commentsExtended.xml"),
                    `<?xml version="1.0"?><w15:commentsEx ${W15_NS}>` +
                        `<w15:commentEx w15:paraId="11111111" w15:paraIdParent="DEADBEEF" w15:done="0"/>` +
                        `</w15:commentsEx>`,
                );
                const v = new DOCXSchemaValidator({ unpackedDir: dir });
                const result = await v.validateCommentThreading();
                expect(result.valid).toBe(false);
                expect(result.issues.some((i) => i.code === "comment-thread-orphan-parent")).toBe(true);
            });
        });

        it("flags commentRangeStart count mismatch", async () => {
            await withTempDir(async (dir) => {
                await writeFile(
                    path.join(dir, "word", "document.xml"),
                    wrapDocument(
                        `<w:p><w:commentRangeStart w:id="0"/>x<w:commentRangeEnd w:id="0"/>` +
                            `<w:r><w:commentReference w:id="0"/></w:r></w:p>`,
                    ),
                );
                await writeFile(
                    path.join(dir, "word", "comments.xml"),
                    `<?xml version="1.0"?><w:comments ${W_NS} ${W14_NS}>` +
                        `<w:comment w:id="0" w:author="A" w:date="2026-01-01T00:00:00Z" w:initials="A">` +
                        `<w:p w14:paraId="11111111"/></w:comment>` +
                        `<w:comment w:id="1" w:author="B" w:date="2026-01-01T00:00:00Z" w:initials="B">` +
                        `<w:p w14:paraId="22222222"/></w:comment></w:comments>`,
                );
                const v = new DOCXSchemaValidator({ unpackedDir: dir });
                const result = await v.validateCommentThreading();
                expect(result.valid).toBe(false);
                expect(
                    result.issues.some(
                        (i) => i.code === "comment-thread-count-mismatch" && i.message.includes("commentRangeStart"),
                    ),
                ).toBe(true);
            });
        });

        it("passes for a well-formed threaded comments fixture", async () => {
            await withTempDir(async (dir) => {
                await writeFile(
                    path.join(dir, "word", "document.xml"),
                    wrapDocument(
                        `<w:p>` +
                            `<w:commentRangeStart w:id="0"/><w:commentRangeStart w:id="1"/>x` +
                            `<w:commentRangeEnd w:id="0"/><w:commentRangeEnd w:id="1"/>` +
                            `<w:r><w:commentReference w:id="0"/></w:r>` +
                            `<w:r><w:commentReference w:id="1"/></w:r>` +
                            `</w:p>`,
                    ),
                );
                await writeFile(
                    path.join(dir, "word", "comments.xml"),
                    `<?xml version="1.0"?><w:comments ${W_NS} ${W14_NS}>` +
                        `<w:comment w:id="0" w:author="A" w:date="2026-01-01T00:00:00Z" w:initials="A">` +
                        `<w:p w14:paraId="AAAAAAAA"/></w:comment>` +
                        `<w:comment w:id="1" w:author="B" w:date="2026-01-01T00:00:00Z" w:initials="B">` +
                        `<w:p w14:paraId="BBBBBBBB"/></w:comment></w:comments>`,
                );
                await writeFile(
                    path.join(dir, "word", "commentsExtended.xml"),
                    `<?xml version="1.0"?><w15:commentsEx ${W15_NS}>` +
                        `<w15:commentEx w15:paraId="AAAAAAAA" w15:done="0"/>` +
                        `<w15:commentEx w15:paraId="BBBBBBBB" w15:paraIdParent="AAAAAAAA" w15:done="0"/>` +
                        `</w15:commentsEx>`,
                );
                const v = new DOCXSchemaValidator({ unpackedDir: dir });
                const result = await v.validateCommentThreading();
                expect(result.valid).toBe(true);
                expect(result.issues).toHaveLength(0);
            });
        });
    });
});
