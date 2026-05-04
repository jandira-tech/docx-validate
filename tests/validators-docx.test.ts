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
});
