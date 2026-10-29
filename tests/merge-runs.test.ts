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
import { mergeRuns } from "../src/scripts/office/helpers/merge-runs";

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';

async function writeDoc(dir: string, body: string): Promise<string> {
    const wordDir = path.join(dir, "word");
    await fs.mkdir(wordDir, { recursive: true });
    const docPath = path.join(wordDir, "document.xml");
    const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:document ${W}>${body}</w:document>`;
    await fs.writeFile(docPath, xml, "utf8");
    return docPath;
}

describe("mergeRuns", () => {
    it("returns an error when document.xml is missing", async () => {
        await withTempDir(async (dir) => {
            const result = await mergeRuns(dir);
            expect(result.count).toBe(0);
            expect(result.message).toMatch(/Error:.*document\.xml not found/);
        });
    });

    it("merges two adjacent runs that share identical empty rPr", async () => {
        await withTempDir(async (dir) => {
            const body = "<w:body><w:p>" + "<w:r><w:t>Hello </w:t></w:r>" + "<w:r><w:t>world</w:t></w:r>" + "</w:p></w:body>";
            const docPath = await writeDoc(dir, body);

            const result = await mergeRuns(dir);
            expect(result.count).toBe(1);
            expect(result.message).toBe("Merged 1 runs");

            const after = await fs.readFile(docPath, "utf8");
            const runMatches = after.match(/<w:r[ >]/g) ?? [];
            expect(runMatches.length).toBe(1);
            expect(after).toMatch(/Hello world/);
        });
    });

    it("preserves leading/trailing whitespace via xml:space when merging", async () => {
        await withTempDir(async (dir) => {
            const body =
                "<w:body><w:p>" +
                '<w:r><w:t xml:space="preserve">trailing </w:t></w:r>' +
                '<w:r><w:t xml:space="preserve">space </w:t></w:r>' +
                "</w:p></w:body>";
            const docPath = await writeDoc(dir, body);

            const result = await mergeRuns(dir);
            expect(result.count).toBe(1);

            const after = await fs.readFile(docPath, "utf8");
            expect(after).toMatch(/trailing space /);
            expect(after).toMatch(/xml:space="preserve"/);
        });
    });

    it("merges adjacent runs that share identical rPr", async () => {
        await withTempDir(async (dir) => {
            const body =
                "<w:body><w:p>" +
                '<w:r w:rsidR="ABC"><w:rPr><w:b/></w:rPr><w:t>foo</w:t></w:r>' +
                '<w:r w:rsidR="DEF"><w:rPr><w:b/></w:rPr><w:t>bar</w:t></w:r>' +
                "</w:p></w:body>";
            const docPath = await writeDoc(dir, body);

            const result = await mergeRuns(dir);
            expect(result.count).toBe(1);

            const after = await fs.readFile(docPath, "utf8");
            const runMatches = after.match(/<w:r[ >]/g) ?? [];
            expect(runMatches.length).toBe(1);
            expect(after).toMatch(/foobar/);
            expect(after).not.toMatch(/w:rsidR/);
        });
    });

    it("does not merge runs with differing rPr", async () => {
        await withTempDir(async (dir) => {
            const body =
                "<w:body><w:p>" +
                "<w:r><w:rPr><w:b/></w:rPr><w:t>bold</w:t></w:r>" +
                "<w:r><w:rPr><w:i/></w:rPr><w:t>italic</w:t></w:r>" +
                "</w:p></w:body>";
            const docPath = await writeDoc(dir, body);

            const result = await mergeRuns(dir);
            expect(result.count).toBe(0);

            const after = await fs.readFile(docPath, "utf8");
            const runMatches = after.match(/<w:r[ >]/g) ?? [];
            expect(runMatches.length).toBe(2);
        });
    });

    it("removes proofErr elements before merging", async () => {
        await withTempDir(async (dir) => {
            const body =
                "<w:body><w:p>" +
                "<w:r><w:t>foo</w:t></w:r>" +
                '<w:proofErr w:type="spellStart"/>' +
                "<w:r><w:t>bar</w:t></w:r>" +
                '<w:proofErr w:type="spellEnd"/>' +
                "</w:p></w:body>";
            const docPath = await writeDoc(dir, body);

            const result = await mergeRuns(dir);
            expect(result.count).toBe(1);

            const after = await fs.readFile(docPath, "utf8");
            expect(after).not.toMatch(/proofErr/);
            expect(after).toMatch(/foobar/);
        });
    });
});
