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
import { validateRedlining } from "../src/scripts/office/validators/redlining";

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';

function wrapDoc(body: string): string {
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:document ${W}>${body}</w:document>`;
}

async function writeUnpacked(dir: string, body: string): Promise<void> {
    const wordDir = path.join(dir, "word");
    await fs.mkdir(wordDir, { recursive: true });
    await fs.writeFile(path.join(wordDir, "document.xml"), wrapDoc(body), "utf8");
}

describe("validateRedlining zip slip prevention", () => {
    it("returns an unpack error when a zip entry escapes the extract dir", async () => {
        await withTempDir(async (dir) => {
            const body = '<w:body><w:p><w:ins w:author="Ritapolis"><w:r><w:t>Hello world</w:t></w:r></w:ins></w:p></w:body>';
            const unpacked = path.join(dir, "unpacked");
            await writeUnpacked(unpacked, body);

            const maliciousZip = new JSZip();
            maliciousZip.file("word/document.xml", wrapDoc(body));
            const buf = await maliciousZip.generateAsync({ type: "nodebuffer" });
            const docxPath = path.join(dir, "malicious.docx");
            await fs.writeFile(docxPath, buf);

            const originalLoadAsync = JSZip.loadAsync;
            JSZip.loadAsync = async function loadAsyncPatched(data, options) {
                const zip = await originalLoadAsync.call(this, data, options);
                zip.files["../../escaped.txt"] = {
                    name: "../../escaped.txt",
                    dir: false,
                    async: async () => Buffer.from("malicious"),
                } as unknown as JSZip.JSZipObject;
                return zip;
            };

            try {
                const result = await validateRedlining({
                    unpackedDir: unpacked,
                    originalDocx: docxPath,
                    author: "Ritapolis",
                    verbose: true,
                });
                expect(result.valid).toBe(false);
                expect(result.issues[0]?.message).toMatch(/Refusing to extract entry outside output dir/);
            } finally {
                JSZip.loadAsync = originalLoadAsync;
            }
        });
    });
});
