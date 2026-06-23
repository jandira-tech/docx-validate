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
    it("throws an error when zip contains paths that escape the target dir", async () => {
        await withTempDir(async (dir) => {
            const body = '<w:body><w:p><w:ins w:author="Ritapolis"><w:r><w:t>Hello world</w:t></w:r></w:ins></w:p></w:body>';
            const unpacked = path.join(dir, "unpacked");
            await writeUnpacked(unpacked, body);

            // Construct malicious zip
            const maliciousZip = new JSZip();
            maliciousZip.file("word/document.xml", wrapDoc(body));

            const buf = await maliciousZip.generateAsync({ type: "nodebuffer" });
            const docxPath = path.join(dir, "malicious.docx");
            await fs.writeFile(docxPath, buf);

            const originalLoadAsync = JSZip.loadAsync;
            JSZip.loadAsync = async function (data) {
                const zip = await originalLoadAsync.call(this, data);
                zip.files["../../escaped.txt"] = {
                    name: "../../escaped.txt",
                    dir: false,
                    async: async () => Buffer.from("malicious"),
                } as unknown as JSZip.JSZipObject;
                return zip;
            };

            const result = await validateRedlining({
                unpackedDir: unpacked,
                originalDocx: docxPath,
                author: "Ritapolis",
                verbose: true,
            });

            JSZip.loadAsync = originalLoadAsync;

            expect(result.valid).toBe(false);
            expect(result.issues[0].message).toMatch(/Refusing to extract entry outside output dir/);
        });
    });
});
