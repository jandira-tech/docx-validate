import { describe, expect, it } from "vitest";
import path from "node:path";
import { promises as fs } from "node:fs";
import { withTempDir } from "../src/lib/run-cli";
import { BaseSchemaValidator } from "../src/scripts/office/validators/base";
import type { BaseSchemaValidatorOptions } from "../src/scripts/office/validators/base";
import JSZip from "jszip";

class TestValidator extends BaseSchemaValidator {
    constructor(opts: BaseSchemaValidatorOptions) {
        super(opts);
    }
    async validate() {
        return { valid: true, issues: [] };
    }
    async repair() {
        return 0;
    }
    public async testGetOriginalFileErrors(xmlFile: string) {
        return this._getOriginalFileErrors(xmlFile);
    }
}

describe("BaseSchemaValidator zip slip prevention", () => {
    it("prevents extracting files outside temp dir during original docx validation", async () => {
        await withTempDir(async (dir) => {
            const unpackedDir = path.join(dir, "unpacked");
            await fs.mkdir(unpackedDir, { recursive: true });
            const xmlFile = path.join(unpackedDir, "../../evil.xml");

            const docxPath = path.join(dir, "original.docx");
            const maliciousZip = new JSZip();
            maliciousZip.file("dummy.xml", "<dummy/>");
            const buf = await maliciousZip.generateAsync({ type: "nodebuffer" });
            await fs.writeFile(docxPath, buf);

            const originalLoadAsync = JSZip.loadAsync;
            JSZip.loadAsync = async function loadAsyncPatched(data, options) {
                const zip = await originalLoadAsync.call(this, data, options);
                // Force an entry with path traversal
                zip.files["../../evil.xml"] = {
                    name: "../../evil.xml",
                    dir: false,
                    async: async () => Buffer.from("<evil/>"),
                } as unknown as JSZip.JSZipObject;
                // JSZip.file() uses .files directly in some versions, or iterates.
                // We mock the file() method to ensure it returns this entry
                const originalFileMethod = zip.file.bind(zip);
                zip.file = function(name: string | RegExp) {
                    if (name === "../../evil.xml") return zip.files["../../evil.xml"] as JSZip.JSZipObject;
                    return originalFileMethod(name as string) as JSZip.JSZipObject;
                } as unknown as typeof JSZip.prototype.file;
                return zip;
            };

            const validator = new TestValidator({
                unpackedDir,
                originalFile: docxPath,
                verbose: true,
                schemasDir: "",
                profile: "lenient",
            });

            try {
                await validator.testGetOriginalFileErrors(xmlFile);
            } catch (e) {
                expect((e as Error).message).toMatch(/Refusing to extract entry outside output dir/);
                return;
            } finally {
                JSZip.loadAsync = originalLoadAsync;
            }
            throw new Error("Should have thrown zip slip error");
        });
    });
});
