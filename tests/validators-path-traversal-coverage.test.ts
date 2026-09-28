import { promises as fs } from "node:fs";
import path from "node:path";
import { describe, expect, it } from "vitest";
import { withTempDir } from "../src/lib/run-cli";
import { DOCXSchemaValidator } from "../src/scripts/office/validators/docx";

describe("Path Traversal Prevention Coverage", () => {
    it("handles path traversal correctly in validator", async () => {
        await withTempDir(async (dir) => {
            const unpackedDir = path.join(dir, "unpacked");
            await fs.mkdir(path.join(unpackedDir, "word", "_rels"), { recursive: true });

            const docXml = `<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body></w:body></w:document>`;
            await fs.writeFile(path.join(unpackedDir, "word", "document.xml"), docXml);

            const relsXml = `<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../../../secret.txt"/>
    <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="/../../../secret2.txt"/>
</Relationships>`;
            await fs.writeFile(path.join(unpackedDir, "word", "_rels", "document.xml.rels"), relsXml);

            const validator = new DOCXSchemaValidator({
                unpackedDir,
                profile: "strict",
            });

            // Invoke validation which will parse relationships
            const result = await validator.validate();

            // We expect broken relationships for the traversed targets
            expect(result.issues.some((i) => i.message.includes("Broken reference to ../../../secret.txt"))).toBe(true);
            expect(result.issues.some((i) => i.message.includes("Broken reference to /../../../secret2.txt"))).toBe(true);
        });
    });
});
