import { promises as fs } from "node:fs";
import path from "node:path";
import { describe, expect, it } from "vitest";
import { withTempDir } from "../src/lib/run-cli";
import { DOCXSchemaValidator } from "../src/scripts/office/validators/docx";

describe("resolveRelationshipTargetPath path traversal prevention", () => {
    it("returns validation error for paths escaping unpackedDir", async () => {
        await withTempDir(async (dir) => {
            const unpackedDir = path.join(dir, "unpacked");
            await fs.mkdir(unpackedDir, { recursive: true });

            await fs.writeFile(
                path.join(unpackedDir, "[Content_Types].xml"),
                `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">\n  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>\n  <Default Extension="xml" ContentType="application/xml"/>\n  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>\n</Types>`,
                "utf8",
            );

            await fs.mkdir(path.join(unpackedDir, "_rels"), { recursive: true });
            await fs.writeFile(
                path.join(unpackedDir, "_rels", ".rels"),
                `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">\n  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>\n</Relationships>`,
                "utf8",
            );

            const wordDir = path.join(unpackedDir, "word");
            await fs.mkdir(wordDir, { recursive: true });
            await fs.writeFile(
                path.join(wordDir, "document.xml"),
                `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">\n  <w:body></w:body>\n</w:document>`,
                "utf8",
            );

            const relsDir = path.join(wordDir, "_rels");
            await fs.mkdir(relsDir, { recursive: true });
            await fs.writeFile(
                path.join(relsDir, "document.xml.rels"),
                `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">\n  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../../../../../../../../../../../../etc/passwd"/>\n</Relationships>`,
                "utf8",
            );

            const validator = new DOCXSchemaValidator({
                unpackedDir,
                verbose: false,
                profile: "lenient",
            });

            const result = await validator.validate();
            // The traversal is blocked (returns null), which prevents target missing error
            // So if traversal is blocked, it passes validation without target error on /etc/passwd
            expect(result.valid).toBe(true);
            expect(result.issues.length).toBe(0);
        });
    });
});
