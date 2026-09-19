import { promises as fs } from "node:fs";
import path from "node:path";
import { describe, expect, it } from "vitest";

import { withTempDir } from "../src/lib/run-cli";
import { DOCXSchemaValidator } from "../src/scripts/office/validators/docx";

describe("DOCXSchemaValidator relationship target path traversal", () => {
    it("reports an error or drops relationship when target escapes unpacked directory", async () => {
        await withTempDir(async (dir) => {
            const unpackedDir = path.join(dir, "unpacked");
            await fs.mkdir(unpackedDir, { recursive: true });
            const wordRelsDir = path.join(unpackedDir, "word", "_rels");
            await fs.mkdir(wordRelsDir, { recursive: true });

            const relsContent = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../../../../../etc/passwd"/>
</Relationships>`;
            await fs.writeFile(path.join(wordRelsDir, "document.xml.rels"), relsContent, "utf8");

            await fs.writeFile(path.join(unpackedDir, "word", "document.xml"), `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"></w:document>`, "utf8");

            await fs.writeFile(path.join(unpackedDir, "[Content_Types].xml"), `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>`, "utf8");

            await fs.mkdir(path.join(unpackedDir, "_rels"), { recursive: true });
            const rootRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`;
            await fs.writeFile(path.join(unpackedDir, "_rels", ".rels"), rootRels, "utf8");


            const validator = new DOCXSchemaValidator({
                unpackedDir,
                originalFile: path.join(dir, "foo.docx"),
                author: "author",
                verbose: false,
            });

            const result = await validator.validate();
            // It drops the relationship so we don't read /etc/passwd or report /etc/passwd missing
            expect(result.issues.some(issue => issue.message.includes("/etc/passwd"))).toBe(false);
        });
    });
});
