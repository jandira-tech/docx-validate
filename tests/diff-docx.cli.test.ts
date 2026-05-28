import { promises as fs } from "node:fs";
import path from "node:path";
import { describe, expect, it } from "vitest";
import { withTempDir } from "../src/lib/run-cli";
import { runDiffDocx } from "../scripts/diff-docx";

const W_NS = `xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"`;
const docXml = (body: string) => `<?xml version="1.0"?><w:document ${W_NS}><w:body>${body}</w:body></w:document>`;

async function writeUnpacked(dir: string, body: string): Promise<void> {
    await fs.mkdir(path.join(dir, "word"), { recursive: true });
    await fs.writeFile(path.join(dir, "word", "document.xml"), docXml(body), "utf-8");
}

describe("diff-docx CLI", () => {
    it("exits 0 when only warn/info diffs (table reshape)", async () => {
        await withTempDir(async (dir) => {
            const a = path.join(dir, "a");
            const b = path.join(dir, "b");
            await writeUnpacked(a, `<w:tbl><w:tblGrid><w:gridCol/><w:gridCol/></w:tblGrid><w:tr><w:tc/><w:tc/></w:tr></w:tbl>`);
            await writeUnpacked(b, `<w:tbl><w:tblGrid><w:gridCol/><w:gridCol/><w:gridCol/></w:tblGrid><w:tr><w:tc/><w:tc/><w:tc/></w:tr></w:tbl>`);
            const { code, markdown } = await runDiffDocx([a, b]);
            expect(code).toBe(0);
            expect(markdown).toContain("## Added");
        });
    });

    it("exits non-zero when a Content element is lost (paragraph removed)", async () => {
        await withTempDir(async (dir) => {
            const a = path.join(dir, "a");
            const b = path.join(dir, "b");
            await writeUnpacked(a, `<w:p><w:r><w:t>one</w:t></w:r></w:p><w:p><w:r><w:t>two</w:t></w:r></w:p>`);
            await writeUnpacked(b, `<w:p><w:r><w:t>one</w:t></w:r></w:p>`);
            const { code } = await runDiffDocx([a, b]);
            expect(code).toBe(1);
        });
    });

    it("rejects an invalid --profile with a non-zero code", async () => {
        await withTempDir(async (dir) => {
            const a = path.join(dir, "a");
            const b = path.join(dir, "b");
            const { code, markdown } = await runDiffDocx([a, b, "--profile", "bogus"]);
            expect(code).not.toBe(0);
            expect(markdown).toContain("Invalid --profile");
        });
    });
});
