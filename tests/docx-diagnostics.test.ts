import { promises as fs } from "node:fs";
import path from "node:path";

import { describe, expect, it } from "vitest";
import { withTempDir } from "../src/lib/run-cli";
import {
    buildRepairPlanIssues,
    collectDocxSemanticInventory,
    compareDocxSemanticInventories,
} from "../src/scripts/office/validators/docx-diagnostics";

const W_NS = `xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"`;

async function writeXml(filePath: string, content: string): Promise<void> {
    await fs.mkdir(path.dirname(filePath), { recursive: true });
    await fs.writeFile(filePath, content, "utf-8");
}

function doc(body: string): string {
    return `<?xml version="1.0"?><w:document ${W_NS}><w:body>${body}</w:body></w:document>`;
}

describe("docx diagnostics", () => {
    it("reports formatting coverage loss without positional character diffs", async () => {
        await withTempDir(async (dir) => {
            const before = path.join(dir, "before");
            const after = path.join(dir, "after");
            await writeXml(
                path.join(before, "word", "document.xml"),
                doc(`<w:p><w:r><w:rPr><w:b/></w:rPr><w:t>ABCDEFGHIJ</w:t></w:r></w:p>`),
            );
            await writeXml(
                path.join(after, "word", "document.xml"),
                doc(`<w:p><w:r><w:rPr><w:b/></w:rPr><w:t>AB</w:t></w:r><w:r><w:t>CDEFGHIJ</w:t></w:r></w:p>`),
            );

            const issues = compareDocxSemanticInventories(
                await collectDocxSemanticInventory(before),
                await collectDocxSemanticInventory(after),
            );

            expect(issues).toHaveLength(1);
            expect(issues[0]).toMatchObject({
                severity: "error",
                code: "repair-content-loss",
                path: "word/document.xml",
            });
            expect(issues[0].message).toContain("formatting 'bold': 10 → 2 (-8 formatted character(s))");
        });
    });

    it("reports comments and tracked-change element loss by semantic bucket", async () => {
        await withTempDir(async (dir) => {
            const before = path.join(dir, "before");
            const after = path.join(dir, "after");
            await writeXml(
                path.join(before, "word", "document.xml"),
                doc(
                    `<w:p><w:ins w:id="1" w:author="A"><w:r><w:t>new</w:t></w:r></w:ins><w:del w:id="2" w:author="A"><w:r><w:delText>old</w:delText></w:r></w:del></w:p>`,
                ),
            );
            await writeXml(
                path.join(before, "word", "comments.xml"),
                `<?xml version="1.0"?><w:comments ${W_NS}><w:comment w:id="0" w:author="A"><w:p><w:r><w:t>note</w:t></w:r></w:p></w:comment></w:comments>`,
            );
            await writeXml(path.join(after, "word", "document.xml"), doc(`<w:p><w:r><w:t>new</w:t></w:r></w:p>`));
            await writeXml(path.join(after, "word", "comments.xml"), `<?xml version="1.0"?><w:comments ${W_NS}/>`);

            const issues = compareDocxSemanticInventories(
                await collectDocxSemanticInventory(before),
                await collectDocxSemanticInventory(after),
            );
            const messages = issues.map((issue) => issue.message).join("\n");

            expect(messages).toContain("comment 'comment entry': 1 → 0 (-1 comment(s))");
            expect(messages).toContain("tracked change 'ins': 1 → 0 (-1 element(s))");
            expect(messages).toContain("tracked change 'del': 1 → 0 (-1 element(s))");
            expect(messages).toContain("text 'deleted text': 3 → 0 (-3 character(s))");
            expect(messages).toContain("text 'comment text': 4 → 0 (-4 character(s))");
        });
    });

    it("emits pre-repair plans for known repairable issue codes", () => {
        const issues = buildRepairPlanIssues([
            { severity: "error", path: "word/document.xml", code: "ws-missing-preserve", message: "missing xml:space" },
            { severity: "error", path: "word/document.xml", code: "ws-missing-preserve", message: "missing xml:space" },
            { severity: "error", path: "word/document.xml", code: "comment-orphan-start", message: "orphan" },
        ]);

        expect(issues).toHaveLength(2);
        expect(issues[0]).toMatchObject({ code: "repair-plan-unavailable", path: "word/document.xml" });
        expect(issues[0].message).toContain("Before repair found 1 [comment-orphan-start] issue(s)");
        expect(issues[1]).toMatchObject({ code: "repair-plan", path: "word/document.xml" });
        expect(issues[1].message).toContain("Before repair found 2 [ws-missing-preserve] issue(s)");
        expect(issues[1].message).toContain("add xml:space='preserve'");
    });
});
