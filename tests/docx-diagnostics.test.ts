import { promises as fs } from "node:fs";
import path from "node:path";

import { describe, expect, it } from "vitest";
import { withTempDir } from "../src/lib/run-cli";
import {
    buildRepairPlanIssues,
    collectDocxSemanticInventory,
    compareDocxSemanticInventories,
    severityClassFor,
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
    it("classifies categories into severity classes", () => {
        expect(severityClassFor("text")).toBe("content");
        expect(severityClassFor("document structure")).toBe("content");
        expect(severityClassFor("tracked change")).toBe("content");
        expect(severityClassFor("content type")).toBe("content");
        expect(severityClassFor("numbering")).toBe("content");
        expect(severityClassFor("table shape")).toBe("table-shape");
        expect(severityClassFor("section geometry")).toBe("section-geometry");
        expect(severityClassFor("image shape")).toBe("image-shape");
        expect(severityClassFor("formatting")).toBe("formatting");
        expect(severityClassFor("style reference")).toBe("formatting");
        expect(severityClassFor("inline mark")).toBe("atomic-marks");
        expect(severityClassFor("package asset")).toBe("bookkeeping");
        expect(severityClassFor("relationship")).toBe("bookkeeping");
        expect(severityClassFor("comment thread")).toBe("bookkeeping");
        expect(severityClassFor("comment")).toBe("content");
    });

    it("tiers repair loss: content → error, other classes → fidelity warning", async () => {
        await withTempDir(async (dir) => {
            const before = path.join(dir, "before");
            const after = path.join(dir, "after");
            await writeXml(
                path.join(before, "word", "document.xml"),
                doc(`<w:p><w:r><w:t>hello</w:t><w:br/></w:r></w:p><w:p><w:r><w:t>world</w:t></w:r></w:p>`),
            );
            // after: lost one paragraph (content) AND the line break (atomic mark)
            await writeXml(path.join(after, "word", "document.xml"), doc(`<w:p><w:r><w:t>hello</w:t></w:r></w:p>`));

            const issues = compareDocxSemanticInventories(
                await collectDocxSemanticInventory(before),
                await collectDocxSemanticInventory(after),
            );
            const byCode = (code: string) => issues.filter((i) => i.code === code);
            expect(byCode("repair-content-loss").some((i) => i.severity === "error")).toBe(true);
            expect(byCode("repair-fidelity-loss").some((i) => i.severity === "warning")).toBe(true);
            // the line-break loss must NOT be an error
            expect(byCode("repair-content-loss").some((i) => i.message.includes("line break"))).toBe(false);
        });
    });

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
                severity: "warning",
                code: "repair-fidelity-loss",
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

    it("counts in-run atomic marks by type and excludes tab-stop definitions", async () => {
        await withTempDir(async (dir) => {
            const root = path.join(dir, "u");
            await writeXml(
                path.join(root, "word", "document.xml"),
                doc(
                    `<w:p><w:pPr><w:tabs><w:tab w:val="left" w:pos="720"/></w:tabs></w:pPr>` +
                        `<w:r><w:br/><w:br w:type="page"/><w:br w:type="column"/><w:tab/><w:sym w:font="Wingdings" w:char="F0E0"/><w:cr/><w:softHyphen/><w:noBreakHyphen/></w:r>` +
                        `<w:r><w:br w:type="separator"/></w:r></w:p>`,
                ),
            );
            const inv = await collectDocxSemanticInventory(root);
            const get = (label: string) =>
                [...inv.counters.values()].find((c) => c.category === "inline mark" && c.label === label)?.count ?? 0;
            expect(get("line break")).toBe(1); // the bare <w:br/>
            expect(get("page break")).toBe(1);
            expect(get("column break")).toBe(1);
            expect(get("tab")).toBe(1); // only the run-level <w:tab/>, NOT the <w:tabs> stop
            expect(get("symbol")).toBe(1);
            expect(get("carriage return")).toBe(1);
            expect(get("soft hyphen")).toBe(1);
            expect(get("non-breaking hyphen")).toBe(1);
            // separator break is excluded entirely
            expect([...inv.counters.values()].some((c) => c.label.includes("separator"))).toBe(false);
        });
    });

    it("captures table shape as rows×cols with a tblGrid-absent fallback", async () => {
        await withTempDir(async (dir) => {
            const root = path.join(dir, "u");
            await writeXml(
                path.join(root, "word", "document.xml"),
                doc(
                    // table A: explicit 1-row, 3-col grid, one cell spans 2
                    `<w:tbl><w:tblGrid><w:gridCol/><w:gridCol/><w:gridCol/></w:tblGrid>` +
                        `<w:tr><w:tc><w:tcPr><w:gridSpan w:val="2"/></w:tcPr></w:tc><w:tc/></w:tr></w:tbl>` +
                        // table B: NO tblGrid, 2 rows; first row has 2 cells, one with gridSpan=2 → 3 cols
                        `<w:tbl><w:tr><w:tc><w:tcPr><w:gridSpan w:val="2"/></w:tcPr></w:tc><w:tc/></w:tr><w:tr><w:tc/><w:tc/><w:tc/></w:tr></w:tbl>`,
                ),
            );
            const inv = await collectDocxSemanticInventory(root);
            const shapes = [...inv.counters.values()].filter((c) => c.category === "table shape");
            const shape = (label: string) => shapes.find((c) => c.label === label)?.count ?? 0;
            expect(shape("table 1×3")).toBe(1); // table A
            expect(shape("table 2×3")).toBe(1); // table B via gridSpan fallback
            expect(shape("merged cell gridSpan=2")).toBe(2); // one in each table
        });
    });

    it("collects section geometry and image shape only under the strict profile", async () => {
        await withTempDir(async (dir) => {
            const root = path.join(dir, "u");
            await writeXml(
                path.join(root, "word", "document.xml"),
                doc(
                    `<w:p><w:r><w:drawing><wp:inline xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing">` +
                        `<wp:extent cx="1905000" cy="1270000"/></wp:inline></w:drawing></w:r></w:p>` +
                        `<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/><w:cols w:num="2"/></w:sectPr>`,
                ),
            );
            const lenient = await collectDocxSemanticInventory(root); // default lenient
            const strict = await collectDocxSemanticInventory(root, "strict");
            const has = (inv: Awaited<ReturnType<typeof collectDocxSemanticInventory>>, category: string) =>
                [...inv.counters.values()].some((c) => c.category === category);

            expect(has(lenient, "section geometry")).toBe(false);
            expect(has(lenient, "image shape")).toBe(false);
            expect(has(strict, "section geometry")).toBe(true);
            expect(has(strict, "image shape")).toBe(true);

            const strictVals = [...strict.counters.values()];
            expect(strictVals.find((c) => c.category === "section geometry" && c.label === "section portrait 12240×15840")?.count).toBe(1);
            expect(strictVals.find((c) => c.category === "section geometry" && c.label === "section columns=2")?.count).toBe(1);
            expect(strictVals.find((c) => c.category === "image shape" && c.label === "image ~1905000×1270000 inline")?.count).toBe(1);

            const wordValid = await collectDocxSemanticInventory(root, "word-valid");
            expect(has(wordValid, "section geometry")).toBe(false); // word-valid behaves like lenient
        });
    });

    it("treats a missing or non-numeric wp:extent dimension as ~0 without throwing", async () => {
        await withTempDir(async (dir) => {
            const root = path.join(dir, "u");
            await writeXml(
                path.join(root, "word", "document.xml"),
                doc(
                    `<w:p><w:r><w:drawing><wp:inline xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing">` +
                        `<wp:extent cy="1270000"/></wp:inline></w:drawing></w:r></w:p>`,
                ),
            );
            const strict = await collectDocxSemanticInventory(root, "strict");
            const strictVals = [...strict.counters.values()];
            expect(strictVals.find((c) => c.category === "image shape" && c.label === "image ~0×1270000 inline")?.count).toBe(1);
        });
    });

    it("emits pre-repair plans for known repairable issue codes", () => {
        const issues = buildRepairPlanIssues([
            {
                severity: "error",
                path: "word/document.xml",
                code: "ws-missing-preserve",
                message: "missing xml:space",
            },
            {
                severity: "error",
                path: "word/document.xml",
                code: "ws-missing-preserve",
                message: "missing xml:space",
            },
            {
                severity: "error",
                path: "word/document.xml",
                code: "comment-orphan-start",
                message: "orphan",
            },
        ]);

        expect(issues).toHaveLength(2);
        expect(issues[0]).toMatchObject({ code: "repair-plan-unavailable", path: "word/document.xml" });
        expect(issues[0].message).toContain("Before repair found 1 [comment-orphan-start] issue(s)");
        expect(issues[1]).toMatchObject({ code: "repair-plan", path: "word/document.xml" });
        expect(issues[1].message).toContain("Before repair found 2 [ws-missing-preserve] issue(s)");
        expect(issues[1].message).toContain("add xml:space='preserve'");
    });
});
