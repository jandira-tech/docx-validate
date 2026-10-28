import { promises as fs } from "node:fs";
import path from "node:path";

import { describe, expect, it } from "vitest";
import { withTempDir } from "../src/lib/run-cli";
import type { ValidationIssue } from "../src/lib/types";
import {
    buildRepairPlanIssues,
    collectDocxSemanticInventory,
    compareDocxSemanticInventories,
    severityClassFor,
} from "../src/scripts/office/validators/docx-diagnostics";
import { fileURLToPath } from "node:url";

import { buildDocx, MINIMAL_PARTS } from "./support/minimal-docx";
import { validate } from "../src/scripts/office/validate";

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

    // Adapted from bot PR #198's test. collectText labels text by its
    // enclosing tracked change or comment; these shapes pin each label and
    // the precedence between them (ins, del, moveFrom, moveTo, comment).
    it("labels text by its enclosing tracked change or comment, in a fixed precedence", async () => {
        await withTempDir(async (dir) => {
            const root = path.join(dir, "u");
            await writeXml(
                path.join(root, "word", "document.xml"),
                doc(
                    `<w:p>` +
                        `<w:fldSimple w:instr="PAGE"/><w:fldSimple/>` +
                        `<w:del><w:r><w:instrText>deleted field</w:instrText></w:r></w:del>` +
                        `<w:r><w:instrText>visible field</w:instrText></w:r>` +
                        `<w:r><w:delInstrText>delInstrText</w:delInstrText><w:delText>gone</w:delText><w:t></w:t></w:r>` +
                        `<w:moveFrom><w:r><w:t>moveFrom</w:t></w:r><w:ins><w:r><w:t>ins in move</w:t></w:r></w:ins></w:moveFrom>` +
                        `<w:moveTo><w:del><w:r><w:t>del</w:t></w:r></w:del><w:r><w:t>moveTo</w:t></w:r></w:moveTo>` +
                        `<w:ins><x:wrap xmlns:x="urn:example"><w:r><w:t>wrapped</w:t></w:r></x:wrap></w:ins>` +
                        `<w:r><w:t>plain</w:t></w:r>` +
                        `</w:p>`,
                ),
            );
            await writeXml(
                path.join(root, "word", "comments.xml"),
                `<?xml version="1.0"?><w:comments ${W_NS}><w:comment w:id="0"><w:p><w:moveFrom><w:r><w:t>moved note</w:t></w:r></w:moveFrom><w:r><w:t>note</w:t></w:r></w:p></w:comment></w:comments>`,
            );
            const inventory = await collectDocxSemanticInventory(root);
            const text = new Map<string, number>();
            for (const counter of inventory.counters.values()) {
                if (counter.category === "text") text.set(counter.label, (text.get(counter.label) ?? 0) + counter.count);
            }

            expect(Object.fromEntries(text)).toEqual({
                "field instruction": 4 + 13, // PAGE, visible field
                "deleted field instruction": 13 + 12, // deleted field, delInstrText
                "deleted text": 4 + 3, // gone, del inside moveTo
                "moved-from text": 8 + 10, // moveFrom, moved note
                "moved-to text": 6, // moveTo
                "inserted text": 11 + 7, // ins in move, wrapped
                "comment text": 4, // note
                "visible text": 5, // plain
            });
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

    it("reports cleared and retained pre-repair groups from their outcomes", () => {
        const retained: ValidationIssue = { severity: "error", path: "word/document.xml", code: "comment-orphan-start", message: "orphan" };
        const issues = buildRepairPlanIssues(
            [
                { severity: "error", path: "word/document.xml", code: "ws-missing-preserve", message: "missing xml:space" },
                { severity: "error", path: "word/document.xml", code: "ws-missing-preserve", message: "missing xml:space" },
                retained,
            ],
            [retained],
        );

        expect(issues).toHaveLength(2);
        expect(issues[0]).toMatchObject({ severity: "info", code: "repair-plan-unavailable", path: "word/document.xml" });
        expect(issues[0].message).toContain("Before repair found 1 [comment-orphan-start] issue(s)");
        expect(issues[0].message).toContain("Auto-repair cleared 0; 1 remain");
        expect(issues[1]).toMatchObject({ severity: "info", code: "repair-plan", path: "word/document.xml" });
        expect(issues[1].message).toContain("Before repair found 2 [ws-missing-preserve] issue(s)");
        expect(issues[1].message).toContain("Auto-repair cleared 2; 0 remain");
    });

    it("gives story-reference guidance that needs no existing id (#623)", () => {
        const missingId: ValidationIssue = {
            severity: "error",
            path: "word/document.xml",
            code: "word-header-footer-relationship",
            message: "<w:footerReference> has a missing r:id",
        };
        // No repair() changes this code, so the finding remains after repair.
        const [issue] = buildRepairPlanIssues([missingId], [missingId]);
        expect(issue?.code).toBe("repair-plan-unavailable");
        expect(issue?.message).toContain(
            "set the w:headerReference or w:footerReference r:id to a non-empty id naming a relationship whose Type is the Transitional or ISO Strict header (respectively footer) relationship URI, adding such a relationship to the part's relationships when none exists",
        );
        expect(issue?.message).not.toContain("add a relationship with that Id and Type");
    });
});

it.each(["[CONTENT_TYPES].XML", "[content_types].xml"])("inventories declarations in stored stream %s (#396)", async (name) => {
    await withTempDir(async (dir) => {
        await writeXml(
            path.join(dir, name),
            `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Override PartName="/word/x.xml" ContentType="application/xml"/><Default Extension="bin" ContentType="application/octet-stream"/></Types>`,
        );
        const inventory = await collectDocxSemanticInventory(dir);
        expect([...inventory.counters.values()].filter((counter) => counter.category === "content type")).toEqual([
            expect.objectContaining({ path: name, label: "default 'application/octet-stream'", count: 1 }),
            expect.objectContaining({ path: name, label: "override 'application/xml'", count: 1 }),
        ]);
    });
});

/*
 * #745 - the inventory counts styles in the parts the main document and
 * the reached glossaries relate, found by relationship as the style checks
 * find them (#688), not in every part named `styles.xml`.
 */
describe("style counters read the related styles parts (#745)", () => {
    const R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    const RELS_NS = `xmlns="http://schemas.openxmlformats.org/package/2006/relationships"`;
    const relationship = (id: string, type: string, target: string): string =>
        `<Relationship Id="${id}" Type="${R}/${type}" Target="${target}"/>`;
    const packageRels = (target: string): string =>
        `<?xml version="1.0"?><Relationships ${RELS_NS}>${relationship("rId1", "officeDocument", target)}</Relationships>`;
    const docRels = (...relationships: string[]): string =>
        `<?xml version="1.0"?><Relationships ${RELS_NS}>${relationships.join("")}</Relationships>`;
    const stylesPart = (...styleIds: string[]): string =>
        `<?xml version="1.0"?><w:styles ${W_NS}>` +
        `${styleIds.map((id) => `<w:style w:type="paragraph" w:styleId="${id}"><w:name w:val="${id}"/></w:style>`).join("")}</w:styles>`;
    const glossaryDoc =
        `<?xml version="1.0"?><w:glossaryDocument ${W_NS}>` +
        `<w:docParts><w:docPart><w:docPartBody><w:p/></w:docPartBody></w:docPart></w:docParts></w:glossaryDocument>`;

    const writePackage = async (dir: string, parts: Record<string, string>): Promise<void> => {
        for (const [name, content] of Object.entries(parts)) await writeXml(path.join(dir, name), content);
    };

    const styleDefinitionCount = (
        inventory: { counters: ReadonlyMap<string, { category: string; path: string; count: number }> },
        partPath: string,
    ): number =>
        [...inventory.counters.values()]
            .filter((counter) => counter.category === "style definition" && counter.path === partPath)
            .reduce((total, counter) => total + counter.count, 0);

    /** The main document relates its styles as `word/sub/s.xml`; no part named `styles.xml` exists. */
    const RENAMED_STYLES: Record<string, string> = {
        "_rels/.rels": packageRels("word/document.xml"),
        "word/document.xml": doc(""),
        "word/_rels/document.xml.rels": docRels(relationship("rId1", "styles", "sub/s.xml")),
        "word/sub/s.xml": stylesPart("Heading1", "Renamed"),
    };

    /** The plain shape: `word/styles.xml`, related by the main document. */
    const PLAIN_STYLES: Record<string, string> = {
        ...RENAMED_STYLES,
        "word/_rels/document.xml.rels": docRels(relationship("rId1", "styles", "styles.xml")),
        "word/styles.xml": stylesPart("Heading1", "Plain"),
    };
    delete PLAIN_STYLES["word/sub/s.xml"];

    it("counts a renamed styles part the main document relates", async () => {
        await withTempDir(async (dir) => {
            await writePackage(dir, RENAMED_STYLES);
            expect(styleDefinitionCount(await collectDocxSemanticInventory(dir), "word/sub/s.xml")).toBe(2);
        });
    });

    it("keeps counting the plain word/styles.xml the main document relates", async () => {
        await withTempDir(async (dir) => {
            await writePackage(dir, PLAIN_STYLES);
            expect(styleDefinitionCount(await collectDocxSemanticInventory(dir), "word/styles.xml")).toBe(2);
        });
    });

    it("keeps counting word/styles.xml by name when the main document relates no styles part (#699)", async () => {
        await withTempDir(async (dir) => {
            await writePackage(dir, { ...PLAIN_STYLES, "word/_rels/document.xml.rels": docRels() });
            expect(styleDefinitionCount(await collectDocxSemanticInventory(dir), "word/styles.xml")).toBe(2);
        });
    });

    it("counts a renamed styles part a reached glossary relates", async () => {
        await withTempDir(async (dir) => {
            await writePackage(dir, {
                ...PLAIN_STYLES,
                "word/_rels/document.xml.rels": docRels(
                    relationship("rId1", "styles", "styles.xml"),
                    relationship("rId9", "glossaryDocument", "zglossary/document.xml"),
                ),
                "word/zglossary/document.xml": glossaryDoc,
                "word/zglossary/_rels/document.xml.rels": docRels(relationship("rId1", "styles", "s.xml")),
                "word/zglossary/s.xml": stylesPart("Glossary"),
            });
            expect(styleDefinitionCount(await collectDocxSemanticInventory(dir), "word/zglossary/s.xml")).toBe(1);
        });
    });

    it("does not count a glossary's styles.xml that nothing relates (#699)", async () => {
        await withTempDir(async (dir) => {
            await writePackage(dir, {
                ...PLAIN_STYLES,
                "word/_rels/document.xml.rels": docRels(
                    relationship("rId1", "styles", "styles.xml"),
                    relationship("rId9", "glossaryDocument", "zglossary/document.xml"),
                ),
                "word/zglossary/document.xml": glossaryDoc,
                "word/zglossary/_rels/document.xml.rels": docRels(),
                "word/zglossary/styles.xml": stylesPart("Unrelated"),
            });
            expect(styleDefinitionCount(await collectDocxSemanticInventory(dir), "word/zglossary/styles.xml")).toBe(0);
        });
    });
});

/*
 * #622 — repair plans report the observed before/after outcome. Every code `repairActionForIssue` maps to
 * a manual action is covered below: one specimen per code, validated with
 * `autoRepair`, asserting `repair-plan` only when the finding is gone after
 * repair and `repair-plan-unavailable` (carrying the manual action as
 * guidance) otherwise.
 */

describe("repair plan vs repair() coverage (#622)", () => {
    const W14_NS = `xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"`;
    const W16CID_NS = `xmlns:w16cid="http://schemas.microsoft.com/office/word/2016/wordml/cid"`;
    const W16CEX_NS = `xmlns:w16cex="http://schemas.microsoft.com/office/word/2018/wordml/cex"`;
    const R_NS = `xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"`;
    const R_URI = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    const WP_NS = `xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"`;
    const A_URI = "http://schemas.openxmlformats.org/drawingml/2006/main";
    const PIC_URI = "http://schemas.openxmlformats.org/drawingml/2006/picture";
    const RELS_NS = `xmlns="http://schemas.openxmlformats.org/package/2006/relationships"`;
    const WMLT = "application/vnd.openxmlformats-officedocument.wordprocessingml";
    const BROKEN_DIR = path.join(path.dirname(fileURLToPath(import.meta.url)), "fixtures", "broken");

    const PAIRED_IDS = `w14:paraId="11111111" w14:textId="11111111"`;
    const SECT =
        `<w:sectPr><w:pgSz w:w="12240" w:h="15840"/>` +
        `<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>`;

    /** MINIMAL_PARTS-shaped main document with `body` inserted before the sectPr. */
    const specDoc = (
        body: string,
        rootAttrs = `${W14_NS} xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="w14" ${R_NS}`,
    ): string =>
        `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n` +
        `<w:document ${W_NS} ${rootAttrs}><w:body><w:p ${PAIRED_IDS}><w:r><w:t>Hello</w:t></w:r></w:p>${body}${SECT}</w:body></w:document>`;

    const addOverride = (ct: string, part: string, type: string): string =>
        ct.replace("</Types>", `<Override PartName="/${part}" ContentType="${type}"/></Types>`);
    const addDefault = (ct: string, ext: string, type: string): string =>
        ct.replace("</Types>", `<Default Extension="${ext}" ContentType="${type}"/></Types>`);
    const addDocRel = (relsPart: string, id: string, typeSuffix: string, target: string): string =>
        relsPart.replace("</Relationships>", `<Relationship Id="${id}" Type="${R_URI}/${typeSuffix}" Target="${target}"/></Relationships>`);

    const COMMENTS_XML =
        `<?xml version="1.0"?><w:comments ${W_NS} ${W14_NS}>` +
        `<w:comment w:id="0" w:author="A" w:date="2026-01-01T00:00:00Z" w:initials="A"><w:p w14:paraId="11111111"/></w:comment></w:comments>`;
    const COMMENT_MARKERS = `<w:p ${PAIRED_IDS}><w:commentRangeStart w:id="0"/>x<w:commentRangeEnd w:id="0"/><w:r><w:commentReference w:id="0"/></w:r></w:p>`;
    const commentsIdsXml = (inner: string): string => `<?xml version="1.0"?><w16cid:commentsIds ${W16CID_NS}>${inner}</w16cid:commentsIds>`;
    const commentsExtensibleXml = (inner: string): string =>
        `<?xml version="1.0"?><w16cex:commentsExtensible ${W16CEX_NS}>${inner}</w16cex:commentsExtensible>`;

    /** MINIMAL_PARTS plus comments.xml, related by the main document (#896), and the optional threading parts, wired by content types. */
    function withCommentParts(opts: { ids?: string; extensible?: string }): Record<string, string> {
        const parts: Record<string, string> = {
            ...MINIMAL_PARTS,
            "word/document.xml": specDoc(COMMENT_MARKERS),
            "word/comments.xml": COMMENTS_XML,
        };
        parts["word/_rels/document.xml.rels"] = addDocRel(parts["word/_rels/document.xml.rels"]!, "rIdC", "comments", "comments.xml");
        parts["[Content_Types].xml"] = addOverride(parts["[Content_Types].xml"]!, "word/comments.xml", `${WMLT}.comments+xml`);
        if (opts.ids !== undefined) {
            parts["word/commentsIds.xml"] = opts.ids;
            parts["[Content_Types].xml"] = addOverride(parts["[Content_Types].xml"]!, "word/commentsIds.xml", `${WMLT}.commentsIds+xml`);
        }
        if (opts.extensible !== undefined) {
            parts["word/commentsExtensible.xml"] = opts.extensible;
            parts["[Content_Types].xml"] = addOverride(
                parts["[Content_Types].xml"]!,
                "word/commentsExtensible.xml",
                `${WMLT}.commentsExtensible+xml`,
            );
        }
        return parts;
    }

    const inlinePicBody =
        `<w:p ${PAIRED_IDS}><w:r><w:drawing><wp:inline ${WP_NS}>` +
        `<a:graphic xmlns:a="${A_URI}"><a:graphicData uri="${PIC_URI}">` +
        `<pic:pic xmlns:pic="${PIC_URI}">` +
        `<pic:nvPicPr><pic:cNvPr id="1" name="image1.png"/><pic:cNvPicPr/></pic:nvPicPr>` +
        `<pic:blipFill><a:blip xmlns:a="${A_URI}" r:embed="rId9"/><a:stretch xmlns:a="${A_URI}"><a:fillRect/></a:stretch></pic:blipFill>` +
        `<pic:spPr><a:xfrm xmlns:a="${A_URI}"><a:off x="0" y="0"/><a:ext cx="100" cy="100"/></a:xfrm>` +
        `<a:prstGeom xmlns:a="${A_URI}" prst="rect"><a:avLst/></a:prstGeom></pic:spPr>` +
        `</pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>`;

    /** `wp:inline` whose a:graphic precedes the required wp:extent/wp:docPr — repairInlinePictureScaffolds' case. */
    function scaffoldSpecimen(): Record<string, string | Uint8Array> {
        const parts: Record<string, string | Uint8Array> = {
            ...MINIMAL_PARTS,
            "word/document.xml": specDoc(inlinePicBody),
            "word/media/image1.png": new Uint8Array([0x89, 0x50, 0x4e, 0x47]),
        };
        parts["[Content_Types].xml"] = addDefault(parts["[Content_Types].xml"] as string, "png", "image/png");
        parts["word/_rels/document.xml.rels"] = addDocRel(
            parts["word/_rels/document.xml.rels"] as string,
            "rId9",
            "image",
            "media/image1.png",
        );
        return parts;
    }

    const numberingParts = (durableId: string): Record<string, string> => {
        const parts: Record<string, string> = { ...MINIMAL_PARTS };
        parts["[Content_Types].xml"] = addOverride(parts["[Content_Types].xml"]!, "word/numbering.xml", `${WMLT}.numbering+xml`);
        parts["word/_rels/document.xml.rels"] = addDocRel(parts["word/_rels/document.xml.rels"]!, "rId2", "numbering", "numbering.xml");
        parts["word/numbering.xml"] =
            `<?xml version="1.0"?><w:numbering ${W_NS} ${W16CID_NS}>` +
            `<w:abstractNum w:abstractNumId="0" w16cid:durableId="${durableId}"/><w:num w:numId="1" w:abstractNumId="0"/></w:numbering>`;
        return parts;
    };

    interface RepairPlanCase {
        code: string;
        /** Expected outcome for this particular specimen, verified against the findings. */
        expected: "repaired" | "unavailable";
        profile?: "word-valid" | "strict" | "lenient";
        parts?: Record<string, string | Uint8Array>;
        fixture?: string;
    }

    const REPAIR_PLAN_CASES: RepairPlanCase[] = [
        {
            code: "ws-missing-preserve",
            expected: "repaired",
            parts: { ...MINIMAL_PARTS, "word/document.xml": specDoc(`<w:p ${PAIRED_IDS}><w:r><w:t> lead</w:t></w:r></w:p>`) },
        },
        {
            code: "rels-empty-part",
            expected: "repaired",
            parts: {
                ...MINIMAL_PARTS,
                "word/_rels/styles.xml.rels": `<?xml version="1.0"?><Relationships ${RELS_NS}></Relationships>`,
            },
        },
        {
            code: "id-durable-overflow",
            expected: "repaired",
            parts: withCommentParts({ ids: commentsIdsXml(`<w16cid:commentId w16cid:paraId="11111111" w16cid:durableId="FFFFFFFF"/>`) }),
        },
        { code: "id-durable-decimal", expected: "repaired", parts: numberingParts("ABCDEF") },
        {
            code: "id-paraid-overflow",
            expected: "repaired",
            parts: { ...MINIMAL_PARTS, "word/document.xml": specDoc(`<w:p ${W14_NS} w14:paraId="FFFFFFFF" w14:textId="00000001"/>`) },
        },
        {
            code: "id-textid-overflow",
            expected: "repaired",
            parts: { ...MINIMAL_PARTS, "word/document.xml": specDoc(`<w:p ${W14_NS} w14:paraId="00000001" w14:textId="FFFFFFFF"/>`) },
        },
        {
            code: "ignorable-undeclared",
            expected: "repaired",
            parts: {
                ...MINIMAL_PARTS,
                "word/document.xml": specDoc(
                    "",
                    `${W14_NS} xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="w14 w15"`,
                ),
            },
        },
        { code: "paraid-missing-element", expected: "repaired", parts: { ...MINIMAL_PARTS, "word/document.xml": specDoc(`<w:p/>`) } },
        {
            code: "textid-missing-element",
            expected: "repaired",
            parts: { ...MINIMAL_PARTS, "word/document.xml": specDoc(`<w:p ${W14_NS} w14:paraId="00000001"/>`) },
        },
        {
            code: "style-reference-undefined",
            expected: "repaired",
            parts: {
                ...MINIMAL_PARTS,
                "word/document.xml": specDoc(
                    `<w:p ${PAIRED_IDS}><w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr><w:t>x</w:t></w:r></w:p>`,
                ),
            },
        },
        {
            code: "style-default-missing",
            expected: "repaired",
            parts: {
                ...MINIMAL_PARTS,
                "word/styles.xml": MINIMAL_PARTS["word/styles.xml"]!.replace(
                    `<w:style w:type="numbering" w:default="1" w:styleId="NoList"><w:name w:val="No List"/></w:style>`,
                    "",
                ),
            },
        },
        {
            code: "comment-thread-commentid-paraid-orphan",
            expected: "repaired",
            parts: withCommentParts({ ids: commentsIdsXml(`<w16cid:commentId w16cid:paraId="DEADBEEF" w16cid:durableId="22222222"/>`) }),
        },
        {
            code: "comment-thread-commentid-missing-paraid",
            expected: "unavailable",
            parts: withCommentParts({ ids: commentsIdsXml(`<w16cid:commentId w16cid:durableId="22222222"/>`) }),
        },
        {
            code: "comment-thread-commentid-missing-durableid",
            expected: "unavailable",
            parts: withCommentParts({ ids: commentsIdsXml(`<w16cid:commentId w16cid:paraId="11111111"/>`) }),
        },
        {
            code: "comment-thread-durableid-orphan",
            expected: "repaired",
            parts: withCommentParts({
                ids: commentsIdsXml(`<w16cid:commentId w16cid:paraId="11111111" w16cid:durableId="22222222"/>`),
                extensible: commentsExtensibleXml(`<w16cex:commentExtensible w16cex:durableId="DEADBEEF"/>`),
            }),
        },
        {
            code: "comment-thread-durableid-missing",
            expected: "unavailable",
            parts: withCommentParts({
                ids: commentsIdsXml(`<w16cid:commentId w16cid:paraId="11111111" w16cid:durableId="22222222"/>`),
                extensible: commentsExtensibleXml(`<w16cex:commentExtensible/>`),
            }),
        },
        {
            code: "word-drawing-scalar-whitespace",
            expected: "repaired",
            parts: { ...MINIMAL_PARTS, "word/document.xml": specDoc(`<wp:posOffset ${WP_NS}> 1 </wp:posOffset>`) },
        },
        { code: "word-drawing-arc-radius", expected: "unavailable", fixture: "package.drawing-arc-radius.docx" },
        { code: "word-empty-drawing", expected: "unavailable", fixture: "package.word-empty-drawing.docx" },
        { code: "word-drawing-extent-negative", expected: "unavailable", fixture: "package.word-drawing-extent-negative.docx" },
        {
            code: "word-anchor-missing-attribute",
            expected: "unavailable",
            parts: {
                ...MINIMAL_PARTS,
                "word/document.xml": specDoc(
                    `<w:p ${PAIRED_IDS}><w:r><w:drawing><wp:anchor ${WP_NS} locked="0" layoutInCell="1"/></w:drawing></w:r></w:p>`,
                ),
            },
        },
        {
            code: "word-position-offset-and-align",
            expected: "unavailable",
            parts: {
                ...MINIMAL_PARTS,
                "word/document.xml": specDoc(
                    `<w:p ${PAIRED_IDS}><w:r><w:drawing><wp:anchor ${WP_NS} locked="0" layoutInCell="1" allowOverlap="1">` +
                        `<wp:positionH><wp:posOffset>1</wp:posOffset><wp:align>left</wp:align></wp:positionH>` +
                        `<wp:positionV><wp:posOffset>1</wp:posOffset></wp:positionV>` +
                        `</wp:anchor></w:drawing></w:r></w:p>`,
                ),
            },
        },
        { code: "word-sdt-id-empty", expected: "unavailable", fixture: "package.word-sdt-id-empty.docx" },
        { code: "word-attached-toolbar-utf8", expected: "unavailable", fixture: "package.word-attached-toolbar-utf8.docx" },
        { code: "word-row-without-cell", expected: "unavailable", fixture: "package.word-row-without-cell.docx" },
        { code: "word-cell-ends-with-table", expected: "unavailable", fixture: "package.word-cell-ends-with-table.docx" },
        { code: "word-numbering-picture-bullet", expected: "unavailable", fixture: "package.word-numbering-picture-bullet.docx" },
        { code: "word-footnote-reference-missing", expected: "unavailable", fixture: "package.word-footnote-reference-missing.docx" },
        { code: "word-textbox-deleted-text", expected: "unavailable", fixture: "package.word-textbox-deleted-text.docx" },
        { code: "word-move-from-deleted-text", expected: "unavailable", fixture: "package.word-move-from-deleted-text.docx" },
        { code: "word-choice-requires-undeclared", expected: "unavailable", fixture: "package.word-choice-requires-undeclared.docx" },
        {
            code: "ct-uncovered-part",
            expected: "unavailable",
            parts: { ...MINIMAL_PARTS, "word/media/image1": new Uint8Array([0x89, 0x50, 0x4e, 0x47]) },
        },
        { code: "word-text-contains-paragraph", expected: "unavailable", fixture: "package.word-text-contains-paragraph.docx" },
        {
            code: "word-hyperlink-relationship",
            expected: "unavailable",
            parts: { ...MINIMAL_PARTS, "word/document.xml": specDoc(`<w:hyperlink r:id="rId9"><w:r><w:t>link</w:t></w:r></w:hyperlink>`) },
        },
        {
            code: "word-header-footer-relationship",
            expected: "unavailable",
            parts: (() => {
                // The session-32 shape from the issue: xmlns:r on w:document,
                // a footerReference r:id="rId9" first in the body sectPr, and
                // document.xml.rels kept without rId9.
                const parts: Record<string, string> = { ...MINIMAL_PARTS };
                parts["word/document.xml"] = parts["word/document.xml"]!.replace(
                    `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"`,
                    `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ${R_NS}`,
                ).replace("<w:sectPr>", `<w:sectPr><w:footerReference w:type="default" r:id="rId9"/>`);
                return parts;
            })(),
        },
        { code: "word-duplicate-theme-relationship", expected: "unavailable", fixture: "package.word-duplicate-theme.docx" },
        { code: "word-duplicate-font-table-relationship", expected: "unavailable", fixture: "package.word-duplicate-font-table.docx" },
        { code: "word-duplicate-web-settings-relationship", expected: "unavailable", fixture: "package.word-duplicate-web-settings.docx" },
        { code: "word-duplicate-settings-relationship", expected: "unavailable", fixture: "package.word-duplicate-settings.docx" },
        { code: "word-duplicate-styles-relationship", expected: "unavailable", fixture: "package.word-duplicate-styles.docx" },
        { code: "word-duplicate-numbering-relationship", expected: "unavailable", fixture: "package.word-duplicate-numbering.docx" },
        { code: "xsd-error", expected: "repaired", profile: "strict", parts: scaffoldSpecimen() },
        { code: "xsd-error", expected: "repaired", profile: "lenient", parts: scaffoldSpecimen() },
        {
            // repairCorePropertiesWhitespace / repairExtendedPropertiesWhitespace
            // only trim app.xml/core.xml scalars; an empty vt:filetime in
            // docProps/custom.xml keeps the finding, so the plan must not
            // promise this repair.
            code: "xsd-error",
            expected: "unavailable",
            profile: "strict",
            parts: (() => {
                const parts: Record<string, string> = { ...MINIMAL_PARTS };
                parts["[Content_Types].xml"] = addOverride(
                    parts["[Content_Types].xml"]!,
                    "docProps/custom.xml",
                    "application/vnd.openxmlformats-officedocument.custom-properties+xml",
                );
                parts["_rels/.rels"] = parts["_rels/.rels"]!.replace(
                    "</Relationships>",
                    `<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties" Target="docProps/custom.xml"/></Relationships>`,
                );
                parts["docProps/custom.xml"] =
                    `<?xml version="1.0"?><Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties" ` +
                    `xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">` +
                    `<property fmtid="{D5CDD505-2E9C-101B-9997-08002B2CF9AE}" pid="2" name="p"><vt:filetime/></property></Properties>`;
                return parts;
            })(),
        },
    ];

    it.each(REPAIR_PLAN_CASES)("plans [$expected] for $code", { timeout: 60_000 }, async (c) => {
        await withTempDir(async (dir) => {
            const file = path.join(dir, `${c.code}.docx`);
            if (c.fixture) {
                await fs.copyFile(path.join(BROKEN_DIR, c.fixture), file);
            } else {
                await fs.writeFile(file, await buildDocx(c.parts!));
            }
            const before = await validate(file, { profile: c.profile ?? "word-valid" });
            expect(before.issues.some((i) => i.code === c.code)).toBe(true);
            const result = await validate(file, { profile: c.profile ?? "word-valid", autoRepair: true });
            const plans = result.issues.filter((i) => i.code?.startsWith("repair-plan") && i.message.includes(`[${c.code}]`));
            expect(plans.length).toBeGreaterThan(0);
            for (const plan of plans) {
                const count = before.issues.filter((i) => i.code === c.code && i.path === plan.path).length;
                const remaining = result.issues.filter((i) => i.code === c.code && i.path === plan.path).length;
                expect(plan.message).toContain(`Before repair found ${count} [${c.code}] issue(s)`);
                expect(plan.message).toContain(`Auto-repair cleared ${Math.max(0, count - remaining)}; ${remaining} remain`);
                expect(plan.message).not.toContain("Auto-repair will");
            }
            if (c.expected === "repaired") {
                for (const plan of plans) {
                    expect(plan.code).toBe("repair-plan");
                    expect(plan.message).toMatch(/Auto-repair cleared [1-9]\d*; 0 remain/);
                }
                expect(result.repairs).toBeGreaterThan(0);
                expect(result.issues.some((i) => i.code === c.code)).toBe(false);
            } else {
                for (const plan of plans) {
                    expect(plan.code).toBe("repair-plan-unavailable");
                    expect(plan.message).not.toContain("Auto-repair will");
                    const remaining = result.issues.filter((i) => i.code === c.code && i.path === plan.path).length;
                    expect(plan.message).toContain(`Auto-repair cleared 0; ${remaining} remain`);
                    expect(plan.message).toContain("Fix manually:");
                }
                expect(result.issues.some((i) => i.code === c.code)).toBe(true);
            }
        });
    });

    it.each([
        ["word-drawing-arc-radius", "non-negative integers"],
        ["word-attached-toolbar-utf8", "restore the original toolbar bytes"],
    ])("manual byte and radius guidance for %s describes a valid correction", (code, guidance) => {
        const issue: ValidationIssue = { severity: "error", code, message: "synthetic" };
        const [plan] = buildRepairPlanIssues([issue], [issue]);
        expect(plan.message).toContain(`Fix manually:`);
        expect(plan.message).toContain(guidance);
    });

    it("reports clearance without requiring a registered guidance action", () => {
        const issue: ValidationIssue = { severity: "error", code: "unregistered", message: "synthetic" };
        expect(buildRepairPlanIssues([issue], [])).toEqual([
            {
                severity: "info",
                path: undefined,
                code: "repair-plan",
                message: "Before repair found 1 [unregistered] issue(s). Auto-repair cleared 1; 0 remain.",
            },
        ]);
    });

    it("carries manual guidance or generic wording only for remaining issues", () => {
        const before: ValidationIssue[] = [
            { severity: "error", path: "word/document.xml", code: "word-hyperlink-relationship", message: "synthetic" },
            { severity: "error", path: "word/document.xml", code: "comment-orphan-start", message: "no action" },
        ];
        const plans = buildRepairPlanIssues(before, before);
        const hyperlink = plans.find((i) => i.message.includes("[word-hyperlink-relationship]"));
        expect(hyperlink?.code).toBe("repair-plan-unavailable");
        expect(hyperlink?.message).toContain("point the w:hyperlink r:id");
        expect(hyperlink?.message).toContain("Fix manually:");
        const unguided = plans.find((i) => i.message.includes("[comment-orphan-start]"));
        expect(unguided?.code).toBe("repair-plan-unavailable");
        expect(unguided?.message).toContain("No manual repair guidance is available");
    });

    it("counts cleared and remaining per (path, code), so a surviving finding blocks 'none remain' (#648)", () => {
        const cleared: ValidationIssue = {
            severity: "error",
            path: "word/document.xml",
            code: "xsd-error",
            message: "}graphic': This element is not expected",
        };
        const survivor: ValidationIssue = { severity: "error", path: "word/document.xml", code: "xsd-error", message: "bogus" };
        const plans = buildRepairPlanIssues([cleared, survivor], [survivor]);
        // One group per (path, code): the survivor's null guidance must not let
        // the cleared member's group claim "0 remain" for [xsd-error].
        expect(plans).toHaveLength(1);
        expect(plans[0]).toMatchObject({ severity: "info", code: "repair-plan-partial", path: "word/document.xml" });
        expect(plans[0].message).toBe(
            "Before repair found 2 [xsd-error] issue(s). Auto-repair cleared 1; 1 remain. No manual repair guidance is available.",
        );
    });

    it("lists each surviving finding's guidance inside the one (path, code) line (#648)", () => {
        const before: ValidationIssue[] = [
            { severity: "error", path: "word/document.xml", code: "xsd-error", message: "}graphic': This element is not expected" },
            { severity: "error", path: "word/document.xml", code: "xsd-error", message: "}filetime': '' is not a valid value" },
            { severity: "error", path: "word/document.xml", code: "xsd-error", message: "bogus" },
        ];
        const plans = buildRepairPlanIssues(before, before);
        expect(plans).toHaveLength(1);
        expect(plans[0].code).toBe("repair-plan-unavailable");
        expect(plans[0].message).toContain("Before repair found 3 [xsd-error] issue(s)");
        expect(plans[0].message).toContain("supply the required DrawingML children in schema order");
        expect(plans[0].message).toContain("give the custom property a valid xs:dateTime value");
        expect(plans[0].message).toContain("No manual repair guidance is available for the other remaining issue(s)");
    });

    it("groups and sorts by path and code identically across permutations", () => {
        const before: ValidationIssue[] = [
            { severity: "warning", path: "word/document.xml", code: "ws-missing-preserve", message: "a" },
            { severity: "error", path: "word/document.xml", code: "ws-missing-preserve", message: "b" },
            { severity: "error", path: "word/header1.xml", code: "ws-missing-preserve", message: "c" },
            { severity: "error", path: "word/document.xml", code: "xsd-error", message: "}graphic': This element is not expected" },
            { severity: "error", path: "word/document.xml", code: "xsd-error", message: "}filetime': '' is not a valid value" },
            { severity: "warning", message: "no path, code or guidance" },
            { severity: "error", path: "word/document.xml", code: "xsd-error", message: "no guidance" },
        ];
        const after = [before[0], before[3], before[5], before[6]];
        const expected = buildRepairPlanIssues(before, after);
        // #648: groups key on (path, code) alone, so the three same-part
        // xsd-error guidance classes merge into one line.
        expect(expected).toHaveLength(4);
        expect(expected.map((i) => i.code)).toEqual([
            "repair-plan-unavailable",
            "repair-plan-partial",
            "repair-plan-partial",
            "repair-plan",
        ]);
        expect(expected[0].message).toContain("[uncoded-validation-issue]");
        expect(expected[1].message).toContain("Auto-repair cleared 1; 1 remain");
        // Two xsd-errors survive at word/document.xml: one carries the scaffold
        // guidance, the other has none, so the line lists both facts.
        expect(expected[2].message).toContain("Auto-repair cleared 1; 2 remain");
        expect(expected[2].message).toContain("supply the required DrawingML children in schema order");
        expect(expected[2].message).toContain("No manual repair guidance is available for the other remaining issue(s)");
        let permutations = 0;
        const visit = (prefix: ValidationIssue[], rest: ValidationIssue[]): void => {
            if (rest.length === 0) {
                expect(buildRepairPlanIssues(prefix, [...after].reverse())).toEqual(expected);
                permutations += 1;
                return;
            }
            rest.forEach((issue, i) =>
                visit(
                    [...prefix, issue],
                    rest.filter((_, j) => j !== i),
                ),
            );
        };
        visit([], before);
        expect(permutations).toBe(5040);
    });

    it("ignores informational findings and does not claim clearance when a group grows", () => {
        const issue: ValidationIssue = { severity: "error", code: "unregistered", message: "one" };
        const info: ValidationIssue = { ...issue, severity: "info" };
        expect(buildRepairPlanIssues([info], [])).toEqual([]);
        expect(buildRepairPlanIssues([issue], [info])[0].code).toBe("repair-plan");
        const [plan] = buildRepairPlanIssues([issue, info], [issue, issue]);
        expect(plan.code).toBe("repair-plan-unavailable");
        expect(plan.message).toContain("Auto-repair cleared 0; 2 remain");
        expect(buildRepairPlanIssues([], [issue])).toEqual([]);
    });
});

// #650 — guidance strings must name the actual fix for the defect the code reports.
describe("repair plan manual guidance wording (#650)", () => {
    it.each([
        ["comment-thread-commentid-missing-paraid", "add the schema-required w16cid:paraId attribute to the w16cid:commentId element"],
        [
            "comment-thread-commentid-missing-durableid",
            "add the schema-required w16cid:durableId attribute to the w16cid:commentId element",
        ],
        ["comment-thread-durableid-missing", "add the schema-required w16cex:durableId attribute to the w16cex:commentExtensible element"],
    ])("missing required attribute guidance for %s says to add the attribute", (code, guidance) => {
        const issue: ValidationIssue = { severity: "error", path: "word/commentsIds.xml", code, message: "synthetic" };
        const [plan] = buildRepairPlanIssues([issue], [issue]);
        expect(plan.code).toBe("repair-plan-unavailable");
        expect(plan.message).toContain(`Fix manually: ${guidance}.`);
        // The attribute is absent, so there is no id to "align".
        expect(plan.message).not.toContain("align commentsIds/commentsExtensible IDs");
    });

    it.each([["comment-thread-commentid-paraid-orphan"], ["comment-thread-durableid-orphan"]])(
        "orphan guidance for %s still names the alignment fix",
        (code) => {
            const issue: ValidationIssue = { severity: "error", code, message: "synthetic" };
            const [plan] = buildRepairPlanIssues([issue], [issue]);
            expect(plan.message).toContain(
                "align commentsIds/commentsExtensible IDs with comments.xml when the target comment thread is unambiguous",
            );
        },
    );

    it("id-durable-decimal names the decimal radix, not a range defect", () => {
        const issue: ValidationIssue = { severity: "error", path: "word/numbering.xml", code: "id-durable-decimal", message: "synthetic" };
        const [plan] = buildRepairPlanIssues([issue], [issue]);
        expect(plan.code).toBe("repair-plan-unavailable");
        expect(plan.message).toContain("Fix manually: write the numbering.xml w16cid:durableId as a decimal number.");
        expect(plan.message).not.toContain("in-range");
    });

    it("word-drawing-extent-negative says non-negative, as ST_PositiveCoordinate allows 0", () => {
        const issue: ValidationIssue = { severity: "error", code: "word-drawing-extent-negative", message: "synthetic" };
        const [plan] = buildRepairPlanIssues([issue], [issue]);
        expect(plan.code).toBe("repair-plan-unavailable");
        expect(plan.message).toContain("Fix manually: replace each negative wp:extent cx or cy with a non-negative integer.");
        expect(plan.message).not.toContain("positive integer");
    });
});

/*
 * #651 — the (issues) → (before, after) signature break and the repair-plan
 * code's meaning change are documented in MIGRATION.md rather than hidden
 * behind a one-argument overload: a silent overload would let old callers
 * read "none remain" as if the code still meant "a repair is registered".
 */
describe("buildRepairPlanIssues migration notes (#651)", () => {
    it("keeps the two-argument signature the migration guide documents", () => {
        expect(buildRepairPlanIssues.length).toBe(2);
    });

    it("documents the signature change and the repair-plan meaning change in MIGRATION.md", async () => {
        const migration = await fs.readFile(path.join(path.dirname(fileURLToPath(import.meta.url)), "..", "MIGRATION.md"), "utf-8");
        expect(migration).toContain("buildRepairPlanIssues");
        expect(migration).toContain("repair-plan");
    });

    it("throws the TypeError the migration guide names for the old one-argument call (#667)", async () => {
        const migration = await fs.readFile(path.join(path.dirname(fileURLToPath(import.meta.url)), "..", "MIGRATION.md"), "utf-8");
        // The engine's wording is not quoted; only the error class is.
        expect(migration.replace(/\s+/g, " ")).toContain("the old one-argument call throws a `TypeError`");
        const oneArgument = buildRepairPlanIssues as unknown as (issues: readonly ValidationIssue[]) => unknown;
        expect(() => oneArgument([{ severity: "error", code: "x", message: "m" }])).toThrow(TypeError);
    });
});
