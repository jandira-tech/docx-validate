import { describe, expect, it } from "vitest";
import type { DocxSemanticInventory } from "../src/scripts/office/validators/docx-diagnostics";
import {
    diffDocxInventories,
    formatInventoryDiffMarkdown,
    inventoryDiffToIssues,
} from "../src/scripts/office/validators/docx-inventory-diff";

function inv(entries: { path: string; category: string; label: string; unit: string; count: number }[]): DocxSemanticInventory {
    const counters = new Map<string, { path: string; category: string; label: string; unit: string; count: number }>();
    for (const e of entries) counters.set(`${e.path}\u0000${e.category}\u0000${e.label}\u0000${e.unit}`, e);
    return { counters };
}

describe("diffDocxInventories", () => {
    it("splits added / removed / changed / unchanged by key and direction", () => {
        const before = inv([
            { path: "word/document.xml", category: "document structure", label: "paragraph", unit: "element(s)", count: 10 },
            { path: "word/document.xml", category: "table shape", label: "table 3×4", unit: "table(s)", count: 1 },
            { path: "word/document.xml", category: "inline mark", label: "tab", unit: "occurrence(s)", count: 5 },
        ]);
        const after = inv([
            { path: "word/document.xml", category: "document structure", label: "paragraph", unit: "element(s)", count: 8 }, // changed ↓
            { path: "word/document.xml", category: "table shape", label: "table 3×2", unit: "table(s)", count: 1 }, // reshape: 3×4 removed, 3×2 added
            { path: "word/document.xml", category: "inline mark", label: "tab", unit: "occurrence(s)", count: 5 }, // unchanged
        ]);
        const d = diffDocxInventories(before, after);
        expect(d.added.map((x) => x.label)).toEqual(["table 3×2"]);
        expect(d.removed.map((x) => x.label)).toEqual(["table 3×4"]);
        expect(d.changed.map((x) => x.label)).toEqual(["paragraph"]);
        expect(d.changed[0]).toMatchObject({ before: 10, after: 8 });
        expect(d.unchangedCount).toBe(1);
    });

    it("treats identical inventories as fully unchanged", () => {
        const a = inv([{ path: "p", category: "text", label: "visible text", unit: "character(s)", count: 3 }]);
        const d = diffDocxInventories(a, a);
        expect(d.added).toHaveLength(0);
        expect(d.removed).toHaveLength(0);
        expect(d.changed).toHaveLength(0);
        expect(d.unchangedCount).toBe(1);
    });

    it("tiers diff deltas into severity-graded issues", () => {
        const before = inv([
            { path: "word/document.xml", category: "document structure", label: "paragraph", unit: "element(s)", count: 10 },
            { path: "word/document.xml", category: "formatting", label: "bold", unit: "formatted character(s)", count: 4 },
            { path: "word/document.xml", category: "table shape", label: "table 3×2", unit: "table(s)", count: 1 },
        ]);
        const after = inv([
            { path: "word/document.xml", category: "document structure", label: "paragraph", unit: "element(s)", count: 8 }, // content ↓ → error
            { path: "word/document.xml", category: "formatting", label: "bold", unit: "formatted character(s)", count: 1 }, // formatting ↓ → warn
            { path: "word/document.xml", category: "table shape", label: "table 3×4", unit: "table(s)", count: 1 }, // shape change → warn (both)
        ]);
        const issues = inventoryDiffToIssues(diffDocxInventories(before, after));
        const find = (code: string) => issues.find((i) => i.code === code);
        expect(find("inventory-content-loss")?.severity).toBe("error");
        expect(find("inventory-formatting-drift")?.severity).toBe("warning");
        expect(issues.filter((i) => i.code === "inventory-shape-change").every((i) => i.severity === "warning")).toBe(true);
        // a content GAIN is a warning, never an error
        const grew = inventoryDiffToIssues(diffDocxInventories(after, before));
        expect(grew.find((i) => i.code === "inventory-content-added")?.severity).toBe("warning");
        expect(grew.some((i) => i.severity === "error")).toBe(false); // table 3×4→3×2 + paragraph growth must not error
    });

    it("tiers a bookkeeping decrease to a warning", () => {
        const before = inv([{ path: "word/document.xml", category: "package asset", label: "part bytes", unit: "byte(s)", count: 100 }]);
        const after = inv([{ path: "word/document.xml", category: "package asset", label: "part bytes", unit: "byte(s)", count: 40 }]);
        const issues = inventoryDiffToIssues(diffDocxInventories(before, after));
        const drift = issues.find((i) => i.code === "inventory-bookkeeping-drift");
        expect(drift?.severity).toBe("warning");
    });

    it("renders a markdown report with sections, severity prefixes, and a summary", () => {
        const before = inv([
            { path: "word/document.xml", category: "document structure", label: "paragraph", unit: "element(s)", count: 10 },
            { path: "word/document.xml", category: "table shape", label: "table 3×4", unit: "table(s)", count: 1 },
        ]);
        const after = inv([
            { path: "word/document.xml", category: "document structure", label: "paragraph", unit: "element(s)", count: 8 },
            { path: "word/document.xml", category: "table shape", label: "table 3×2", unit: "table(s)", count: 1 },
        ]);
        const md = formatInventoryDiffMarkdown(diffDocxInventories(before, after));
        expect(md).toContain("## Removed");
        expect(md).toContain("## Added");
        expect(md).toContain("## Changed");
        expect(md).toContain("## Reshaped"); // shape loss+gain paired for readability
        expect(md).toMatch(/table 3×4.*→.*table 3×2/);
        expect(md).toMatch(/1 added, 1 removed, 1 changed, 0 unchanged/);
        // paragraph 10→8 is a content loss (1 error); the table reshape splits into a
        // removed (3×4) + added (3×2) shape change, each a warning → 2 warn.
        expect(md).toContain("1 error / 2 warn");
        expect(md).toMatch(/🔴|🟠|⚪/); // severity prefixes present
    });
});
