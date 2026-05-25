import { mkdtempSync, rmSync } from "node:fs";
import { tmpdir } from "node:os";
import path from "node:path";
import { describe, expect, it } from "vitest";
import { DiskPartFS, MemoryPartFS } from "../src/lib/part-fs";

describe("MemoryPartFS", () => {
    it("lists, reads, writes and removes parts via absolute paths rooted at root", async () => {
        const fs = new MemoryPartFS([
            ["word/document.xml", "<w:document/>"],
            ["word/_rels/document.xml.rels", "<Relationships/>"],
            ["[Content_Types].xml", "<Types/>"],
        ]);
        const docAbs = path.join(fs.root, "word/document.xml");
        expect(fs.exists(docAbs)).toBe(true);
        expect(await fs.readText(docAbs)).toBe("<w:document/>");

        // list with extension filter returns absolute paths
        const xmls = fs.list([".xml"]);
        expect(xmls).toContain(docAbs);
        expect(fs.list([".rels"]).length).toBe(1);

        // path.relative(root, abs) is the part name — the arithmetic validators use
        expect(path.relative(fs.root, docAbs)).toBe(path.join("word", "document.xml"));

        await fs.write(docAbs, "<w:document><w:body/></w:document>");
        expect(await fs.readText(docAbs)).toContain("<w:body/>");

        await fs.remove(docAbs);
        expect(fs.exists(docAbs)).toBe(false);
    });

    it("entries() snapshots parts for repacking", () => {
        const fs = new MemoryPartFS([["a.xml", "1"], ["b/c.xml", "2"]]);
        const map = new Map(fs.entries().map(([k, v]) => [k, v.toString("utf-8")]));
        expect(map.get("a.xml")).toBe("1");
        expect(map.get("b/c.xml")).toBe("2");
    });

    it("missing reads throw ENOENT", async () => {
        const fs = new MemoryPartFS();
        await expect(fs.readText(path.join(fs.root, "nope.xml"))).rejects.toMatchObject({ code: "ENOENT" });
    });
});

describe("DiskPartFS", () => {
    it("mirrors fs behavior for list/read/write/remove", async () => {
        const dir = mkdtempSync(path.join(tmpdir(), "partfs-"));
        try {
            const fs = new DiskPartFS(dir);
            const abs = path.join(dir, "word", "document.xml");
            await fs.write(abs, "<w:document/>"); // auto-creates word/
            expect(fs.exists(abs)).toBe(true);
            expect(await fs.readText(abs)).toBe("<w:document/>");
            expect(fs.list([".xml"])).toContain(abs);
            await fs.remove(abs);
            expect(fs.exists(abs)).toBe(false);
        } finally {
            rmSync(dir, { force: true, recursive: true });
        }
    });
});
