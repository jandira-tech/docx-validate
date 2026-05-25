import { describe, expect, it } from "vitest";
import { repairDocxInMemory } from "../src/scripts/office/repair-in-memory";

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';

// A minimal unpacked DOCX held entirely in memory whose body has a
// whitespace-bearing <w:t> missing xml:space="preserve" — a real defect the
// repair suite (repairWhitespacePreservation) fixes. No disk, no temp dir.
function minimalParts(): Array<[string, string]> {
    return [
        [
            "[Content_Types].xml",
            `<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
                `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
                `<Default Extension="xml" ContentType="application/xml"/>` +
                `<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>` +
                `</Types>`,
        ],
        [
            "_rels/.rels",
            `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
                `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>` +
                `</Relationships>`,
        ],
        [
            "word/document.xml",
            `<?xml version="1.0"?><w:document ${W}><w:body>` +
                `<w:p><w:r><w:t> leading and trailing </w:t></w:r></w:p>` +
                `</w:body></w:document>`,
        ],
    ];
}

describe("repairDocxInMemory", () => {
    it("repairs whitespace preservation purely in memory (no disk)", async () => {
        const { parts, repairs } = await repairDocxInMemory(minimalParts());
        expect(repairs).toBeGreaterThan(0);

        const doc = parts.find(([rel]) => rel === "word/document.xml");
        expect(doc).toBeDefined();
        const text = (doc as [string, Buffer])[1].toString("utf-8");
        // the whitespace-bearing run now carries xml:space="preserve"
        expect(text).toMatch(/<w:t[^>]*xml:space="preserve"[^>]*> leading and trailing <\/w:t>/);
    });

    it("returns the full part set so it can be repacked", async () => {
        const { parts } = await repairDocxInMemory(minimalParts());
        const names = parts.map(([rel]) => rel).sort();
        expect(names).toContain("[Content_Types].xml");
        expect(names).toContain("_rels/.rels");
        expect(names).toContain("word/document.xml");
    });
});
