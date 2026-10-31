import { describe, expect, it, vi } from "vitest";

// Run the validator under Windows path semantics, so relPath sees the
// backslash-separated relatives path.relative returns there (#1183).
vi.mock("node:path", async () => {
    const actual = await vi.importActual<typeof import("node:path")>("node:path");
    return { ...actual.win32, default: actual.win32 };
});

const { DOCXSchemaValidator } = await import("../src/scripts/office/validators/docx");

class RelPathProbe extends DOCXSchemaValidator {
    probe(absPath: string): string {
        return this.relPath(absPath);
    }
}

describe("BaseSchemaValidator.relPath on Windows (#1183)", () => {
    it.each([
        ["C:\\pkg\\word\\_rels\\document.xml.rels", "word/_rels/document.xml.rels"],
        ["C:\\pkg\\_rels\\.rels", "_rels/.rels"],
        ["C:\\pkg\\[Content_Types].xml", "[Content_Types].xml"],
    ])("writes %s as the part name %s", (absPath, expected) => {
        expect(new RelPathProbe({ unpackedDir: "C:\\pkg" }).probe(absPath)).toBe(expected);
    });
});
