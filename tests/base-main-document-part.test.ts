/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

/**
 * `BaseSchemaValidator.isMainDocumentPart` runs once per part whose root
 * markup compatibility ignores. It must not re-walk the unpacked tree or
 * re-read `_rels/.rels` on every call, or the cost grows with the square of
 * the part count (#355).
 */
import * as nodeFs from "node:fs";
import path from "node:path";

import { afterEach, describe, expect, it, vi } from "vitest";

import { withTempDir } from "../src/lib/run-cli";
import { BaseSchemaValidator } from "../src/scripts/office/validators/base";
import { MINIMAL_PARTS } from "./support/minimal-docx";

vi.mock("node:fs", async (importOriginal) => {
    const actual = await importOriginal<typeof import("node:fs")>();
    return { ...actual, default: actual, readdirSync: vi.fn(actual.readdirSync), readFileSync: vi.fn(actual.readFileSync) };
});

interface MainProbe {
    isMainDocumentPart(file: string): boolean;
}

const writeParts = (dir: string, parts: Readonly<Record<string, string>>): void => {
    for (const [name, content] of Object.entries(parts)) {
        nodeFs.mkdirSync(path.dirname(path.join(dir, name)), { recursive: true });
        nodeFs.writeFileSync(path.join(dir, name), content);
    }
};

const relsReads = (): number =>
    vi.mocked(nodeFs.readFileSync).mock.calls.filter(([file]) => String(file).endsWith(path.join("_rels", ".rels"))).length;

afterEach(() => {
    vi.mocked(nodeFs.readdirSync).mockClear();
    vi.mocked(nodeFs.readFileSync).mockClear();
});

describe("isMainDocumentPart (#355, #357, #362)", () => {
    it("walks the tree and reads the package relationships once per validator", async () => {
        await withTempDir(async (dir) => {
            const extra = Object.fromEntries(Array.from({ length: 5 }, (_, i) => [`word/garbage${i}.xml`, "<x/>"]));
            writeParts(dir, { ...MINIMAL_PARTS, ...extra });
            const validator = new BaseSchemaValidator({ unpackedDir: dir, verbose: false, profile: "strict" });
            const probe = validator as unknown as MainProbe;
            vi.mocked(nodeFs.readdirSync).mockClear();
            vi.mocked(nodeFs.readFileSync).mockClear();

            expect(probe.isMainDocumentPart(path.join(dir, "word", "document.xml"))).toBe(true);
            const walked = vi.mocked(nodeFs.readdirSync).mock.calls.length;
            for (const name of Object.keys(extra)) {
                expect(probe.isMainDocumentPart(path.join(dir, name))).toBe(false);
            }
            expect(probe.isMainDocumentPart(path.join(dir, "word", "document.xml"))).toBe(true);

            expect(vi.mocked(nodeFs.readdirSync).mock.calls.length).toBe(walked);
            expect(relsReads()).toBe(1);
        });
    });

    // A file name with a literal backslash cannot exist on Windows.
    it.skipIf(process.platform === "win32")(
        "calls a main part stored under a literal backslash name main, as the package view does (#357)",
        async () => {
            await withTempDir(async (dir) => {
                const relationships = MINIMAL_PARTS["_rels/.rels"].replace('Target="word/document.xml"', 'Target="word/a\\b.xml"');
                writeParts(dir, { "_rels/.rels": relationships, "word/a\\b.xml": "<x/>" });
                const validator = new BaseSchemaValidator({ unpackedDir: dir, verbose: false, profile: "strict" });
                const probe = validator as unknown as MainProbe;

                expect(nodeFs.readdirSync(path.join(dir, "word"))).toEqual(["a\\b.xml"]);
                expect(probe.isMainDocumentPart(path.join(dir, "word", "a\\b.xml"))).toBe(true);
            });
        },
    );

    it("reads a UTF-16 _rels/.rels as the package view does (#362)", async () => {
        await withTempDir(async (dir) => {
            writeParts(dir, MINIMAL_PARTS);
            nodeFs.writeFileSync(
                path.join(dir, "_rels", ".rels"),
                Buffer.from(`\ufeff${MINIMAL_PARTS["_rels/.rels"].replace('encoding="UTF-8"', 'encoding="UTF-16"')}`, "utf16le"),
            );
            const probe = new BaseSchemaValidator({ unpackedDir: dir, verbose: false, profile: "strict" }) as unknown as MainProbe;

            expect(probe.isMainDocumentPart(path.join(dir, "word", "document.xml"))).toBe(true);
        });
    });

    it.skipIf(process.platform === "win32")(
        "finds package relationships stored as _rels\\.rels, as the package view does (#362)",
        async () => {
            await withTempDir(async (dir) => {
                const { "_rels/.rels": relationships, ...rest } = MINIMAL_PARTS;
                writeParts(dir, { ...rest, "_rels\\.rels": relationships });
                const probe = new BaseSchemaValidator({ unpackedDir: dir, verbose: false, profile: "strict" }) as unknown as MainProbe;

                expect(nodeFs.existsSync(path.join(dir, "_rels"))).toBe(false);
                expect(probe.isMainDocumentPart(path.join(dir, "word", "document.xml"))).toBe(true);
            });
        },
    );

    it("forgets the main parts when a repair deletes the package relationships (#362)", async () => {
        await withTempDir(async (dir) => {
            writeParts(dir, MINIMAL_PARTS);
            const validator = new BaseSchemaValidator({ unpackedDir: dir, verbose: false, profile: "strict" });
            const probe = validator as unknown as MainProbe;
            const main = path.join(dir, "word", "document.xml");
            expect(probe.isMainDocumentPart(main)).toBe(true);

            nodeFs.writeFileSync(
                path.join(dir, "_rels", ".rels"),
                '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>',
            );
            expect(await validator.repairEmptyRelsParts()).toBeGreaterThan(0);
            expect(nodeFs.existsSync(path.join(dir, "_rels", ".rels"))).toBe(false);

            expect(probe.isMainDocumentPart(main)).toBe(false);
        });
    });
});
