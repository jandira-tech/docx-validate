/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

/**
 * Package-level content-type resolution (docs/plans/2026-10-23-fixing-validator.md,
 * Task 4; issues #208, #237, #249).
 *
 * Builds tiny DOCX packages in memory from MINIMAL_PARTS, adds media parts
 * plus [Content_Types].xml declarations, and drives the public `validate()`
 * entry point under every profile. Resolver-level behaviour is unit-tested in
 * `tests/content-types.test.ts`.
 */

import { writeFileSync } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

import { describe, expect, it } from "vitest";

import { withTempDir } from "../src/lib/run-cli";
import type { Profile } from "../src/lib/types";
import { validate } from "../src/scripts/office/validate";
import { buildDocx, MINIMAL_PARTS } from "./support/minimal-docx";

const HERE = path.dirname(fileURLToPath(import.meta.url));
const LANE = path.join(HERE, "lane", "jubarte-0.13.0");
const PROFILES: readonly Profile[] = ["lenient", "strict", "word-valid"];

const CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types";
const RELS_DEFAULT = `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>`;
const XML_DEFAULT = `<Default Extension="xml" ContentType="application/xml"/>`;
const DOC_OVERRIDE =
    `<Override PartName="/word/document.xml"` +
    ` ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>`;
const STYLES_OVERRIDE =
    `<Override PartName="/word/styles.xml"` + ` ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>`;

/** [Content_Types].xml for the minimal package plus extra declarations. */
const contentTypes = (extra: string): string =>
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<Types xmlns="${CT_NS}">${RELS_DEFAULT}${XML_DEFAULT}${DOC_OVERRIDE}${STYLES_OVERRIDE}${extra}</Types>`;

const JPEG_BYTES = new Uint8Array([0xff, 0xd8, 0xff, 0xe0, 0x00, 0x10]);

type Parts = Record<string, string | Uint8Array>;

const withMedia = (extraDeclarations: string, media: Parts): Parts => ({
    ...MINIMAL_PARTS,
    "[Content_Types].xml": contentTypes(extraDeclarations),
    ...media,
});

const ctIssues = (issues: readonly { code?: string }[]): { code?: string }[] =>
    issues.filter((issue) => (issue.code ?? "").startsWith("ct-"));

async function validateBytes(bytes: Uint8Array, profile: Profile) {
    return withTempDir(async (dir) => {
        const file = path.join(dir, "specimen.docx");
        writeFileSync(file, bytes);
        return validate(file, { profile });
    });
}

describe("content-type resolution across profiles", { timeout: 120_000 }, () => {
    for (const profile of PROFILES) {
        it(`${profile}: media covered only by an exact Override yields neither ct-undeclared-ext nor ct-uncovered-part`, async () => {
            // No <Default Extension="jpeg"> — the ONLY declaration is the
            // exact Override, which the old extension-only media check never
            // consulted (issues #208/#237: jubarte redlines declare every
            // media part this way).
            const bytes = await buildDocx(
                withMedia(`<Override PartName="/word/media/image1.jpeg" ContentType="image/jpeg"/>`, {
                    "word/media/image1.jpeg": JPEG_BYTES,
                }),
            );
            const result = await validateBytes(bytes, profile);
            expect(ctIssues(result.issues)).toEqual([]);
        });
    }

    it("resolves OPC-equivalent case variants of an Override PartName (lenient profile, direct validator)", async () => {
        // Part-name equivalence is ASCII case-insensitive: /Word/Media/IMAGE1.BIN
        // covers word/media/image1.bin. The .bin extension is NOT known media,
        // so the residual ct-uncovered-part pass owns this case today — with a
        // case-sensitive Override lookup it false-positives.
        const bytes = await buildDocx(
            withMedia(`<Override PartName="/Word/Media/IMAGE1.BIN" ContentType="application/octet-stream"/>`, {
                "word/media/image1.bin": new Uint8Array([0x00, 0x01]),
            }),
        );
        const result = await validateBytes(bytes, "lenient");
        expect(ctIssues(result.issues)).toEqual([]);
    });

    it("resolves an escaped Override PartName against the decoded package path", async () => {
        // PartName is a URI; the on-disk path is its decoded form (%20 → space).
        const bytes = await buildDocx(
            withMedia(`<Override PartName="/word/media/my%20pic.jpeg" ContentType="image/jpeg"/>`, {
                "word/media/my pic.jpeg": JPEG_BYTES,
            }),
        );
        const result = await validateBytes(bytes, "lenient");
        expect(ctIssues(result.issues)).toEqual([]);
    });

    it("resolves a case-variant Override for a declarable XML root without ct-undeclared-part", async () => {
        // A second word/document.xml-like part whose Override differs only by
        // ASCII case must satisfy the XML-root coverage check.
        const parts: Parts = {
            ...MINIMAL_PARTS,
            "word/report.xml": MINIMAL_PARTS["word/document.xml"]!,
        };
        parts["[Content_Types].xml"] = contentTypes(
            `<Override PartName="/Word/Report.XML" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>`,
        );
        const bytes = await buildDocx(parts);
        const result = await validateBytes(bytes, "lenient");
        expect(result.issues.some((i) => i.code === "ct-undeclared-part" && i.path === "word/report.xml")).toBe(false);
        expect(result.issues.some((i) => i.code === "ct-uncovered-part" && i.path === "word/report.xml")).toBe(false);
    });

    it("matches Default extensions ASCII case-insensitively for media parts", async () => {
        // <Default Extension="JPEG"> covers x.jpeg — extension comparison is
        // ASCII case-insensitive in both directions.
        const bytes = await buildDocx(
            withMedia(`<Default Extension="JPEG" ContentType="image/jpeg"/>`, {
                "word/media/upper default.jpeg": JPEG_BYTES,
            }),
        );
        const result = await validateBytes(bytes, "lenient");
        expect(ctIssues(result.issues)).toEqual([]);
    });
});

/**
 * The scenarios below reference the media part from word/_rels/document.xml.rels
 * so `rels-unreferenced` noise stays out and `valid` can be asserted.
 */
const IMAGE_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";

const relsWithImage = (target: string): string =>
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
    `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>` +
    `<Relationship Id="rId2" Type="${IMAGE_REL_TYPE}" Target="${target}"/>` +
    `</Relationships>`;

interface MediaCase {
    /** ZIP entry name of the media part — the on-disk part path. */
    entry: string;
    /** Override PartName exactly as written into [Content_Types].xml; omit for none. */
    override?: string;
    /** Extra Default Extension to declare; omit for none. */
    defaultExtension?: string;
}

const mediaPackage = (c: MediaCase): Parts => {
    const declarations =
        (c.override ? `<Override PartName="${c.override}" ContentType="image/jpeg"/>` : "") +
        (c.defaultExtension ? `<Default Extension="${c.defaultExtension}" ContentType="image/jpeg"/>` : "");
    return {
        ...MINIMAL_PARTS,
        "[Content_Types].xml": contentTypes(declarations),
        "word/_rels/document.xml.rels": relsWithImage(c.entry.replace(/^word\//, "")),
        [c.entry]: JPEG_BYTES,
    };
};

describe("content-type resolution: OPC part-name equivalence", { timeout: 120_000 }, () => {
    for (const profile of PROFILES) {
        it(`${profile}: a case-variant Override PartName covers the media part in every pass`, async () => {
            // OPC part-name comparison is ASCII case-insensitive, so
            // /Word/Media/IMAGE1.JPEG declares word/media/image1.jpeg.
            const bytes = await buildDocx(mediaPackage({ entry: "word/media/image1.jpeg", override: "/Word/Media/IMAGE1.JPEG" }));
            const result = await validateBytes(bytes, profile);
            expect(ctIssues(result.issues)).toEqual([]);
            expect(result.valid).toBe(true);
        });

        it(`${profile}: an escaped Override PartName covers the decoded on-disk media part`, async () => {
            // The PartName is a URI (%20); the ZIP entry carries the decoded
            // space. Resolution must decode before comparing.
            const bytes = await buildDocx(mediaPackage({ entry: "word/media/my image.jpeg", override: "/word/media/my%20image.jpeg" }));
            const result = await validateBytes(bytes, profile);
            expect(ctIssues(result.issues)).toEqual([]);
            expect(result.valid).toBe(true);
        });

        it(`${profile}: a genuinely undeclared known media part still fails as ct-undeclared-ext with the fix-it message`, async () => {
            // No Override, no Default for jpeg — the defect the check exists
            // for, with its actionable suggestion (kept verbatim).
            const bytes = await buildDocx(mediaPackage({ entry: "word/media/image1.jpeg" }));
            const result = await validateBytes(bytes, profile);
            const hit = result.issues.find((i) => i.code === "ct-undeclared-ext");
            expect(hit?.path).toBe("word/media/image1.jpeg");
            expect(hit?.message).toContain(`should add: <Default Extension="jpeg" ContentType="image/jpeg"/>`);
            // ct-undeclared-ext owns known media extensions; the residual
            // pass must not double-flag them.
            expect(result.issues.some((i) => i.code === "ct-uncovered-part")).toBe(false);
        });

        it(`${profile}: a genuinely undeclared non-media part still fails as ct-uncovered-part`, async () => {
            // .bin is not a known media extension and nothing declares it.
            const bytes = await buildDocx(mediaPackage({ entry: "word/media/data.bin" }));
            const result = await validateBytes(bytes, profile);
            const hit = result.issues.find((i) => i.code === "ct-uncovered-part");
            expect(hit?.path).toBe("word/media/data.bin");
            expect(result.issues.some((i) => i.code === "ct-undeclared-ext")).toBe(false);
        });
    }
});

/**
 * Lane regression (issues #208, #237, #249): the jubarte 0.13.0 redline
 * specimens declare every media part by exact `<Override>` — jpeg/png
 * (k01) and emf plus mis-extensioned .bin images (k02). They must report
 * neither `ct-undeclared-ext` nor `ct-uncovered-part` under strict. The
 * specimens are large, so these assert only those two codes.
 */
describe("lane regression: jubarte 0.13.0 specimens", { timeout: 180_000 }, () => {
    for (const name of ["k01_1", "k01_2", "k02_1", "k02_2"]) {
        it(`${name}.redline.docx reports no ct-undeclared-ext and no ct-uncovered-part under strict`, async () => {
            const result = await validate(path.join(LANE, `${name}.redline.docx`), { profile: "strict" });
            expect(result.issues.filter((i) => i.code === "ct-undeclared-ext" || i.code === "ct-uncovered-part")).toEqual([]);
        });
    }
});
