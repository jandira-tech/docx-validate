/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

// Integration specimens for Word sessions 89 (s89a) and 90 (s90a), #1186.
// Each observed row ran twice, with all four readings agreeing.
import { promises as fs } from "node:fs";
import path from "node:path";

import { describe, expect, it } from "vitest";

import { withTempDir } from "../src/lib/run-cli";
import type { Profile } from "../src/lib/types";
import { ValidationContext } from "../src/lib/validation-context";
import { createDefaultCheckRegistry, WORD_RULE_CHECKS } from "../src/scripts/office/checks";
import { validate } from "../src/scripts/office/validate";
import { buildDocx, MINIMAL_PARTS } from "./support/minimal-docx";

const CODE = "word-package-thumbnail-type";
const CT = "[Content_Types].xml";
const RELS = "_rels/.rels";
const THUMBNAIL = "http://schemas.openxmlformats.org/package/2006/relationships/metadata/thumbnail";
const PNG_PART = "docProps/thumbnail.png";
const PNG = Buffer.from("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg==", "base64");
// A 96 x 96 JPEG calibration, embedded so tests need no image tools.
const JPEG = Buffer.from(
    "/9j/4AAQSkZJRgABAQAASABIAAD/4QBMRXhpZgAATU0AKgAAAAgAAYdpAAQAAAABAAAAGgAAAAAAA6ABAAMAAAABAAEAAKACAAQAAAABAAAAYKADAAQAAAABAAAAYAAAAAD/7QA4UGhvdG9zaG9wIDMuMAA4QklNBAQAAAAAAAA4QklNBCUAAAAAABDUHYzZjwCyBOmACZjs+EJ+/8AAEQgAYABgAwEiAAIRAQMRAf/EAB8AAAEFAQEBAQEBAAAAAAAAAAABAgMEBQYHCAkKC//EALUQAAIBAwMCBAMFBQQEAAABfQECAwAEEQUSITFBBhNRYQcicRQygZGhCCNCscEVUtHwJDNicoIJChYXGBkaJSYnKCkqNDU2Nzg5OkNERUZHSElKU1RVVldYWVpjZGVmZ2hpanN0dXZ3eHl6g4SFhoeIiYqSk5SVlpeYmZqio6Slpqeoqaqys7S1tre4ubrCw8TFxsfIycrS09TV1tfY2drh4uPk5ebn6Onq8fLz9PX29/j5+v/EAB8BAAMBAQEBAQEBAQEAAAAAAAABAgMEBQYHCAkKC//EALURAAIBAgQEAwQHBQQEAAECdwABAgMRBAUhMQYSQVEHYXETIjKBCBRCkaGxwQkjM1LwFWJy0QoWJDThJfEXGBkaJicoKSo1Njc4OTpDREVGR0hJSlNUVVZXWFlaY2RlZmdoaWpzdHV2d3h5eoKDhIWGh4iJipKTlJWWl5iZmqKjpKWmp6ipqrKztLW2t7i5usLDxMXGx8jJytLT1NXW19jZ2uLj5OXm5+jp6vLz9PX29/j5+v/bAEMAAgICAgICAwICAwUDAwMFBgUFBQUGCAYGBgYGCAoICAgICAgKCgoKCgoKCgwMDAwMDA4ODg4ODw8PDw8PDw8PDw8PD//bAEMBAgICBAQEBwQEBxALCQsQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEP/dAAQABv/aAAwDAQACEQMRAD8A7iiiiv8AQA+XCiiigAooooAKKKKACiiigAooooA//9DuKKKK/wBAD5cKKKKACiiigAooooAKKKKACiiigD//0e4ooor/AEAPlwooooAKKKKACiiigAooooAKKKKAP//S7iiiiv8AQA+XCiiigAooooAKKKKACiiigAooooA//9PuKKKK/wBAD5cKKKKACiiigAooooAKKKKACiiigD//1O4ooor/AEAPlwooooAKKKKACiiigAooooAKKKKAP//Z",
    "base64",
);

type Parts = Record<string, string | Uint8Array>;
const relationship = (target = PNG_PART, extra = "", id = "rIdThumbnail") =>
    `<Relationship Id="${id}" Type="${THUMBNAIL}" Target="${target}"${extra}/>`;
const specimen = (declarations = '<Default Extension="png" ContentType="image/png"/>', rel = relationship()): Parts => ({
    ...MINIMAL_PARTS,
    [CT]: MINIMAL_PARTS[CT]!.replace("</Types>", `${declarations}</Types>`),
    [RELS]: MINIMAL_PARTS[RELS]!.replace("</Relationships>", `${rel}</Relationships>`),
    [PNG_PART]: PNG,
});
const run = async (parts: Parts, profile: Profile = "word-valid", checks?: readonly string[]) =>
    withTempDir(async (dir) => {
        const file = path.join(dir, "thumbnail.docx");
        await fs.writeFile(file, await buildDocx(parts));
        const result = await validate(file, { profile, ...(checks ? { checks: [...checks] } : {}) });
        return { ...result, findings: result.issues.filter((issue) => issue.code === CODE) };
    });

describe("word-package-thumbnail-type (#1186)", () => {
    it("registers a DOCX content-types rule gated by XML syntax", () => {
        const definition = createDefaultCheckRegistry()
            .list()
            .find((check) => check.id === CODE);
        expect(definition).toMatchObject({ groups: ["content-types"], formats: ["docx"], requires: ["xml-syntax"] });
        expect(WORD_RULE_CHECKS).toContain(definition);
    });

    for (const profile of ["word-valid", "strict"] as const) {
        it(`${profile}: blocks a 1 x 1 PNG package thumbnail typed by Default`, async () => {
            const result = await run(specimen(), profile);
            expect(result.findings).toEqual([
                expect.objectContaining({ severity: "error", path: PNG_PART, message: expect.stringContaining(PNG_PART) }),
            ]);
            expect(result.findings[0]!.message).toMatch(/Word 16\.115 offered recovery for a PNG package thumbnail/);
            expect(result.findings[0]!.message).toMatch(/opened a JPEG one clean/);
            expect(result.findings[0]!.message).toContain("sessions 89 and 90, #1186");
            expect(result.valid).toBe(false);
        });

        it(`${profile}: blocks an Override-only PNG`, async () => {
            const result = await run(specimen(`<Override PartName="/${PNG_PART}" ContentType="image/png"/>`), profile);
            expect(result.findings.map((issue) => issue.severity)).toEqual(["error"]);
            expect(result.valid).toBe(false);
        });

        for (const source of ["Default", "Override"] as const) {
            it(`${profile}: accepts a JPEG package thumbnail typed by ${source}`, async () => {
                const part = "docProps/thumbnail.jpeg";
                const declaration =
                    source === "Default"
                        ? '<Default Extension="jpeg" ContentType="image/jpeg"/>'
                        : `<Override PartName="/${part}" ContentType="image/jpeg"/>`;
                const parts = specimen(declaration, relationship(part));
                delete parts[PNG_PART];
                parts[part] = JPEG;
                const result = await run(parts, profile);
                expect(result.findings).toEqual([]);
                expect(result.valid).toBe(true);
            });
        }
    }

    it("accepts a stored PNG no relationship names", async () => {
        const result = await run(specimen(undefined, ""));
        expect(result.findings).toEqual([]);
        expect(result.valid).toBe(true);
    });

    it("does not judge an External thumbnail even beside a stored PNG", async () => {
        expect((await run(specimen(undefined, relationship(PNG_PART, ' TargetMode="External"')))).findings).toEqual([]);
    });

    it("does not run under lenient", async () => {
        const result = await run(specimen(), "lenient");
        expect(result.findings).toEqual([]);
        expect(result.metadata?.executedChecks).not.toContain(CODE);
        expect(result.valid).toBe(true);
    });

    it("blocks when selected by id, with its prerequisite", async () => {
        const result = await run(specimen(), "word-valid", [CODE]);
        expect(result.metadata?.executedChecks).toEqual(["xml-syntax", CODE]);
        expect(result.findings.map((issue) => issue.severity)).toEqual(["error"]);
        expect(result.valid).toBe(false);
    });

    it("reports each internal thumbnail relationship, including repeated targets", async () => {
        const result = await run(specimen(undefined, relationship() + relationship(PNG_PART, "", "rIdOther")));
        expect(result.findings.map((issue) => [issue.path, issue.severity])).toEqual([
            [PNG_PART, "error"],
            [PNG_PART, "error"],
        ]);
        expect(result.valid).toBe(false);
    });

    it("compares MIME type/subtype case-insensitively", async () => {
        const result = await run(specimen('<Default Extension="png" ContentType="IMAGE/PNG"/>'));
        expect(result.findings.map((issue) => issue.severity)).toEqual(["error"]);
        expect(result.valid).toBe(false);
    });

    it("leaves parameterized PNG types unjudged", async () => {
        const result = await run(specimen('<Default Extension="png" ContentType="image/png; charset=binary"/>'));
        expect(result.findings).toEqual([]);
    });

    it("lets a JPEG Override take precedence over a PNG Default", async () => {
        const result = await run(
            specimen(`<Default Extension="png" ContentType="image/png"/><Override PartName="/${PNG_PART}" ContentType="image/jpeg"/>`),
        );
        expect(result.findings).toEqual([]);
        expect(result.valid).toBe(true);
    });

    it("lets a PNG Override take precedence over a JPEG Default", async () => {
        const result = await run(
            specimen(`<Default Extension="png" ContentType="image/jpeg"/><Override PartName="/${PNG_PART}" ContentType="image/png"/>`),
        );
        expect(result.findings.map((issue) => issue.severity)).toEqual(["error"]);
        expect(result.valid).toBe(false);
    });

    it("does not judge an absent target", async () => {
        const parts = specimen();
        delete parts[PNG_PART];
        expect((await run(parts)).findings).toEqual([]);
    });

    it("does not judge a thumbnail from the main document's sidecar", async () => {
        const parts = specimen(undefined, "");
        parts["word/_rels/document.xml.rels"] = MINIMAL_PARTS["word/_rels/document.xml.rels"]!.replace(
            "</Relationships>",
            `${relationship("../docProps/thumbnail.png")}</Relationships>`,
        );
        expect((await run(parts)).findings).toEqual([]);
    });

    it("finds OPC-equivalent package relationships, targets and Overrides", async () => {
        const parts = specimen(
            '<Override PartName="/DOCPROPS/THUMBNAIL.PNG" ContentType="image/png"/>',
            relationship("DOCPROPS/%74humbnail.png"),
        );
        parts["_rels/.RELS"] = parts[RELS]!;
        delete parts[RELS];
        const result = await run(parts);
        expect(result.findings.map((issue) => [issue.path, issue.severity])).toEqual([[PNG_PART, "error"]]);
        expect(result.valid).toBe(false);
    });

    it("keeps the error but drops the session citation in an ISO Strict package", async () => {
        const parts = specimen();
        parts["word/document.xml"] = MINIMAL_PARTS["word/document.xml"]!.replace(
            "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
            "http://purl.oclc.org/ooxml/wordprocessingml/main",
        );
        const result = await run(parts);
        expect(result.findings.map((issue) => issue.severity)).toEqual(["error"]);
        expect(result.findings[0]!.message).not.toMatch(/Word 16\.115|sessions 89|sessions 90/);
        expect(result.valid).toBe(false);
    });

    // In-memory unit rows exercise unreadable prerequisites without disk or
    // parser mocks; full-run XML diagnostics belong to the syntax check.
    it.each([
        ["missing content-types stream", CT, undefined],
        ["malformed content-types stream", CT, "<Types>"],
        ["missing package relationships", RELS, undefined],
        ["malformed package relationships", RELS, "<Relationships>"],
        ["uncovered thumbnail", CT, MINIMAL_PARTS[CT]!],
        ["empty target", RELS, MINIMAL_PARTS[RELS]!.replace("</Relationships>", `${relationship("")}</Relationships>`)],
        ["folder target", RELS, MINIMAL_PARTS[RELS]!.replace("</Relationships>", `${relationship(`${PNG_PART}/.`)}</Relationships>`)],
        [
            "foreign relationship namespace",
            RELS,
            MINIMAL_PARTS[RELS]!.replace("</Relationships>", `${relationship(PNG_PART, ' xmlns="urn:foreign"')}</Relationships>`),
        ],
    ])("leaves %s to its owning check", async (_row, part, value) => {
        const parts = specimen();
        if (value === undefined) delete parts[part];
        else parts[part] = value;
        const definition = createDefaultCheckRegistry()
            .list()
            .find((check) => check.id === CODE)!;
        expect(await definition.run(ValidationContext.fromEntries(parts))).toEqual([]);
    });

    it("reads only the thumbnail relationship selected by markup compatibility", async () => {
        const parts = specimen(undefined, "");
        const choice = `<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x="urn:unrecognized"><mc:Choice Requires="x">${relationship()}</mc:Choice><mc:Fallback/></mc:AlternateContent>`;
        parts[RELS] = MINIMAL_PARTS[RELS]!.replace("</Relationships>", `${choice}</Relationships>`);
        expect((await run(parts)).findings).toEqual([]);
        parts[RELS] = MINIMAL_PARTS[RELS]!.replace(
            "</Relationships>",
            `${choice.replace(`<mc:Fallback/>`, `<mc:Fallback>${relationship()}</mc:Fallback>`)}</Relationships>`,
        );
        const result = await run(parts);
        expect(result.findings.map((issue) => issue.severity)).toEqual(["error"]);
        expect(result.valid).toBe(false);
    });
});
