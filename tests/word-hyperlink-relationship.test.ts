/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

import { writeFileSync } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

import { describe, expect, it } from "vitest";

import { withTempDir } from "../src/lib/run-cli";
import { validate } from "../src/scripts/office/validate";
import { buildDocx, MINIMAL_PARTS } from "./support/minimal-docx";

const HERE = path.dirname(fileURLToPath(import.meta.url));
const BROKEN = path.join(HERE, "fixtures", "broken", "package.word-hyperlink-relationship.docx");
const CLEAN = path.join(HERE, "fixtures", "working", "package.word-hyperlink-relationship-repaired.docx");

describe("word-valid hyperlink relationship", () => {
    it("rejects a w:hyperlink whose r:id is not an external hyperlink relationship", async () => {
        const result = await validate(BROKEN, { profile: "word-valid" });
        expect(result.valid).toBe(false);
        expect(result.issues.some((issue) => issue.severity === "error" && issue.code === "word-hyperlink-relationship")).toBe(true);
    });

    it("does not reject the same hyperlink once that relationship is external", async () => {
        const result = await validate(CLEAN, { profile: "word-valid" });
        expect(result.issues.some((issue) => issue.code === "word-hyperlink-relationship")).toBe(false);
        expect(result.valid).toBe(true);
    });

    it("does not report the cause on lenient", async () => {
        const result = await validate(BROKEN, { profile: "lenient" });
        expect(result.issues.some((issue) => issue.code === "word-hyperlink-relationship")).toBe(false);
    });

    // strict includes every Word-valid check (validation contract, Task 1).
    it("reports the cause on strict", async () => {
        const result = await validate(BROKEN, { profile: "strict" });
        expect(result.issues.some((issue) => issue.severity === "error" && issue.code === "word-hyperlink-relationship")).toBe(true);
    });
});

/**
 * #1106, Word session 82: Word judges a hyperlink's relationship only in the
 * parts it reads. Each placement holds session 28's refused shape, a
 * `w:hyperlink` whose relationship has the Type `…/relationships/HYPERLINK`,
 * TargetMode External; its calibration uses the exact Type. Word offered
 * recovery for the refused shape in every placement marked refused and
 * opened the others clean (s82a–s82c, two runs each, all readings agree);
 * every calibration opened clean.
 */
describe("word-hyperlink-relationship judges the parts Word reads (#1106)", () => {
    const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    const R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    const PR = "http://schemas.openxmlformats.org/package/2006/relationships";
    const WML = "application/vnd.openxmlformats-officedocument.wordprocessingml";
    const CT = "[Content_Types].xml";
    const SIDECAR = "word/_rels/document.xml.rels";
    const DECL = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n`;
    const P = MINIMAL_PARTS;
    const MAIN = P["word/document.xml"]!;

    const once = (text: string, from: string, to: string): string => {
        if (text.split(from).length !== 2) throw new Error(`expected exactly one ${from}`);
        return text.replace(from, to);
    };
    const relationship = (id: string, segment: string, target: string) =>
        `<Relationship Id="${id}" Type="${R}/${segment}" Target="${target}"/>`;
    const addRelationship = (rels: string, extra: string) => once(rels, "</Relationships>", `${extra}</Relationships>`);
    const relsPart = (body: string) => `${DECL}<Relationships xmlns="${PR}">${body}</Relationships>`;
    const override = (part: string, type: string) =>
        once(P[CT]!, "</Types>", `<Override PartName="/${part}" ContentType="${type}"/></Types>`);
    const withR = (doc: string) => once(doc, `<w:document xmlns:w="${W}"`, `<w:document xmlns:w="${W}" xmlns:r="${R}"`);
    const link = (segment: string) =>
        `<Relationship Id="rIdL" Type="${R}/${segment}" Target="https://example.com/" TargetMode="External"/>`;
    const LINK = '<w:p><w:hyperlink r:id="rIdL"><w:r><w:t>Link</w:t></w:r></w:hyperlink></w:p>';
    const HEADER = `${DECL}<w:hdr xmlns:w="${W}" xmlns:r="${R}">${LINK}</w:hdr>`;
    const GLOSSARY =
        `${DECL}<w:glossaryDocument xmlns:w="${W}" xmlns:r="${R}"><w:docParts><w:docPart><w:docPartPr><w:name w:val="Part"/>` +
        `<w:category><w:name w:val="General"/><w:gallery w:val="placeholder"/></w:category><w:guid w:val="{6A9E2C40-3F1B-4C55-9D2E-7B1A0C9D4E11}"/></w:docPartPr>` +
        `<w:docPartBody>${LINK}</w:docPartBody></w:docPart></w:docParts></w:glossaryDocument>`;
    const FOOTNOTES = `${DECL}<w:footnotes xmlns:w="${W}" xmlns:r="${R}"><w:footnote w:id="1">${LINK}</w:footnote></w:footnotes>`;

    /** The main document stored at `part` holding the hyperlink. */
    const main = (part: string, segment: string): Record<string, string> => {
        const doc = once(withR(MAIN), "<w:sectPr>", `${LINK}<w:sectPr>`);
        const sidecar = addRelationship(P[SIDECAR]!, link(segment));
        if (part === "word/document.xml") return { ...P, [part]: doc, [SIDECAR]: sidecar };
        const { [SIDECAR]: _sidecar, "word/document.xml": _main, ...rest } = P;
        return {
            ...rest,
            [part]: doc,
            [`word/_rels/${path.posix.basename(part)}.rels`]: sidecar,
            "_rels/.rels": once(P["_rels/.rels"]!, 'Target="word/document.xml"', `Target="${part}"`),
            [CT]: once(P[CT]!, 'PartName="/word/document.xml"', `PartName="/${part}"`),
        };
    };
    /** A header stored as word/`name` holding the hyperlink; the main document relates it, and references it when `used`. */
    const header = (name: string, used: boolean, segment: string): Record<string, string> => ({
        ...P,
        ...(used
            ? { "word/document.xml": once(withR(MAIN), "<w:sectPr>", '<w:sectPr><w:headerReference w:type="default" r:id="rIdH"/>') }
            : {}),
        [SIDECAR]: addRelationship(P[SIDECAR]!, relationship("rIdH", "header", name)),
        [`word/${name}`]: HEADER,
        [`word/_rels/${name}.rels`]: relsPart(link(segment)),
        [CT]: override(`word/${name}`, `${WML}.header+xml`),
    });
    /** The same header stored as word/header9.xml, typed header, which no relationship names. */
    const unrelatedHeader = (segment: string): Record<string, string> => ({
        ...P,
        "word/header9.xml": HEADER,
        "word/_rels/header9.xml.rels": relsPart(link(segment)),
        [CT]: override("word/header9.xml", `${WML}.header+xml`),
    });
    /** A glossary holding the hyperlink; the main document relates it when `reached`. */
    const glossary = (reached: boolean, segment: string): Record<string, string> => ({
        ...P,
        "word/glossary/document.xml": GLOSSARY,
        "word/glossary/_rels/document.xml.rels": relsPart(link(segment)),
        ...(reached ? { [SIDECAR]: addRelationship(P[SIDECAR]!, relationship("rIdG", "glossaryDocument", "glossary/document.xml")) } : {}),
        [CT]: override("word/glossary/document.xml", `${WML}.document.glossary+xml`),
    });
    /** word/footnotes.xml whose footnote 1 holds the hyperlink; the main document relates it, and references it when `used`. */
    const footnotes = (used: boolean, segment: string): Record<string, string> => ({
        ...P,
        ...(used
            ? {
                  "word/document.xml": once(
                      MAIN,
                      "<w:r><w:t>Hello</w:t></w:r>",
                      '<w:r><w:t>Hello</w:t></w:r><w:r><w:footnoteReference w:id="1"/></w:r>',
                  ),
              }
            : {}),
        [SIDECAR]: addRelationship(P[SIDECAR]!, relationship("rIdN", "footnotes", "footnotes.xml")),
        "word/footnotes.xml": FOOTNOTES,
        "word/_rels/footnotes.xml.rels": relsPart(link(segment)),
        [CT]: override("word/footnotes.xml", `${WML}.footnotes+xml`),
    });

    /** Each placement, its builder, and the part Word judged the hyperlink in, or `null` where Word opened it clean. */
    const PLACEMENTS: readonly [string, (segment: string) => Record<string, string>, string | null][] = [
        ["word/document.xml (s28a, s82a)", (s) => main("word/document.xml", s), "word/document.xml"],
        ["a main document stored as word/main.xml (s82a)", (s) => main("word/main.xml", s), "word/main.xml"],
        ["a used header word/header1.xml (s82a)", (s) => header("header1.xml", true, s), "word/header1.xml"],
        ["a used header stored as word/hdr.xml (s82a)", (s) => header("hdr.xml", true, s), "word/hdr.xml"],
        ["an unrelated word/header9.xml (s82b)", unrelatedHeader, null],
        ["a glossary no relationship reaches (s82b)", (s) => glossary(false, s), null],
        ["a reached glossary (s82b)", (s) => glossary(true, s), "word/glossary/document.xml"],
        ["a header the main document relates but no section references (s82c)", (s) => header("header1.xml", false, s), null],
        ["a footnotes part nothing references (s82c)", (s) => footnotes(false, s), null],
        ["a footnotes part the body references (s82c)", (s) => footnotes(true, s), "word/footnotes.xml"],
    ];

    const run = async (parts: Record<string, string>, profile: "word-valid" | "strict") =>
        withTempDir(async (dir) => {
            const file = path.join(dir, "specimen.docx");
            writeFileSync(file, await buildDocx(parts));
            const result = await validate(file, { profile });
            return {
                valid: result.valid,
                judged: result.issues
                    .filter((issue) => issue.code === "word-hyperlink-relationship")
                    .map((issue) => `${issue.severity} ${issue.path}`),
                legacy: result.issues.filter((issue) => issue.code?.startsWith("rels-id-")).map((issue) => `${issue.code} ${issue.path}`),
            };
        });

    it.each(
        PLACEMENTS.flatMap(([name, build, part]) =>
            (["word-valid", "strict"] as const).map((profile) => [name, profile, build, part] as const),
        ),
    )("the refused Type in %s under %s is judged where Word judged it", async (_name, profile, build, part) => {
        const found = await run(build("HYPERLINK"), profile);
        expect(found.judged).toEqual(part === null ? [] : [`error ${part}`]);
        // Under word-valid the rule's error is the only blocking finding.
        if (profile === "word-valid") expect(found.valid).toBe(part === null);
    });

    it.each(PLACEMENTS.map(([name, build]) => [name, build] as const))("the calibration in %s passes word-valid", async (_name, build) => {
        expect(await run(build("hyperlink"), "word-valid")).toEqual({ valid: true, judged: [], legacy: [] });
    });

    it.each(["word-valid", "strict"] as const)(
        "a dangling id in a main document stored as word/main.xml is one report under %s, the legacy copy handed off",
        async (profile) => {
            const parts = main("word/main.xml", "hyperlink");
            parts["word/main.xml"] = once(parts["word/main.xml"]!, 'r:id="rIdL"', 'r:id="rIdX"');
            const found = await run(parts, profile);
            expect(found.judged).toEqual(["error word/main.xml"]);
            expect(found.legacy).toEqual([]);
        },
    );

    // #816: with no main document the read model reads nothing, so the rule
    // keeps judging the story parts by name, as the note checks do.
    it("keeps judging word/document.xml by name when the officeDocument relationship names no stored part", async () => {
        const parts = main("word/document.xml", "HYPERLINK");
        parts["_rels/.rels"] = once(parts["_rels/.rels"]!, 'Target="word/document.xml"', 'Target="word/missing.xml"');
        const found = await run(parts, "word-valid");
        expect(found.judged).toEqual(["error word/document.xml"]);
    });

    // #1118 (review 1115 F1): the read model lists a part once per segment
    // that reaches it; a header the body uses as its header and its footer is
    // judged once.
    it.each(["word-valid", "strict"] as const)(
        "a part the body uses as header and footer is judged once under %s (#1118)",
        async (profile) => {
            const parts = header("hf1.xml", true, "HYPERLINK");
            parts["word/document.xml"] = once(
                parts["word/document.xml"]!,
                '<w:headerReference w:type="default" r:id="rIdH"/>',
                '<w:headerReference w:type="default" r:id="rIdH"/><w:footerReference w:type="default" r:id="rIdF"/>',
            );
            parts[SIDECAR] = addRelationship(parts[SIDECAR]!, relationship("rIdF", "footer", "hf1.xml"));
            expect((await run(parts, profile)).judged).toEqual(["error word/hf1.xml"]);
        },
    );
});
