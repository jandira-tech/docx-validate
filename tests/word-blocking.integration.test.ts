/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

import { promises as fs } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

import { describe, expect, it } from "vitest";

import { withTempDir } from "../src/lib/run-cli";
import type { ValidationIssue } from "../src/lib/types";
import { validate } from "../src/scripts/office/validate";
import { buildDocx, MINIMAL_PARTS } from "./support/minimal-docx";

// Package shapes from the Word 16.115 probes in
// docs/plans/evidence/word-probes-2026-10-24.md: Override + package relationship.
const PROPERTIES = [
    {
        kind: "core",
        part: "docProps/core.xml",
        type: "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties",
        contentType: "application/vnd.openxmlformats-package.core-properties+xml",
        xml: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/"><dc:title>t</dc:title></cp:coreProperties>`,
    },
    {
        kind: "app",
        part: "docProps/app.xml",
        type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties",
        contentType: "application/vnd.openxmlformats-officedocument.extended-properties+xml",
        xml: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"><Application>Microsoft Office Word</Application></Properties>`,
    },
    {
        kind: "custom",
        part: "docProps/custom.xml",
        type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties",
        contentType: "application/vnd.openxmlformats-officedocument.custom-properties+xml",
        xml: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"><property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="2" name="test"><vt:lpwstr>t</vt:lpwstr></property></Properties>`,
    },
] as const;

type Properties = (typeof PROPERTIES)[number];
const withProperties = (
    properties: Properties,
    xml: string,
    part: string = properties.part,
    referenced = true,
): Readonly<Record<string, string>> => ({
    ...MINIMAL_PARTS,
    [part]: xml,
    "_rels/.rels": referenced
        ? MINIMAL_PARTS["_rels/.rels"]!.replace(
              "</Relationships>",
              `<Relationship Id="rId2" Type="${properties.type}" Target="${part}"/></Relationships>`,
          )
        : MINIMAL_PARTS["_rels/.rels"]!,
    "[Content_Types].xml": MINIMAL_PARTS["[Content_Types].xml"]!.replace(
        "</Types>",
        `<Override PartName="/${part}" ContentType="${properties.contentType}"/></Types>`,
    ),
});

interface Specimen {
    name: string;
    parts: Readonly<Record<string, string>>;
    wordValid: boolean;
    code?: string;
    part?: string;
}
const SPECIMENS: readonly Specimen[] = [
    ...PROPERTIES.flatMap((properties): Specimen[] => [
        {
            name: `truncated-${properties.kind}`,
            parts: withProperties(properties, properties.xml.slice(0, -15)),
            wordValid: false,
            code: "xml-syntax",
            part: properties.part,
        },
        {
            name: `prefix-${properties.kind}`,
            parts: withProperties(properties, properties.xml.replace(/(<\/)([^>]+)>$/, `<dcx:title>t</dcx:title>$1$2>`)),
            wordValid: false,
            code: "xml-undeclared-prefix",
            part: properties.part,
        },
        {
            name: `unknown-encoding-${properties.kind}`,
            parts: withProperties(properties, properties.xml.replace('encoding="UTF-8"', 'encoding="X-UNKNOWN-ENCODING"')),
            wordValid: false,
            code: "xml-encoding",
            part: properties.part,
        },
        {
            name: `unreferenced-truncated-${properties.kind}`,
            parts: withProperties(properties, properties.xml.slice(0, -15), properties.part, false),
            wordValid: true,
            code: "xml-syntax",
            part: properties.part,
        },
        { name: `${properties.kind}-control`, parts: withProperties(properties, properties.xml), wordValid: true },
    ]),
    {
        name: "truncated-renamed-core",
        parts: withProperties(PROPERTIES[0], PROPERTIES[0].xml.slice(0, -15), "props/meta.xml"),
        wordValid: false,
        code: "xml-syntax",
        part: "props/meta.xml",
    },
    { name: "renamed-core-control", parts: withProperties(PROPERTIES[0], PROPERTIES[0].xml, "props/meta.xml"), wordValid: true },
];

const errorPairs = (issues: readonly ValidationIssue[]): string[] =>
    issues.filter((issue) => issue.severity === "error").map((issue) => `${issue.code}@${issue.path ?? ""}`);

describe.each(["packed", "directory"] as const)("relationship-selected properties policy on %s inputs (#409)", (input) => {
    it.each(["core", "app", "custom", "renamed-core"])(
        "judges %s defects and controls in full, quick, named and strict runs",
        { timeout: 120_000 },
        async (kind) => {
            const specimens = SPECIMENS.filter((specimen) =>
                kind === "core" ? !specimen.name.includes("renamed") && specimen.name.includes(kind) : specimen.name.includes(kind),
            );
            for (const specimen of specimens) {
                await withTempDir(async (dir) => {
                    const target = path.join(dir, `${specimen.name}.docx`);
                    const packed = await buildDocx(specimen.parts);
                    if (input === "packed") {
                        await fs.writeFile(target, packed);
                    } else {
                        for (const [name, xml] of Object.entries(specimen.parts)) {
                            const file = path.join(target, name);
                            await fs.mkdir(path.dirname(file), { recursive: true });
                            await fs.writeFile(file, xml);
                        }
                    }

                    const full = await validate(target, { profile: "word-valid" });
                    expect.soft(full.valid, `${specimen.name}: full word-valid verdict`).toBe(specimen.wordValid);
                    if (specimen.code !== undefined) {
                        expect.soft(full.issues, `${specimen.name}: full diagnostic`).toContainEqual(
                            expect.objectContaining({
                                code: specimen.code,
                                path: specimen.part,
                                severity: specimen.wordValid ? "warning" : "error",
                            }),
                        );
                    }
                    for (const selection of [{ quick: true }, { checks: ["xml-syntax"] }] as const) {
                        const partial = await validate(target, { profile: "word-valid", ...selection });
                        const label = `${specimen.name}: ${JSON.stringify(selection)}`;
                        expect.soft(partial.valid, label).toBe(specimen.wordValid);
                        if (specimen.code !== undefined) {
                            expect.soft(partial.issues, label).toContainEqual(
                                expect.objectContaining({
                                    code: specimen.code,
                                    path: specimen.part,
                                    severity: specimen.wordValid ? "warning" : "error",
                                }),
                            );
                        }
                        // #381: valid full implies valid partial; partial errors ⊆ full errors.
                        if (full.valid) expect.soft(partial.valid, `${label} valid-full implication`).toBe(true);
                        expect
                            .soft(errorPairs(full.issues), `${label} error subset`)
                            .toEqual(expect.arrayContaining(errorPairs(partial.issues)));
                    }

                    const strict = await validate(target, { profile: "strict" });
                    expect.soft(strict.valid, `${specimen.name}: strict verdict`).toBe(specimen.code === undefined);
                    if (specimen.code !== undefined) {
                        expect
                            .soft(strict.issues, `${specimen.name}: strict diagnostic`)
                            .toContainEqual(expect.objectContaining({ code: specimen.code, path: specimen.part, severity: "error" }));
                    }

                    if (input === "packed") {
                        expect(await fs.readFile(target), "packed input bytes").toEqual(packed);
                    } else {
                        for (const [name, xml] of Object.entries(specimen.parts)) {
                            expect(await fs.readFile(path.join(target, name)), name).toEqual(Buffer.from(xml));
                        }
                    }
                });
            }
        },
    );
});

it("keeps the Word-clean SuperDoc custom-properties prefix counterexample valid (#409, #381)", { timeout: 120_000 }, async () => {
    const fixture = fileURLToPath(new URL("./fixtures/external/superdoc/super-editor/annotations_import_2.docx", import.meta.url));
    const before = await fs.readFile(fixture);
    const full = await validate(fixture, { profile: "word-valid" });
    expect(full.valid).toBe(true);
    expect(full.issues).toContainEqual(
        expect.objectContaining({ code: "xml-undeclared-prefix", path: "docProps/custom.xml", severity: "warning" }),
    );
    for (const selection of [{ quick: true }, { checks: ["xml-syntax"] }] as const) {
        const partial = await validate(fixture, { profile: "word-valid", ...selection });
        expect(partial.valid, JSON.stringify(selection)).toBe(true);
        expect(errorPairs(full.issues)).toEqual(expect.arrayContaining(errorPairs(partial.issues)));
    }
    expect(await fs.readFile(fixture)).toEqual(before);
});
