/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

/**
 * Issue #290: a BOM-prefixed UTF-16 XML part is legal OOXML — Word accepts it
 * and `validateXml` (libxml2, byte-level) already passes it — but every
 * downstream reader used to decode part bytes as UTF-8, so DOM reads produced
 * NUL-interleaved garbage and strict reported id-parse / ws-parse / del-parse
 * / ins-parse / rels-id-parse / comment-marker-parse plus an xsd-error
 * "missing root element". These tests pin the end-to-end acceptance:
 * `validate()` must pass a UTF-16 main part under every profile.
 */

import { promises as fs } from "node:fs";
import path from "node:path";

import { describe, expect, it } from "vitest";

import { withTempDir } from "../src/lib/run-cli";
import { validate } from "../src/scripts/office/validate";

import { buildDocx, MINIMAL_PARTS } from "./support/minimal-docx";

const DOCUMENT_XML = MINIMAL_PARTS["word/document.xml"]!;

const utf16le = (text: string): Buffer => Buffer.concat([Buffer.from([0xff, 0xfe]), Buffer.from(text, "utf16le")]);
const utf16be = (text: string): Buffer => Buffer.concat([Buffer.from([0xfe, 0xff]), Buffer.from(text, "utf16le").swap16()]);
const utf8Bom = (text: string): Buffer => Buffer.concat([Buffer.from([0xef, 0xbb, 0xbf]), Buffer.from(text, "utf-8")]);

describe("UTF-16 / BOM'd XML parts (#290)", () => {
    it("accepts a legal UTF-16 main part end to end", async () =>
        withTempDir(async (dir) => {
            const text = DOCUMENT_XML.replace("UTF-8", "UTF-16");
            const file = path.join(dir, "utf16.docx");
            await fs.writeFile(
                file,
                await buildDocx({
                    ...MINIMAL_PARTS,
                    "word/document.xml": utf16le(text),
                }),
            );
            expect((await validate(file, { profile: "strict" })).valid).toBe(true);
        }));

    it.each(["lenient", "word-valid"] as const)("accepts a legal UTF-16 main part under %s", async (profile) =>
        withTempDir(async (dir) => {
            const text = DOCUMENT_XML.replace("UTF-8", "UTF-16");
            const file = path.join(dir, "utf16.docx");
            await fs.writeFile(
                file,
                await buildDocx({
                    ...MINIMAL_PARTS,
                    "word/document.xml": utf16le(text),
                }),
            );
            expect((await validate(file, { profile })).valid).toBe(true);
        }),
    );

    it("accepts a UTF-16BE main part (control)", async () =>
        withTempDir(async (dir) => {
            const text = DOCUMENT_XML.replace("UTF-8", "UTF-16");
            const file = path.join(dir, "utf16be.docx");
            await fs.writeFile(
                file,
                await buildDocx({
                    ...MINIMAL_PARTS,
                    "word/document.xml": utf16be(text),
                }),
            );
            for (const profile of ["lenient", "word-valid", "strict"] as const) {
                expect((await validate(file, { profile })).valid).toBe(true);
            }
        }));

    it("accepts a UTF-8-with-BOM main part (control)", async () =>
        withTempDir(async (dir) => {
            const file = path.join(dir, "utf8bom.docx");
            await fs.writeFile(
                file,
                await buildDocx({
                    ...MINIMAL_PARTS,
                    "word/document.xml": utf8Bom(DOCUMENT_XML),
                }),
            );
            for (const profile of ["lenient", "word-valid", "strict"] as const) {
                expect((await validate(file, { profile })).valid).toBe(true);
            }
        }));
});
