/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

/**
 * #866: an editing step must put a part back the way it found it. A UTF-16
 * part read as a UTF-8 string turns into NUL-interleaved garbage, and
 * writing that garbage back corrupts the part.
 */

import { promises as fs } from "node:fs";
import path from "node:path";

import { describe, expect, it } from "vitest";

import { withTempDir } from "../src/lib/run-cli";
import { readXmlPart, writeXmlPart, writeXmlText } from "../src/lib/xml-part-io";
import { parseXml } from "../src/lib/xml-helpers";

const DECLARATION = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
const BODY = "<document><body/></document>";

/** A part stored in UTF-16, with the byte order mark that says which way round it is. */
function utf16Part(text: string, endian: "le" | "be"): Buffer {
    const body = Buffer.from(text, "utf16le");
    if (endian === "be") body.swap16();
    return Buffer.concat([Buffer.from(endian === "le" ? [0xff, 0xfe] : [0xfe, 0xff]), body]);
}

describe("readXmlPart", () => {
    it("reads a part and reports the encoding it is stored in", async () =>
        withTempDir(async (dir) => {
            const file = path.join(dir, "document.xml");
            await fs.writeFile(file, `${DECLARATION}${BODY}`, "utf8");

            const part = await readXmlPart(file);
            expect(part?.text).toBe(`${DECLARATION}${BODY}`);
            expect(part?.encoding).toBe("utf-8");
            expect(part?.bom).toBeUndefined();
        }));

    it("reads a UTF-16 part with its byte order mark", async () =>
        withTempDir(async (dir) => {
            const file = path.join(dir, "document.xml");
            await fs.writeFile(file, utf16Part(`<?xml version="1.0" encoding="UTF-16"?>${BODY}`, "be"));

            const part = await readXmlPart(file);
            expect(part?.text).toBe(`<?xml version="1.0" encoding="UTF-16"?>${BODY}`);
            expect(part?.encoding).toBe("utf-16be");
            expect(part?.bom).toBe("utf-16be");
        }));

    it("refuses a part whose declared encoding it cannot decode", async () =>
        withTempDir(async (dir) => {
            const file = path.join(dir, "document.xml");
            await fs.writeFile(file, '<?xml version="1.0" encoding="macintosh"?><a/>', "utf8");

            expect(await readXmlPart(file)).toBeUndefined();
        }));
});

describe("writeXmlPart", () => {
    it("writes a UTF-16 part back as UTF-16, byte order mark and declaration included", async () =>
        withTempDir(async (dir) => {
            const file = path.join(dir, "document.xml");
            await fs.writeFile(file, utf16Part(`<?xml version="1.0" encoding="UTF-16"?>${BODY}`, "le"));
            const part = await readXmlPart(file);

            await writeXmlPart(file, parseXml(part!.text), part!);

            const bytes = await fs.readFile(file);
            expect([...bytes.subarray(0, 2)]).toEqual([0xff, 0xfe]);
            expect(Buffer.from(bytes.subarray(2)).toString("utf16le")).toContain('encoding="UTF-16"');
        }));

    it("writes a UTF-8 part back as UTF-8", async () =>
        withTempDir(async (dir) => {
            const file = path.join(dir, "document.xml");
            await fs.writeFile(file, `${DECLARATION}${BODY}`, "utf8");
            const part = await readXmlPart(file);

            await writeXmlPart(file, parseXml(part!.text), part!);

            const bytes = await fs.readFile(file);
            expect(bytes.toString("utf8")).toContain('encoding="UTF-8"');
            expect(bytes.toString("utf8")).toContain("<body/>");
        }));

    it("keeps the encoding of the part on disk when the caller holds only a DOM", async () =>
        withTempDir(async (dir) => {
            const file = path.join(dir, "document.xml");
            await fs.writeFile(file, utf16Part(`<?xml version="1.0" encoding="UTF-16"?>${BODY}`, "le"));
            const part = await readXmlPart(file);

            await writeXmlPart(file, parseXml(part!.text));

            const bytes = await fs.readFile(file);
            expect([...bytes.subarray(0, 2)]).toEqual([0xff, 0xfe]);
        }));

    it("leaves a part it cannot decode untouched", async () =>
        withTempDir(async (dir) => {
            const file = path.join(dir, "document.xml");
            const original = '<?xml version="1.0" encoding="macintosh"?><a/>';
            await fs.writeFile(file, original, "utf8");

            await writeXmlPart(file, parseXml("<a/>"));

            expect((await fs.readFile(file)).toString("utf8")).toBe(original);
        }));
});

describe("writeXmlText", () => {
    it("writes text in the part's own encoding", async () =>
        withTempDir(async (dir) => {
            const file = path.join(dir, "document.xml");
            await fs.writeFile(file, utf16Part(`<?xml version="1.0" encoding="UTF-16"?>${BODY}`, "le"));
            const part = await readXmlPart(file);

            await writeXmlText(file, part!.text.replace("<body/>", "<body><p/></body>"), part!);

            const bytes = await fs.readFile(file);
            expect([...bytes.subarray(0, 2)]).toEqual([0xff, 0xfe]);
            expect(Buffer.from(bytes.subarray(2)).toString("utf16le")).toContain("<p/>");
        }));
});
