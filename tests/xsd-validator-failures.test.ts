/*
 * Copyright 2026 Jandira Technologies, LLC
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *     http://www.apache.org/licenses/LICENSE-2.0
 */

import { afterEach, beforeAll, beforeEach, describe, expect, it, vi } from "vitest";
import { existsSync } from "node:fs";
import { readFile } from "node:fs/promises";
import path from "node:path";
import { XmlDocument, XmlValidateError, XsdValidator as WasmValidator } from "libxml2-wasm";
import {
    xmlRegisterInputProvider,
    xmlSchemaFreeValidCtxt,
    xmlSchemaNewValidCtxt,
    xmlSchemaSetValidStructuredErrors,
    xmlSchemaValidateDoc,
    type XmlInputProvider,
} from "libxml2-wasm/lib/libxml2.mjs";
import { fsInputProviders } from "libxml2-wasm/lib/nodejs.mjs";
import {
    _resetSchemaResourceRoots,
    _resetXsdValidatorMemo,
    createXsdValidator,
    schemaResourceRoots,
    xsdValidateStats,
} from "../src/lib/xsd-validator";

const engine = vi.hoisted(() => ({ hasAddFunction: true }));

// Namespace exports are immutable under ESM. Keep filesystem failures local to
// this suite, leaving the existing real-engine integration tests unaffected.
vi.mock("node:fs/promises", async (orig) => {
    const actual = await orig<typeof import("node:fs/promises")>();
    return { ...actual, readFile: vi.fn(actual.readFile) };
});

vi.mock("node:fs", async (orig) => {
    const actual = await orig<typeof import("node:fs")>();
    return { ...actual, existsSync: vi.fn(() => true) };
});

vi.mock("libxml2-wasm/lib/libxml2.mjs", async (orig) => {
    const actual = await orig<typeof import("libxml2-wasm/lib/libxml2.mjs")>();
    // addFunction is shipped by the dependency but missing from its types.
    const { addFunction } = actual as unknown as { addFunction: (fn: () => void, signature: string) => number };
    return {
        ...actual,
        get addFunction() {
            return engine.hasAddFunction ? addFunction : undefined;
        },
        xmlRegisterInputProvider: vi.fn(actual.xmlRegisterInputProvider),
        xmlSchemaNewValidCtxt: vi.fn(actual.xmlSchemaNewValidCtxt),
        xmlSchemaFreeValidCtxt: vi.fn(actual.xmlSchemaFreeValidCtxt),
        xmlSchemaSetValidStructuredErrors: vi.fn(actual.xmlSchemaSetValidStructuredErrors),
        xmlSchemaValidateDoc: vi.fn(actual.xmlSchemaValidateDoc),
    };
});

const SCHEMA = '<xs:schema xmlns:xs="http://www.w3.org/2001/XMLSchema"><xs:element name="root" type="xs:string"/></xs:schema>';
const SCHEMA_PATH = path.resolve(__dirname, "fixtures", "in-memory.xsd");
const VALID_XML = "<root>hello</root>";

// fromDoc is the same seam used by xsd-validator.test.ts. These intentionally
// incomplete instances model private fields disappearing in a future engine.
const fakeSchema = (value: unknown): void => {
    vi.spyOn(WasmValidator, "fromDoc").mockReturnValue(value as WasmValidator);
};

describe("xsd-validator engine and filesystem failures", () => {
    let provider: XmlInputProvider;

    beforeAll(async () => {
        await createXsdValidator();
        provider = vi.mocked(xmlRegisterInputProvider).mock.calls[0]![0];
    });

    beforeEach(() => {
        _resetXsdValidatorMemo();
        vi.clearAllMocks();
        vi.mocked(readFile).mockResolvedValue(SCHEMA);
    });

    afterEach(() => {
        vi.restoreAllMocks();
        engine.hasAddFunction = true;
        vi.mocked(existsSync).mockReturnValue(true);
        _resetXsdValidatorMemo();
        _resetSchemaResourceRoots();
    });

    it("reports a failed validation context without validating or freeing a null context", async () => {
        const publicValidate = vi.fn();
        fakeSchema({ _ptr: 1, validate: publicValidate });
        // The branded pointer type does not admit zero, though C returns it on
        // allocation failure. No fake pointer reaches the real WASM functions.
        vi.mocked(xmlSchemaNewValidCtxt).mockReturnValueOnce(0 as ReturnType<typeof xmlSchemaNewValidCtxt>);
        const validator = await createXsdValidator();

        expect(await validator.validate(VALID_XML, SCHEMA_PATH)).toEqual([
            { severity: "error", code: "xsd-validation-failed", message: "XSD engine failed: invalid input or internal error" },
        ]);
        expect(xmlSchemaNewValidCtxt).toHaveBeenCalledWith(1);
        expect(xmlSchemaValidateDoc).not.toHaveBeenCalled();
        expect(xmlSchemaSetValidStructuredErrors).not.toHaveBeenCalled();
        expect(xmlSchemaFreeValidCtxt).not.toHaveBeenCalled();
        expect(publicValidate).not.toHaveBeenCalled();
    });

    it.each([
        ["missing", {}],
        ["undefined", { _ptr: undefined }],
        ["string", { _ptr: "1" }],
        ["zero", { _ptr: 0 }],
        ["negative", { _ptr: -1 }],
        ["fractional", { _ptr: 1.5 }],
        ["unsafe integer", { _ptr: Number.MAX_SAFE_INTEGER + 1 }],
        ["NaN", { _ptr: Number.NaN }],
    ])("uses public validation for a %s schema pointer", async (_name, pointer) => {
        const publicValidate = vi.fn();
        fakeSchema({ ...pointer, validate: publicValidate });
        const validator = await createXsdValidator();

        expect(await validator.validate(VALID_XML, SCHEMA_PATH)).toEqual([]);
        expect(publicValidate).toHaveBeenCalledExactlyOnceWith(expect.any(XmlDocument));
        expect(xmlSchemaNewValidCtxt).not.toHaveBeenCalled();
    });

    it("uses public validation when a schema is callable rather than an object", async () => {
        const publicValidate = vi.fn();
        fakeSchema(Object.assign(() => undefined, { _ptr: 1, validate: publicValidate }));
        const validator = await createXsdValidator();

        expect(await validator.validate(VALID_XML, SCHEMA_PATH)).toEqual([]);
        expect(publicValidate).toHaveBeenCalledOnce();
        expect(xmlSchemaNewValidCtxt).not.toHaveBeenCalled();
    });

    it("reports a null schema as an engine failure without entering WASM", async () => {
        fakeSchema(null);
        const validator = await createXsdValidator();

        expect(await validator.validate(VALID_XML, SCHEMA_PATH)).toEqual([
            { severity: "error", code: "xsd-validation-failed", message: expect.stringContaining("XSD engine failed:") },
        ]);
        expect(xmlSchemaNewValidCtxt).not.toHaveBeenCalled();
    });

    it("uses public validation when the instance document has no private pointer", async () => {
        const publicValidate = vi.fn();
        const dispose = vi.fn();
        fakeSchema({ _ptr: 1, validate: publicValidate });
        const fromString = XmlDocument.fromString.bind(XmlDocument);
        vi.spyOn(XmlDocument, "fromString").mockImplementation((xml, options) =>
            xml === VALID_XML ? ({ dispose } as unknown as XmlDocument) : fromString(xml, options),
        );
        const validator = await createXsdValidator();

        expect(await validator.validate(VALID_XML, SCHEMA_PATH)).toEqual([]);
        expect(publicValidate).toHaveBeenCalledExactlyOnceWith({ dispose });
        expect(xmlSchemaNewValidCtxt).not.toHaveBeenCalled();
        expect(dispose).toHaveBeenCalledOnce();
    });

    it("accepts a zero status without collecting public validation messages", async () => {
        const publicValidate = vi.spyOn(WasmValidator.prototype, "validate");
        const before = { ...xsdValidateStats };
        const validator = await createXsdValidator();

        expect(await validator.validate(VALID_XML, SCHEMA_PATH)).toEqual([]);
        expect(xmlSchemaValidateDoc).toHaveReturnedWith(0);
        expect(xmlSchemaSetValidStructuredErrors).toHaveBeenCalledOnce();
        expect(xmlSchemaFreeValidCtxt).toHaveBeenCalledExactlyOnceWith(vi.mocked(xmlSchemaNewValidCtxt).mock.results[0]!.value);
        expect(publicValidate).not.toHaveBeenCalled();
        expect(xsdValidateStats).toEqual({ ...before, statusOnlyPasses: before.statusOnlyPasses + 1 });
    });

    it("bypasses the status probe when addFunction is unavailable", async () => {
        engine.hasAddFunction = false;
        const publicValidate = vi.fn();
        fakeSchema({ _ptr: 1, validate: publicValidate });
        const before = { ...xsdValidateStats };
        const validator = await createXsdValidator();

        expect(await validator.validate(VALID_XML, SCHEMA_PATH)).toEqual([]);
        expect(publicValidate).toHaveBeenCalledOnce();
        expect(xmlSchemaNewValidCtxt).not.toHaveBeenCalled();
        expect(xsdValidateStats).toEqual({ ...before, messageCollections: before.messageCollections + 1 });
    });

    it("retains a non-Error schema read failure and retries the failed cache entry", async () => {
        vi.mocked(readFile).mockRejectedValueOnce("schema storage unavailable");
        const validator = await createXsdValidator();

        expect(await validator.validate(VALID_XML, SCHEMA_PATH)).toEqual([
            {
                severity: "error",
                code: "xsd-schema-load-failed",
                message: `Schema load failed (${SCHEMA_PATH}): schema storage unavailable`,
                path: SCHEMA_PATH,
            },
        ]);
        expect(await validator.validate(VALID_XML, SCHEMA_PATH)).toEqual([]);
        expect(readFile).toHaveBeenCalledTimes(2);
    });

    it("retains a non-Error instance parse failure", async () => {
        const fromString = XmlDocument.fromString.bind(XmlDocument);
        vi.spyOn(XmlDocument, "fromString").mockImplementation((xml, options) => {
            if (xml === VALID_XML) throw "instance parser unavailable";
            return fromString(xml, options);
        });
        const validator = await createXsdValidator();

        expect(await validator.validate(VALID_XML, SCHEMA_PATH)).toEqual([
            { severity: "error", code: "xml-parse-error", message: "instance parser unavailable" },
        ]);
    });

    it.each([new Error("engine unavailable"), "engine unavailable"])(
        "retains a generic public validation failure (%s)",
        async (failure) => {
            fakeSchema({
                validate: () => {
                    throw failure;
                },
            });
            const validator = await createXsdValidator();

            expect(await validator.validate(VALID_XML, SCHEMA_PATH)).toEqual([
                { severity: "error", code: "xsd-validation-failed", message: "XSD engine failed: engine unavailable" },
            ]);
        },
    );

    it("preserves available error positions and omits unavailable positions", async () => {
        // Use the dependency's error constructor; its declared line field is
        // required even though the adapter defensively handles missing lines.
        const failure = XmlValidateError.fromDetails([
            { message: "located", level: 2, col: 4, line: 7, file: "word/document.xml" },
            { message: "unlocated", level: 2, col: 0, line: undefined as unknown as number },
        ]);
        fakeSchema({
            validate: () => {
                throw failure;
            },
        });
        const validator = await createXsdValidator();

        expect(await validator.validate(VALID_XML, SCHEMA_PATH)).toEqual([
            { severity: "error", code: "xsd-validation-failed", message: "located", path: "word/document.xml", line: 7 },
            { severity: "error", code: "xsd-validation-failed", message: "unlocated" },
        ]);
    });

    it("keeps the source schema root when neither bundled candidate exists", () => {
        vi.mocked(existsSync).mockReturnValue(false);

        _resetSchemaResourceRoots();

        expect(schemaResourceRoots()).toEqual([path.resolve(__dirname, "../src/scripts/office/schemas")]);
        expect(vi.mocked(existsSync).mock.calls.map(([candidate]) => candidate)).toEqual([
            path.resolve(__dirname, "../src/scripts/office/schemas"),
            path.resolve(__dirname, "../src/lib/schemas"),
        ]);
    });

    it.each(["relative.xsd", "https://example.invalid/schema.xsd", "file:///bad%ZZ.xsd", "file://remote-host/schema.xsd"])(
        "refuses the resource %s before delegating to filesystem IO",
        (filename) => {
            const open = vi.spyOn(fsInputProviders, "open");

            expect(provider.open(filename)).toBeUndefined();
            expect(open).not.toHaveBeenCalled();
        },
    );

    it("refuses a relative resource inside an allowed schema root before filesystem IO", () => {
        const absolute = path.join(schemaResourceRoots()[0]!, "probe.xsd");
        const relative = path.relative(process.cwd(), absolute);
        const open = vi.spyOn(fsInputProviders, "open").mockReturnValue(42);

        // The root allow-list accepts this target, so it cannot mask a missing
        // relative-path refusal. The fake opener never touches the filesystem.
        expect(provider.open(absolute)).toBe(42);
        expect(open).toHaveBeenCalledExactlyOnceWith(absolute);
        open.mockClear();

        expect(provider.open(relative)).toBeUndefined();
        expect(open).not.toHaveBeenCalled();
    });
});
