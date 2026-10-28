/*
 * Copyright 2026 Jandira Technologies, LLC
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *     http://www.apache.org/licenses/LICENSE-2.0
 */

import { describe, it, expect } from "vitest";
import { readFileSync } from "node:fs";
import { tmpdir } from "node:os";
import path from "node:path";
import { Repair } from "../src/repair";
import { MINIMAL_PARTS, buildDocx } from "./support/minimal-docx";

const WORKING_FIXTURE = path.resolve(__dirname, "fixtures", "working", "sample-document.afterword-repaired-word-repaired.docx");

describe("Repair class (PR C task C.2)", () => {
    it("instantiates with no options", () => {
        expect(new Repair()).toBeInstanceOf(Repair);
    });

    it("returns a RepairResult shape (bytes + repairs + diagnostics)", async () => {
        const bytes = readFileSync(WORKING_FIXTURE);
        const result = await new Repair().run(new Uint8Array(bytes));
        expect(result.bytes).toBeInstanceOf(Uint8Array);
        expect(result.bytes.byteLength).toBeGreaterThan(0);
        expect(typeof result.repairs).toBe("number");
        expect(Array.isArray(result.diagnostics)).toBe(true);
    }, 30_000);
});

// #1157: a step that refuses the package is an error naming the step, not
// the ENOENT of reading an output pack never wrote, and no temp path leaks.
describe("Repair refusing a package (#1157)", () => {
    const refusal = async (bytes: Uint8Array): Promise<Error & { step?: string; issues?: unknown[] }> => {
        const error = await new Repair().run(bytes).then(
            () => undefined,
            (e: unknown) => e as Error & { step?: string; issues?: unknown[] },
        );
        expect(error, "the run resolved").toBeInstanceOf(Error);
        expect(error!.message).not.toContain("ENOENT");
        expect(error!.message).not.toContain(tmpdir());
        return error!;
    };

    it("rejects bytes that are not a package at the unpack step", async () => {
        const error = await refusal(new Uint8Array([1, 2, 3]));
        expect(error).toMatchObject({ name: "PackageStepError", step: "unpack" });
    }, 30_000);

    it("rejects a package that still fails validation at the pack step, with its issues", async () => {
        const bytes = await buildDocx({ ...MINIMAL_PARTS, "word/styles.xml": MINIMAL_PARTS["word/styles.xml"].replace("</w:styles>", "") });
        const error = await refusal(new Uint8Array(bytes));
        expect(error).toMatchObject({ name: "PackageStepError", step: "pack" });
        expect(error.issues).toEqual(expect.arrayContaining([expect.objectContaining({ severity: "error", code: "xml-syntax" })]));
        expect(JSON.stringify(error.issues)).not.toContain(tmpdir());
    }, 30_000);

    it("still returns bytes for the minimal package", async () => {
        const result = await new Repair().run(new Uint8Array(await buildDocx()));
        expect(result.bytes.byteLength).toBeGreaterThan(0);
    }, 30_000);

    it("exports the error from the barrel", async () => {
        const api: Record<string, unknown> = await import("../src/index");
        expect(typeof api.PackageStepError).toBe("function");
    });
});
