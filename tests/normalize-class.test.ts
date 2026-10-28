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
import { Normalize } from "../src/normalize";
import { MINIMAL_PARTS, buildDocx } from "./support/minimal-docx";

const WORKING_FIXTURE = path.resolve(__dirname, "fixtures", "working", "sample-document.afterword-repaired-word-repaired.docx");

describe("Normalize class (PR C task C.3)", () => {
    it("returns a NormalizeResult shape (bytes + changed + passes)", async () => {
        const bytes = readFileSync(WORKING_FIXTURE);
        const result = await new Normalize().run(new Uint8Array(bytes));
        expect(result.bytes).toBeInstanceOf(Uint8Array);
        expect(result.bytes.byteLength).toBeGreaterThan(0);
        expect(typeof result.changed).toBe("boolean");
        expect(typeof result.passes.mergeRuns).toBe("number");
    }, 30_000);

    it("is content-idempotent — second pass reports changed=false (mergeRuns count = 0)", async () => {
        // Content-level idempotency: once mergeRuns has collapsed sibling
        // runs on first call, the second call finds nothing to merge.
        // BYTE-level idempotency is NOT asserted here — JSZip's pack is
        // non-deterministic (CLAUDE.md note 3) so a.bytes !== b.bytes
        // even when the unpacked XML content is identical.
        const bytes = readFileSync(WORKING_FIXTURE);
        await new Normalize().run(new Uint8Array(bytes));
        const a = await new Normalize().run(new Uint8Array(bytes));
        const b = await new Normalize().run(a.bytes);
        expect(b.changed).toBe(false);
        expect(b.passes.mergeRuns).toBe(0);
    }, 60_000);
});

// #1157: a step that refuses the package is an error naming the step, not
// the ENOENT of reading an output pack never wrote, and no temp path leaks.
describe("Normalize refusing a package (#1157)", () => {
    const refusal = async (bytes: Uint8Array): Promise<Error & { step?: string; issues?: unknown[] }> => {
        const error = await new Normalize().run(bytes).then(
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
        const result = await new Normalize().run(new Uint8Array(await buildDocx()));
        expect(result.bytes.byteLength).toBeGreaterThan(0);
    }, 30_000);

    it("exports the error from the barrel", async () => {
        const api: Record<string, unknown> = await import("../src/index");
        expect(typeof api.PackageStepError).toBe("function");
    });
});
