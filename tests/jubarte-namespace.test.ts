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
import path from "node:path";
import { jubarte, Validate, Repair, Normalize, Measure } from "../src/index";

const WORKING_FIXTURE = path.resolve(__dirname, "fixtures", "working", "sample-document.afterword-repaired-word-repaired.docx");

describe("jubarte namespace (PR C task C.5)", () => {
    it("exposes validate / repair / normalize / measure as functions", () => {
        expect(typeof jubarte.validate).toBe("function");
        expect(typeof jubarte.repair).toBe("function");
        expect(typeof jubarte.normalize).toBe("function");
        expect(typeof jubarte.measure).toBe("function");
    });

    it("jubarte.validate(bytes) matches new Validate().run(bytes) — same result shape", async () => {
        const bytes = new Uint8Array(readFileSync(WORKING_FIXTURE));
        const fromNamespace = await jubarte.validate(bytes);
        const fromClass = await new Validate().run(bytes);
        expect(fromNamespace.valid).toBe(fromClass.valid);
        expect(fromNamespace.issues.length).toBe(fromClass.issues.length);
    }, 60_000);

    it("re-exports the four classes from src/index", () => {
        expect(typeof Validate).toBe("function");
        expect(typeof Repair).toBe("function");
        expect(typeof Normalize).toBe("function");
        expect(typeof Measure).toBe("function");
    });
});
