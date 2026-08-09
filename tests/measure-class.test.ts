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
import { Measure } from "../src/measure";
import type { AstAdapter } from "../src/measure";

const WORKING_FIXTURE = path.resolve(
  __dirname,
  "fixtures",
  "working",
  "sample-document.afterword-repaired-word-repaired.docx",
);

describe("Measure class (PR C task C.4)", () => {
  it("returns 'no-ast-adapter' classification when no adapter is injected", async () => {
    const bytes = readFileSync(WORKING_FIXTURE);
    const result = await new Measure().runOne(new Uint8Array(bytes));
    expect(result.classification).toBe("no-ast-adapter");
    expect(result.metrics.bodyBailoutCount).toBe(0);
    expect(result.metrics.t2ElementCarrierCount).toBe(0);
    expect(Array.isArray(result.diagnostics)).toBe(true);
  }, 30_000);

  it("runs the full pipeline when an AstAdapter is injected", async () => {
    const bytes = readFileSync(WORKING_FIXTURE);
    const fakeAdapter: AstAdapter = {
      async read(_b) {
        return { ast: { tag: "fake-ast" }, bodyBailoutCount: 1, t2ElementCarrierCount: 2 };
      },
      async write(_ast) {
        // Return the same bytes to test the byte-equivalent path.
        return new Uint8Array(bytes);
      },
    };
    const result = await new Measure({ astAdapter: fakeAdapter }).runOne(new Uint8Array(bytes));
    expect(result.classification).toBe("byte-equivalent");
    expect(result.metrics.bodyBailoutCount).toBe(1);
    expect(result.metrics.t2ElementCarrierCount).toBe(2);
  });

  it("classifies as 'ast-equivalent-byte-differs' when adapter rewrites bytes", async () => {
    const bytes = readFileSync(WORKING_FIXTURE);
    const fakeAdapter: AstAdapter = {
      async read(_b) {
        return { ast: {}, bodyBailoutCount: 0, t2ElementCarrierCount: 0 };
      },
      async write(_ast) {
        // Return slightly different bytes.
        const out = new Uint8Array(bytes);
        out[0] = (out[0]! ^ 0x01) as number;
        return out;
      },
    };
    const result = await new Measure({ astAdapter: fakeAdapter }).runOne(new Uint8Array(bytes));
    expect(result.classification).toBe("ast-equivalent-byte-differs");
  });
});
