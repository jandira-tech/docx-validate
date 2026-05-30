/*
 * Copyright 2026 Jandira Technologies, LLC
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *     http://www.apache.org/licenses/LICENSE-2.0
 */

/**
 * PR A Task A.3: segregation enforcement skeleton.
 *
 * The four-class architecture spec (§1.3) requires that no file under `src/`
 * outside of `src/node/` imports from `node:*`. The mechanical guarantee is
 * this test — it ripgreps for the pattern and fails if any match shows up.
 *
 * For now the test is SKIPPED. PR C is what relocates the current ~10
 * `node:*` importers under `src/` into `src/node/`; until then this would
 * fail. Skipping signals the intent + reserves the test slot so PR C just
 * has to flip `it.skip` to `it`.
 *
 * Spec: docs/superpowers/specs/2026-05-29-four-class-architecture-design.md §1.3
 * Plan: docs/superpowers/plans/2026-05-30-four-class-architecture-implementation.md §PR-C task C.8
 */

import { spawnSync } from "node:child_process";
import { describe, it, expect } from "vitest";

describe("src/ segregation: no node:* imports outside src/node/ (PR C un-skips this)", () => {
    it.skip("zero src/ files outside src/node/ import from node:*", () => {
        const result = spawnSync(
            "rg",
            [
                "-n",
                "--type",
                "ts",
                "-e",
                String.raw`^import .* from ['"]node:`,
                "src/",
                "--glob",
                "!src/node/**",
            ],
            { encoding: "utf-8" },
        );

        // ripgrep exits 1 when no matches — that's our success signal.
        // exits 0 means at least one match (failure for this assertion).
        const lines = result.stdout.split("\n").filter((l) => l.length > 0);
        expect(lines, `expected zero node:* imports outside src/node/, got:\n${lines.join("\n")}`).toEqual([]);
    });
});
