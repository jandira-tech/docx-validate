/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

import { writeFileSync } from "node:fs";
import path from "node:path";

import { describe, expect, it, vi } from "vitest";

import type { CheckDefinition } from "../src/lib/check-registry";
import { withTempDir } from "../src/lib/run-cli";
import type { Profile } from "../src/lib/types";
import { validate } from "../src/scripts/office/validate";
import { buildDocx } from "./support/minimal-docx";

// A stand-in word rule: one advisory per package that has a document part.
const probe = vi.hoisted((): CheckDefinition => ({
    id: "probe-word-rule",
    description: "Test probe registered as a word rule",
    groups: ["story"],
    formats: ["docx"],
    run: (ctx) =>
        ctx.has("word/document.xml") ? [{ severity: "warning", code: "probe-word-rule", message: "probe", path: "word/document.xml" }] : [],
}));

vi.mock("../src/scripts/office/checks/index", async (importOriginal) => {
    const actual = await importOriginal<typeof import("../src/scripts/office/checks/index")>();
    const { CheckRegistry } = await import("../src/lib/check-registry");
    const rules = [...actual.WORD_RULE_CHECKS, probe];
    return {
        ...actual,
        WORD_RULE_CHECKS: rules,
        createDefaultCheckRegistry: () => new CheckRegistry([...actual.createDefaultCheckRegistry().list(), probe]),
    };
});

const validateMinimal = async (profile: Profile) => {
    const bytes = await buildDocx();
    return withTempDir(async (dir) => {
        const file = path.join(dir, "minimal.docx");
        writeFileSync(file, bytes);
        return validate(file, { profile });
    });
};

describe("registered Word rule checks inside the DOCX validator", () => {
    it.each<Profile>(["word-valid", "strict"])("runs every registered word rule under %s", async (profile) => {
        const result = await validateMinimal(profile);
        const probes = result.issues.filter((issue) => issue.code === "probe-word-rule");
        expect(probes).toEqual([expect.objectContaining({ severity: "warning", path: "word/document.xml" })]);
        expect(result.valid).toBe(true);
    });

    it("leaves lenient output unchanged", async () => {
        const result = await validateMinimal("lenient");
        expect(result.issues.map((issue) => issue.code)).not.toContain("probe-word-rule");
    });
});
