/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

/**
 * Unit tests for the shared mode policy (`docs/plans/2026-10-23-fixing-validator.md`
 * § 1 "Contracts", task 10). Pure functions, no I/O: profile defaults, the
 * quick preset, partial-request detection, and every conflict rule.
 */

import { describe, expect, it } from "vitest";

import {
    assertModeCoherent,
    defaultProfileForFormat,
    isPartialRequest,
    ModeConflictError,
    profileWalksSchemas,
    QUICK_CHECK_GROUPS,
    resolveMode,
    wordNameIssueSeverity,
    type ModeRequest,
} from "../src/lib/mode-policy";
import { CHECK_GROUPS } from "../src/lib/check-registry";

describe("defaultProfileForFormat", () => {
    it("defaults DOCX to word-valid and every other format to lenient", () => {
        expect(defaultProfileForFormat("docx")).toBe("word-valid");
        expect(defaultProfileForFormat("pptx")).toBe("lenient");
        expect(defaultProfileForFormat("xlsx")).toBe("lenient");
    });
});

describe("QUICK_CHECK_GROUPS", () => {
    it("names the six quick-preset families, all of them declared groups", () => {
        expect([...QUICK_CHECK_GROUPS]).toEqual(["package", "xml", "content-types", "relationships", "tables", "drawings"]);
        for (const group of QUICK_CHECK_GROUPS) {
            expect(CHECK_GROUPS).toContain(group);
        }
    });
});

describe("isPartialRequest", () => {
    it("is true only for quick or a non-empty check list", () => {
        expect(isPartialRequest({})).toBe(false);
        expect(isPartialRequest({ checks: [] })).toBe(false);
        expect(isPartialRequest({ quick: true })).toBe(true);
        expect(isPartialRequest({ checks: ["tables"] })).toBe(true);
    });
});

describe("profileWalksSchemas", () => {
    it.each([
        ["docx", "lenient", true],
        ["docx", "strict", true],
        // A DOCX under word-valid compiles no schema, so it never needs the XSD engine.
        ["docx", "word-valid", false],
        // The PPTX validator walks its schemas under every profile (#341).
        ["pptx", "lenient", true],
        ["pptx", "strict", true],
        ["pptx", "word-valid", true],
        // No format validator runs for XLSX, so nothing compiles a schema (#341).
        ["xlsx", "lenient", false],
        ["xlsx", "strict", false],
        ["xlsx", "word-valid", false],
    ] as const)("%s under %s: %s", (format, profile, walks) => {
        expect(profileWalksSchemas(format, profile)).toBe(walks);
    });
});

describe("resolveMode", () => {
    it("falls back to the per-format default and covers every check", () => {
        expect(resolveMode("docx", {})).toEqual({ profile: "word-valid", select: undefined });
        expect(resolveMode("pptx", {})).toEqual({ profile: "lenient", select: undefined });
    });

    it("honours an explicit profile", () => {
        expect(resolveMode("docx", { profile: "lenient" }).profile).toBe("lenient");
        expect(resolveMode("pptx", { profile: "word-valid" }).profile).toBe("word-valid");
        expect(resolveMode("docx", { profile: "strict" }).profile).toBe("strict");
    });

    it("treats --strict as shorthand for the strict profile", () => {
        expect(resolveMode("docx", { strict: true })).toEqual({ profile: "strict", select: undefined });
    });

    it("selects the quick preset for --quick", () => {
        expect(resolveMode("docx", { quick: true })).toEqual({ profile: "word-valid", select: QUICK_CHECK_GROUPS });
        expect(resolveMode("pptx", { quick: true })).toEqual({ profile: "lenient", select: QUICK_CHECK_GROUPS });
    });

    it("copies the named check list rather than aliasing the caller's array", () => {
        const checks = ["tables"];
        const mode = resolveMode("docx", { checks });
        expect(mode.select).toEqual(checks);
        expect(mode.select).not.toBe(checks);
    });

    it("treats an empty check list as no selection at all", () => {
        expect(resolveMode("docx", { checks: [] }).select).toBeUndefined();
    });
});

describe("assertModeCoherent", () => {
    const ACCEPTED: readonly ModeRequest[] = [
        {},
        { profile: "lenient" },
        { profile: "word-valid" },
        { strict: true },
        { strict: true, profile: "strict" },
        { profile: "strict" },
        { quick: true },
        { checks: ["tables"] },
        { quick: true, profile: "lenient" },
        // `--checks ""` names nothing, so there is nothing to conflict with.
        { quick: true, checks: [] },
    ];

    for (const request of ACCEPTED) {
        it(`accepts ${JSON.stringify(request)}`, () => {
            expect(() => assertModeCoherent(request)).not.toThrow();
        });
    }

    const REJECTED: readonly { request: ModeRequest; message: RegExp }[] = [
        { request: { quick: true, checks: ["tables"] }, message: /--quick with --checks/ },
        { request: { strict: true, quick: true }, message: /--strict with --quick or --checks/ },
        { request: { strict: true, checks: ["tables"] }, message: /--strict with --quick or --checks/ },
        { request: { strict: true, profile: "lenient" }, message: /--strict with --profile lenient/ },
        { request: { strict: true, profile: "word-valid" }, message: /--strict with --profile word-valid/ },
        { request: { profile: "strict", quick: true }, message: /profile 'strict' with --quick or --checks/ },
        { request: { profile: "strict", checks: ["tables"] }, message: /profile 'strict' with --quick or --checks/ },
    ];

    for (const { request, message } of REJECTED) {
        it(`rejects ${JSON.stringify(request)}`, () => {
            expect(() => assertModeCoherent(request)).toThrow(ModeConflictError);
            expect(() => assertModeCoherent(request)).toThrow(message);
        });
    }

    it("throws a named, catchable error", () => {
        try {
            assertModeCoherent({ quick: true, checks: ["tables"] });
            expect.unreachable("expected a ModeConflictError");
        } catch (err) {
            expect(err).toBeInstanceOf(ModeConflictError);
            expect((err as Error).name).toBe("ModeConflictError");
        }
    });
});

describe("wordNameIssueSeverity", () => {
    it("errors only for DOCX/DOCM under word-valid and strict, warning for every PPTX/XLSX profile", () => {
        for (const profile of ["word-valid", "strict", "lenient"] as const)
            for (const format of ["docx", "pptx", "xlsx"] as const)
                expect(wordNameIssueSeverity(profile, format)).toBe(format === "docx" && profile !== "lenient" ? "error" : "warning");
    });
});
