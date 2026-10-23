/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

import { describe, expect, it } from "vitest";

import { type CheckDefinition, CheckRegistry, UnknownCheckError } from "../src/lib/check-registry";
import { runChecks } from "../src/lib/check-scheduler";
import type { ValidationIssue } from "../src/lib/types";
import { ValidationContext } from "../src/lib/validation-context";

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const DOCUMENT = `<w:document xmlns:w="${W}"><w:body><w:p/></w:body></w:document>`;

const error = (code: string, path?: string): ValidationIssue => ({ severity: "error", code, message: code, ...(path ? { path } : {}) });

function check(id: string, extra: Partial<CheckDefinition> = {}): CheckDefinition {
    return { id, description: `${id} check`, groups: ["xml"], formats: ["docx"], run: () => [], ...extra };
}

describe("CheckRegistry", () => {
    it("lists definitions in registration order and filters by format", () => {
        const registry = new CheckRegistry([check("a"), check("b", { formats: ["pptx"] }), check("c", { formats: ["docx", "pptx"] })]);
        expect(registry.ids).toEqual(["a", "b", "c"]);
        expect(registry.list("docx").map((d) => d.id)).toEqual(["a", "c"]);
        expect(registry.list("pptx").map((d) => d.id)).toEqual(["b", "c"]);
        expect(registry.list().map((d) => d.id)).toEqual(["a", "b", "c"]);
        expect(registry.get("b")?.formats).toEqual(["pptx"]);
        expect(registry.get("zz")).toBeUndefined();
    });

    it("rejects duplicate ids, unknown prerequisites, and prerequisites registered later", () => {
        expect(() => new CheckRegistry([check("a"), check("a")])).toThrow(/duplicate/i);
        expect(() => new CheckRegistry([check("a", { requires: ["missing"] })])).toThrow(/missing/);
        expect(() => new CheckRegistry([check("a", { requires: ["b"] }), check("b")])).toThrow(/before/i);
    });

    it("rejects a prerequisite that cannot run for the dependent check's format", () => {
        // Issue #299: a docx-only dependent of a pptx-only prerequisite never
        // runs — the scheduler would skip it for every docx package and report
        // a valid run with nothing executed, so the configuration is rejected
        // up front like an unknown prerequisite.
        expect(() => new CheckRegistry([check("a", { formats: ["pptx"] }), check("b", { formats: ["docx"], requires: ["a"] })])).toThrow(
            /"b" requires "a".*docx/,
        );
        expect(
            () => new CheckRegistry([check("a", { formats: ["pptx"] }), check("b", { formats: ["docx", "pptx"], requires: ["a"] })]),
        ).toThrow(/does not apply.*docx/);
        // A prerequisite that covers every dependent format constructs fine.
        expect(
            () => new CheckRegistry([check("gate", { formats: ["docx", "pptx"] }), check("dependent", { requires: ["gate"] })]),
        ).not.toThrow();
    });

    it("resolves a selection to its checks and their prerequisites, in registration order", () => {
        const registry = new CheckRegistry([
            check("xml"),
            check("content-types", { requires: ["xml"], groups: ["content-types"] }),
            check("rows", { requires: ["xml"], groups: ["tables"] }),
            check("cells", { requires: ["rows"], groups: ["tables"] }),
        ]);
        expect(registry.resolve(["cells"])).toEqual(["xml", "rows", "cells"]);
        expect(registry.resolve(["content-types", "rows"])).toEqual(["xml", "content-types", "rows"]);
        expect(registry.resolve(["tables"])).toEqual(["xml", "rows", "cells"]);
    });

    it("fails configuration for unknown ids instead of silently running nothing", () => {
        const registry = new CheckRegistry([check("xml")]);
        expect(() => registry.resolve(["xml", "nope", "nada"])).toThrow(UnknownCheckError);
        try {
            registry.resolve(["nope"]);
        } catch (err) {
            expect((err as UnknownCheckError).ids).toEqual(["nope"]);
        }
        expect(() => registry.resolve([])).toThrow(/empty/i);
    });

    it("limits a group to the package format but rejects an explicitly named check of another format", () => {
        const registry = new CheckRegistry([
            check("docx-rows", { groups: ["tables"] }),
            check("pptx-cells", { groups: ["tables"], formats: ["pptx"] }),
        ]);
        expect(registry.resolve(["tables"], "docx")).toEqual(["docx-rows"]);
        expect(registry.resolve(["tables"], "pptx")).toEqual(["pptx-cells"]);
        expect(registry.resolve(["tables"])).toEqual(["docx-rows", "pptx-cells"]);
        expect(() => registry.resolve(["pptx-cells"], "docx")).toThrow(/pptx-cells does not apply to docx/);
        expect(() => registry.resolve(["tables", "pptx-cells"], "docx")).toThrow(/does not apply/);
    });
});

describe("runChecks", () => {
    it("yields the defect of an in-memory package from its named check", async () => {
        const registry = new CheckRegistry([
            check("has-body", {
                run: (ctx) => {
                    const parsed = ctx.xml("word/document.xml");
                    if (!parsed?.ok) return [];
                    return parsed.document.getElementsByTagNameNS(W, "body").length === 0 ? [error("no-body", "word/document.xml")] : [];
                },
            }),
        ]);
        const ctx = ValidationContext.fromEntries({ "word/document.xml": `<w:document xmlns:w="${W}"/>` });
        const run = await runChecks(ctx, registry, { format: "docx" });
        expect(run.valid).toBe(false);
        expect(run.issues).toEqual([error("no-body", "word/document.xml")]);
        expect(run.outcomes).toEqual([{ id: "has-body", status: "failed", issues: [error("no-body", "word/document.xml")] }]);
        expect(run.executed).toEqual(["has-body"]);
        expect(run.omitted).toEqual([]);
    });

    it("runs a shared prerequisite once and skips dependents of a failed prerequisite without crashing them", async () => {
        const calls: string[] = [];
        const registry = new CheckRegistry([
            check("xml-syntax", {
                run: (ctx) => {
                    calls.push("xml-syntax");
                    const parsed = ctx.xml("word/document.xml");
                    return parsed && !parsed.ok ? [error("xml-syntax", "word/document.xml")] : [];
                },
            }),
            check("dom-a", {
                requires: ["xml-syntax"],
                run: (ctx) => {
                    calls.push("dom-a");
                    const parsed = ctx.xml("word/document.xml");
                    if (!parsed?.ok) throw new Error("dom-a ran on malformed XML");
                    return [];
                },
            }),
            check("dom-b", {
                requires: ["xml-syntax"],
                run: () => {
                    calls.push("dom-b");
                    return [];
                },
            }),
        ]);
        const malformed = ValidationContext.fromEntries({ "word/document.xml": "<w:document>" });
        const run = await runChecks(malformed, registry, { format: "docx", select: ["dom-a", "dom-b"] });
        expect(calls).toEqual(["xml-syntax"]);
        expect(run.issues.map((i) => i.code)).toEqual(["xml-syntax"]);
        expect(run.outcomes.map((o) => [o.id, o.status])).toEqual([
            ["xml-syntax", "failed"],
            ["dom-a", "skipped"],
            ["dom-b", "skipped"],
        ]);
        expect(run.outcomes[1]?.reason).toMatch(/xml-syntax/);
        expect(run.skipped).toEqual(["dom-a", "dom-b"]);
        expect(run.valid).toBe(false);

        calls.length = 0;
        const clean = ValidationContext.fromEntries({ "word/document.xml": DOCUMENT });
        const ok = await runChecks(clean, registry, { format: "docx", select: ["dom-a", "dom-b"] });
        expect(calls).toEqual(["xml-syntax", "dom-a", "dom-b"]);
        expect(ok.valid).toBe(true);
        expect(ok.executed).toEqual(["xml-syntax", "dom-a", "dom-b"]);
    });

    it("lets two checks that share a part reuse its bytes and parsed DOM", async () => {
        let loads = 0;
        const ctx = ValidationContext.fromEntries([
            [
                "word/document.xml",
                () => {
                    loads += 1;
                    return new TextEncoder().encode(DOCUMENT);
                },
            ],
        ]);
        const seen: Document[] = [];
        const reader = (id: string) =>
            check(id, {
                run: (c) => {
                    const parsed = c.xml("word/document.xml");
                    if (parsed?.ok) seen.push(parsed.document);
                    return [];
                },
            });
        await runChecks(ctx, new CheckRegistry([reader("one"), reader("two")]), { format: "docx" });
        expect(loads).toBe(1);
        expect(ctx.parseCount("word/document.xml")).toBe(1);
        expect(seen).toHaveLength(2);
        expect(seen[0]).toBe(seen[1]);
    });

    it("turns a check that throws into an error finding and keeps running the others", async () => {
        const registry = new CheckRegistry([
            check("boom", {
                run: () => {
                    throw new Error("kaput");
                },
            }),
            check("fine"),
        ]);
        const run = await runChecks(ValidationContext.fromEntries({}), registry, { format: "docx" });
        expect(run.outcomes.map((o) => [o.id, o.status])).toEqual([
            ["boom", "failed"],
            ["fine", "passed"],
        ]);
        expect(run.issues).toHaveLength(1);
        expect(run.issues[0]).toMatchObject({ severity: "error", code: "check-crashed" });
        expect(run.issues[0]?.message).toMatch(/boom.*kaput/);
        expect(run.valid).toBe(false);
    });

    it("reports archive-only checks unavailable for an unpacked package rather than passed", async () => {
        const registry = new CheckRegistry([
            check("zip", { archiveOnly: true, groups: ["package"], run: () => [error("zip")] }),
            check("dom"),
        ]);
        const run = await runChecks(ValidationContext.fromEntries({}, { source: "directory" }), registry, { format: "docx" });
        expect(run.outcomes.map((o) => [o.id, o.status])).toEqual([
            ["zip", "unavailable"],
            ["dom", "passed"],
        ]);
        expect(run.unavailable).toEqual(["zip"]);
        expect(run.executed).toEqual(["dom"]);
        expect(run.valid).toBe(true);
    });

    it("runs archive-only checks when the archive bytes are present", async () => {
        const registry = new CheckRegistry([check("zip", { archiveOnly: true, run: (ctx) => (ctx.archiveBytes ? [error("zip")] : []) })]);
        const ctx = ValidationContext.fromEntries({}, { source: "archive", archiveBytes: new Uint8Array([0x50, 0x4b]) });
        const run = await runChecks(ctx, registry, { format: "docx" });
        expect(run.issues.map((i) => i.code)).toEqual(["zip"]);
    });

    it("reports the applicable checks a partial selection omitted", async () => {
        const registry = new CheckRegistry([check("a"), check("b"), check("c", { formats: ["pptx"] }), check("d", { requires: ["a"] })]);
        const run = await runChecks(ValidationContext.fromEntries({}), registry, { format: "docx", select: ["d"] });
        expect(run.executed).toEqual(["a", "d"]);
        expect(run.omitted).toEqual(["b"]);
    });

    it("skips a dependent whose prerequisite was unavailable", async () => {
        const registry = new CheckRegistry([check("zip", { archiveOnly: true }), check("after-zip", { requires: ["zip"] })]);
        const run = await runChecks(ValidationContext.fromEntries({}, { source: "directory" }), registry, { format: "docx" });
        expect(run.outcomes.map((o) => [o.id, o.status])).toEqual([
            ["zip", "unavailable"],
            ["after-zip", "skipped"],
        ]);
    });

    it("treats warnings as passing, so dependents still run", async () => {
        const registry = new CheckRegistry([
            check("warns", { run: () => [{ severity: "warning", code: "w", message: "w" }] }),
            check("next", { requires: ["warns"] }),
        ]);
        const run = await runChecks(ValidationContext.fromEntries({}), registry, { format: "docx" });
        expect(run.outcomes.map((o) => [o.id, o.status])).toEqual([
            ["warns", "passed"],
            ["next", "passed"],
        ]);
        expect(run.valid).toBe(true);
        expect(run.issues).toHaveLength(1);
    });

    it("rejects selecting a check that does not apply to the package format", async () => {
        const registry = new CheckRegistry([check("pptx-only", { formats: ["pptx"] })]);
        await expect(runChecks(ValidationContext.fromEntries({}), registry, { format: "docx", select: ["pptx-only"] })).rejects.toThrow(
            /does not apply to docx/,
        );
    });
});

describe("runChecks: a prerequisite that gates by part (#307)", () => {
    const parts = { "word/a.xml": DOCUMENT, "word/b.xml": DOCUMENT };
    // Reports which parts it could read.
    const reader = check("reader", {
        requires: ["gate"],
        run: (ctx) =>
            ["word/a.xml", "word/b.xml"].map((part): ValidationIssue => ({
                severity: "info",
                code: ctx.xml(part)?.ok ? "read" : "unreadable",
                message: ctx.xml(part)?.ok ? "read" : (ctx.xml(part) as { message: string }).message,
                path: part,
            })),
    });

    it("runs the dependents on the parts the gate passed, with the failed parts unparseable", async () => {
        const gate = check("gate", { gatesByPart: true, run: () => [error("broken", "word/a.xml")] });
        const ctx = ValidationContext.fromEntries(parts);
        const run = await runChecks(ctx, new CheckRegistry([gate, reader]), { format: "docx" });
        expect(run.executed).toEqual(["gate", "reader"]);
        expect(run.issues.filter((issue) => issue.severity === "info")).toEqual([
            expect.objectContaining({ code: "unreadable", path: "word/a.xml", message: expect.stringContaining("broken") }),
            expect.objectContaining({ code: "read", path: "word/b.xml" }),
        ]);
        expect(run.valid).toBe(false);
        // The shared context is not changed for checks outside the gate.
        expect(ctx.xml("word/a.xml")?.ok).toBe(true);
    });

    it("still skips the dependents when the failing gate names no part, or does not gate by part", async () => {
        for (const gate of [
            check("gate", { gatesByPart: true, run: () => [error("broken")] }),
            check("gate", { run: () => [error("broken", "word/a.xml")] }),
        ]) {
            const run = await runChecks(ValidationContext.fromEntries(parts), new CheckRegistry([gate, reader]), { format: "docx" });
            expect(run.skipped).toEqual(["reader"]);
        }
    });
});
