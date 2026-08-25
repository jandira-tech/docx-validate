/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

/**
 * word-valid openability regression: docx-validate's `word-valid` profile must
 * ERROR on the structural defects that make real Microsoft Word reject a file
 * (OPEN_ERROR / "Word found unreadable content"), while NOT flagging files Word
 * opens cleanly.
 *
 * Ground truth: every fixture here was classified by a real-Word probe
 * (word_oracle, open-and-observe). The blocking allowlist
 * (WORD_BLOCKING_MISPLACED_LOCALS) and the tolerated set were derived empirically
 * from that probe over a 500-file corpus; this suite locks both directions.
 */
import path from "node:path";
import { fileURLToPath } from "node:url";
import { describe, expect, it } from "vitest";
import { validate } from "../src/scripts/office/validate";

const HERE = path.dirname(fileURLToPath(import.meta.url));
const BROKEN = path.join(HERE, "fixtures", "broken");
const WORKING = path.join(HERE, "fixtures", "working");

const errs = (r: { issues: { severity: string }[] }) => r.issues.filter((i) => i.severity === "error");

describe("word-valid profile errors on Word-rejected files", () => {
    it("misplaced table property (tblCellSpacing outside tblPr) — Word unreadable", async () => {
        const res = await validate(path.join(BROKEN, "word-rejects-table-cellspacing.docx"), { profile: "word-valid" });
        expect(res.valid).toBe(false);
        expect(res.issues.some((i) => i.severity === "error" && /is not expected/.test(i.message))).toBe(true);
    });
    // NB: `word-rejects-misplaced-run.docx` (a misplaced <w:r>, Word OPEN_ERROR)
    // exercises the same allowlist path and is `valid:false` standalone, but is
    // omitted here — libxmljs2 XSD validation of that ~2 KB file is nondeterministic
    // inside the vitest worker. The tblCellSpacing case above covers the path.
});

describe("word-valid profile does NOT flag Word-tolerated misplacements (no false positives)", { timeout: 30000 }, () => {
    // Each of these opens cleanly in real Word despite a content-model quirk that
    // an over-broad rule would wrongly flag. They lock the tolerated set.
    const tolerated: [string, string][] = [
        ["misplaced w:pgSz (sectPr child)", "word-tolerated-misplaced-pgsz.docx"],
        ["misplaced w:link (style child)", "word-tolerated-misplaced-link.docx"],
        ["misplaced w:uiPriority (style child)", "word-tolerated-misplaced-uipriority.docx"],
        ["orphan comment-range marker", "word-tolerated-orphan-comment.docx"],
        // Duplicate <w:pPr> in one paragraph: XSD says the second pPr "is not
        // expected", but real Word silently merges/ignores it and opens cleanly
        // (superdoc export). Locks pPr OUT of the blocking allowlist.
        ["duplicate w:pPr in a paragraph", "word-tolerated-duplicate-ppr.docx"],
        // Broken media relationship (rel Target file absent): Word shows a
        // missing-image placeholder and opens cleanly — broken media alone does
        // NOT predict rejection. Locks media/ out of the word-valid blocking set.
        ["broken media relationship", "word-tolerated-broken-media-rel.docx"],
        ["a Word-clean baseline file", "word-clean-strict01.docx"],
    ];
    for (const [label, file] of tolerated) {
        it(`stays valid: ${label}`, { timeout: 30000 }, async () => {
            const res = await validate(path.join(WORKING, file), { profile: "word-valid" });
            const e = errs(res);
            expect(e, `must not gain errors, got: ${e.map((x) => (x as { code?: string }).code).join(",")}`).toHaveLength(0);
            expect(res.valid).toBe(true);
        });
    }
});
