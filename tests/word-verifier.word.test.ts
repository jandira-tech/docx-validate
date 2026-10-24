/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

/**
 * Optional real-Word integration lane for `src/lib/word-verifier.ts`.
 *
 * SKIPPED unless `WORD_VERIFY=1` and the platform is macOS with Microsoft
 * Word installed — this opens the real application and requires
 * Accessibility automation permission for the host terminal. Run with:
 *
 *   WORD_VERIFY=1 pnpm exec vitest run tests/word-verifier.word.test.ts
 */

import { readFileSync } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

import { describe, expect, it } from "vitest";

import { createMacWordTransport, verifyInWord } from "../src/lib/word-verifier";
import { buildDocx } from "./support/minimal-docx";

const RUN_WORD = process.env.WORD_VERIFY === "1" && process.platform === "darwin";

// Malformed XML in word/document.xml; the manifest records Word's open error for it.
const OPEN_ERROR_SPECIMEN = path.join(path.dirname(fileURLToPath(import.meta.url)), "fixtures/broken/single-paragraph.malformed-xml.docx");

describe("verifyInWord (real Microsoft Word)", () => {
    it.skipIf(!RUN_WORD)(
        "opens the minimal docx cleanly in Word",
        async () => {
            const bytes = await buildDocx();
            const result = await verifyInWord(bytes, { transport: createMacWordTransport() });
            expect(result).toMatchObject({
                requested: true,
                outcome: "clean-open",
                satisfied: true,
                platform: "darwin",
            });
            expect(result.fileSha256).toMatch(/^[0-9a-f]{64}$/);
        },
        120_000,
    );

    it.skipIf(!RUN_WORD)(
        "clears the dialog its own open raised, so the next run is attributable (#315)",
        async () => {
            const transport = createMacWordTransport();
            const rejected = await verifyInWord(readFileSync(OPEN_ERROR_SPECIMEN), { transport });
            expect(rejected.outcome, rejected.details).toBe("open-error");
            const clean = await verifyInWord(await buildDocx(), { transport });
            expect(clean.outcome, clean.details).toBe("clean-open");
        },
        240_000,
    );
});
