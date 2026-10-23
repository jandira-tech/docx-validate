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

import { describe, expect, it } from "vitest";

import { createMacWordTransport, verifyInWord } from "../src/lib/word-verifier";
import { buildDocx } from "./support/minimal-docx";

const RUN_WORD = process.env.WORD_VERIFY === "1" && process.platform === "darwin";

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
});
