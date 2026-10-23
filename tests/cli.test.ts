/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

/**
 * Unit surface of `src/cli.ts` (#314): the packaged validator entry must run
 * `runValidateFromArgv` only when it is the process entry point — importing
 * it as a module is inert. The integration twin that exercises the emitted
 * `dist/cli.mjs` lives in tests/package-cli.integration.test.ts.
 */

import { fileURLToPath } from "node:url";

import { afterEach, describe, expect, it, vi } from "vitest";

const runValidateFromArgv = vi.fn<(argv: readonly string[]) => Promise<number>>().mockResolvedValue(0);
vi.doMock("../src/scripts/office/validate", () => ({ runValidateFromArgv }));

const importCli = async (): Promise<void> => {
    vi.resetModules();
    await import("../src/cli");
};

describe("src/cli.ts entry-point guard", () => {
    afterEach(() => {
        vi.restoreAllMocks();
        runValidateFromArgv.mockClear();
    });

    it("runs the validator CLI when src/cli.ts is the process entry", async () => {
        const argv1 = process.argv[1];
        process.argv[1] = fileURLToPath(new URL("../src/cli.ts", import.meta.url));
        const exit = vi.spyOn(process, "exit").mockImplementation((() => undefined) as never);
        try {
            await importCli();
            await vi.waitFor(() => expect(runValidateFromArgv).toHaveBeenCalledTimes(1));
            expect(runValidateFromArgv).toHaveBeenCalledWith(process.argv.slice(2));
            expect(exit).toHaveBeenCalledWith(0);
        } finally {
            process.argv[1] = argv1;
        }
    });

    it("is inert when imported as a module", async () => {
        const argv1 = process.argv[1];
        process.argv[1] = fileURLToPath(new URL("./cli.test.ts", import.meta.url));
        try {
            await importCli();
            await new Promise((resolve) => setTimeout(resolve, 10));
            expect(runValidateFromArgv).not.toHaveBeenCalled();
        } finally {
            process.argv[1] = argv1;
        }
    });
});
