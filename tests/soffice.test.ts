/*
 * Copyright 2026 Jandira Technologies, LLC
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *     http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */

import { spawnSync } from "node:child_process";
import { existsSync, unlinkSync } from "node:fs";

import { afterEach, beforeEach, describe, expect, it } from "vitest";

import { __test, ensureShim, getSofficeEnv, needsShim, runSoffice } from "../src/scripts/office/soffice";

const isLinux = process.platform === "linux";
const hasGcc = isLinux && spawnSync("gcc", ["--version"], { stdio: "ignore" }).status === 0;
const hasSoffice = Boolean(process.env.SOFFICE_AVAILABLE);

function clearShim(): void {
    for (const p of [__test.SHIM_SO, __test.SHIM_C]) {
        if (existsSync(p)) {
            try {
                unlinkSync(p);
            } catch {
                /* best effort */
            }
        }
    }
}

describe("needsShim", () => {
    it("returns false on non-Linux platforms", () => {
        if (isLinux) {
            // On Linux it depends on the host; we just assert it returns a bool.
            expect(typeof needsShim()).toBe("boolean");
        } else {
            expect(needsShim()).toBe(false);
        }
    });
});

describe("ensureShim", () => {
    if (!hasGcc) {
        it.skip("requires gcc + Linux to build the shim", () => undefined);
        return;
    }

    beforeEach(() => {
        clearShim();
    });
    afterEach(() => {
        clearShim();
    });

    it("compiles the shim and caches the .so", () => {
        const so = ensureShim();
        expect(so).toBe(__test.SHIM_SO);
        expect(existsSync(so)).toBe(true);

        // Second call should hit the cache (no recompile, but still returns path).
        const so2 = ensureShim();
        expect(so2).toBe(so);
        expect(existsSync(so2)).toBe(true);
    });

    it("removes the .c source after a successful build", () => {
        ensureShim();
        expect(existsSync(__test.SHIM_C)).toBe(false);
    });
});

describe("getSofficeEnv", () => {
    it("always sets SAL_USE_VCLPLUGIN=svp", () => {
        const env = getSofficeEnv({});
        expect(env.SAL_USE_VCLPLUGIN).toBe("svp");
    });

    it("preserves caller-provided env vars", () => {
        const env = getSofficeEnv({ FOO: "bar" });
        expect(env.FOO).toBe("bar");
        expect(env.SAL_USE_VCLPLUGIN).toBe("svp");
    });

    it("does not set LD_PRELOAD on non-Linux platforms", () => {
        if (isLinux) {
            // skip — behaviour is host-dependent on Linux
            return;
        }
        const env = getSofficeEnv({});
        expect(env.LD_PRELOAD).toBeUndefined();
    });
});

describe("SHIM_SOURCE", () => {
    it("contains the expected hook symbols", () => {
        expect(__test.SHIM_SOURCE).toContain("__attribute__((constructor))");
        expect(__test.SHIM_SOURCE).toContain("RTLD_NEXT");
        expect(__test.SHIM_SOURCE).toContain("AF_UNIX");
        // Ensure all six syscalls intercepted by the Python original are present.
        for (const sym of ["real_socket", "real_socketpair", "real_listen", "real_accept", "real_close", "real_read"]) {
            expect(__test.SHIM_SOURCE).toContain(sym);
        }
    });
});

describe("runSoffice", () => {
    if (!hasSoffice) {
        it.skip("requires SOFFICE_AVAILABLE=1 to invoke real soffice", () => undefined);
        return;
    }

    it("returns a help banner via --help", async () => {
        const result = await runSoffice(["--help"]);
        expect(result.exitCode).toBe(0);
        expect(result.stdout.length + result.stderr.length).toBeGreaterThan(0);
    }, 30000);
});
