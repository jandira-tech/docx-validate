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

import { existsSync, promises as fs } from "node:fs";
import * as os from "node:os";
import path from "node:path";

import { afterEach, beforeEach, describe, expect, it } from "vitest";

import {
    ACCEPT_CHANGES_MACRO,
    acceptChanges,
    acceptChangesResult,
    buildAcceptChangesCommand,
    LIBREOFFICE_PROFILE,
    MACRO_DIR,
    setupLibreofficeMacro,
} from "../src/scripts/accept-changes";

const hasSoffice = Boolean(process.env.SOFFICE_AVAILABLE);

async function makeTmpDir(prefix: string): Promise<string> {
    return fs.mkdtemp(path.join(os.tmpdir(), prefix));
}

describe("ACCEPT_CHANGES_MACRO constant", () => {
    it("declares the StarBasic Module1 with the AcceptAllTrackedChanges sub", () => {
        expect(ACCEPT_CHANGES_MACRO).toContain('script:name="Module1"');
        expect(ACCEPT_CHANGES_MACRO).toContain('script:language="StarBasic"');
        expect(ACCEPT_CHANGES_MACRO).toContain("Sub AcceptAllTrackedChanges()");
        expect(ACCEPT_CHANGES_MACRO).toContain(".uno:AcceptAllTrackedChanges");
        expect(ACCEPT_CHANGES_MACRO).toContain("ThisComponent.store()");
        expect(ACCEPT_CHANGES_MACRO).toContain("ThisComponent.close(True)");
    });
});

describe("LIBREOFFICE_PROFILE / MACRO_DIR", () => {
    it("matches the Python paths verbatim", () => {
        expect(LIBREOFFICE_PROFILE).toBe("/tmp/libreoffice_docx_profile");
        expect(MACRO_DIR).toBe("/tmp/libreoffice_docx_profile/user/basic/Standard");
    });
});

describe("acceptChangesResult validation", () => {
    it("returns Error when the input file does not exist", async () => {
        const result = await acceptChangesResult("/definitely/does/not/exist.docx", "/tmp/out.docx");
        expect(result.ok).toBe(false);
        expect(result.message).toMatch(/Input file not found/);
    });

    it("returns Error when the input file is not .docx", async () => {
        const dir = await makeTmpDir("accept-changes-test-");
        try {
            const wrong = path.join(dir, "input.txt");
            await fs.writeFile(wrong, "not a docx");
            const result = await acceptChangesResult(wrong, path.join(dir, "out.docx"));
            expect(result.ok).toBe(false);
            expect(result.message).toMatch(/not a DOCX file/);
        } finally {
            await fs.rm(dir, { recursive: true, force: true });
        }
    });
});

describe("acceptChanges (throwing wrapper)", () => {
    it("throws on validation errors", async () => {
        await expect(acceptChanges("/definitely/does/not/exist.docx", "/tmp/out.docx")).rejects.toThrow(/Input file not found/);
    });
});

describe("buildAcceptChangesCommand", () => {
    it("declares two positional args (input + output)", () => {
        const cmd = buildAcceptChangesCommand();
        expect(cmd.name()).toBe("accept-changes");
        // commander stores positional definitions on `_args`; fall back
        // through `args` if internals change.
        const args = (cmd as unknown as { _args: { name: () => string }[] })._args;
        expect(args.map((a) => a.name())).toEqual(["input_file", "output_file"]);
    });
});

describe("setupLibreofficeMacro (smoke test)", () => {
    // We re-point MACRO_DIR/macroFile at a private temp tree by patching
    // the module's behaviour through a wrapper — but since the constants
    // are exported as plain strings, we test the real macro setup against
    // an isolated fake by writing into the *real* MACRO_DIR location only
    // when we explicitly opt in. By default we only check that the macro
    // file ends up containing ACCEPT_CHANGES_MACRO without spawning soffice.

    let savedMacro: string | null = null;
    let savedExisted = false;

    beforeEach(async () => {
        if (existsSync(MACRO_DIR)) {
            savedExisted = true;
            const macroFile = path.join(MACRO_DIR, "Module1.xba");
            if (existsSync(macroFile)) {
                savedMacro = await fs.readFile(macroFile, "utf-8");
            } else {
                savedMacro = null;
            }
        } else {
            savedExisted = false;
            savedMacro = null;
        }
    });

    afterEach(async () => {
        // Restore prior state to keep the host's profile untouched.
        if (!savedExisted) {
            // We may have created the dir tree; remove the parent profile
            // only if we created it — safest is to leave it alone, but we
            // do remove the macro file we wrote.
            const macroFile = path.join(MACRO_DIR, "Module1.xba");
            if (existsSync(macroFile)) {
                await fs.rm(macroFile, { force: true });
            }
        } else if (savedMacro !== null) {
            const macroFile = path.join(MACRO_DIR, "Module1.xba");
            await fs.writeFile(macroFile, savedMacro, "utf-8");
        }
    });

    it("writes the macro file when MACRO_DIR is writable", async () => {
        if (!hasSoffice && !existsSync(MACRO_DIR)) {
            // The function will try to spawn soffice to bootstrap the dir.
            // Without soffice that fails silently, then the mkdir fallback
            // runs. If /tmp is writable we should still end up with the
            // macro file.
        }
        const ok = await setupLibreofficeMacro();
        expect(ok).toBe(true);
        const macroFile = path.join(MACRO_DIR, "Module1.xba");
        expect(existsSync(macroFile)).toBe(true);
        const written = await fs.readFile(macroFile, "utf-8");
        expect(written).toContain("AcceptAllTrackedChanges");
        expect(written).toBe(ACCEPT_CHANGES_MACRO);
    });

    it("is idempotent — second call returns true without rewriting", async () => {
        const ok1 = await setupLibreofficeMacro();
        expect(ok1).toBe(true);
        const macroFile = path.join(MACRO_DIR, "Module1.xba");
        const stat1 = await fs.stat(macroFile);
        // Bump mtime granularity by waiting a hair, then re-run.
        await new Promise((r) => setTimeout(r, 5));
        const ok2 = await setupLibreofficeMacro();
        expect(ok2).toBe(true);
        const stat2 = await fs.stat(macroFile);
        expect(stat2.mtimeMs).toBe(stat1.mtimeMs);
    });
});

describe("acceptChanges (real soffice)", () => {
    // Tracked-changes fixture lives in the jubarte test corpus (101 insertions,
    // 69 deletions). We reference it directly rather than duplicating it.
    const FIXTURE = path.resolve(__dirname, "../../jubarte/tests/fixtures/potpourritest.docx");
    const FIXTURE_PRESENT = existsSync(FIXTURE);

    if (!hasSoffice) {
        it.skip("requires SOFFICE_AVAILABLE=1 to run end-to-end", () => undefined);
        return;
    }

    it.skipIf(!FIXTURE_PRESENT)(
        "accepts tracked changes in a real DOCX",
        async () => {
            const dir = await makeTmpDir("accept-changes-e2e-");
            try {
                const out = path.join(dir, "out.docx");
                await acceptChanges(FIXTURE, out);
                expect(existsSync(out)).toBe(true);
                const buf = await fs.readFile(out);
                expect(buf.length).toBeGreaterThan(0);
            } finally {
                await fs.rm(dir, { recursive: true, force: true });
            }
        },
        60000,
    );
});
