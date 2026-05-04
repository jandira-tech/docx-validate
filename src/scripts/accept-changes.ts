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

/**
 * Accept all tracked changes in a DOCX file using LibreOffice.
 *
 * 1:1 TypeScript port of `src/docx-validate/scripts/accept_changes.py`
 * (task #14). Requires LibreOffice (`soffice`) to be installed.
 *
 * Usage:
 *   bunx tsx scripts/accept-changes.ts <input.docx> <output.docx>
 */

import { promises as fs, existsSync } from "node:fs";
import path from "node:path";

import { Command } from "commander";

import { commanderExitCode, runCli } from "../lib/run-cli.ts";
import { getSofficeEnv, runSoffice } from "./office/soffice.ts";

export const LIBREOFFICE_PROFILE = "/tmp/libreoffice_docx_profile";
export const MACRO_DIR = `${LIBREOFFICE_PROFILE}/user/basic/Standard`;

export const ACCEPT_CHANGES_MACRO = `<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE script:module PUBLIC "-//OpenOffice.org//DTD OfficeDocument 1.0//EN" "module.dtd">
<script:module xmlns:script="http://openoffice.org/2000/script" script:name="Module1" script:language="StarBasic">
    Sub AcceptAllTrackedChanges()
        Dim document As Object
        Dim dispatcher As Object

        document = ThisComponent.CurrentController.Frame
        dispatcher = createUnoService("com.sun.star.frame.DispatchHelper")

        dispatcher.executeDispatch(document, ".uno:AcceptAllTrackedChanges", "", 0, Array())
        ThisComponent.store()
        ThisComponent.close(True)
    End Sub
</script:module>`;

export interface AcceptChangesResult {
    message: string;
    ok: boolean;
}

/**
 * Accept all tracked changes in `inputFile` and write the result to
 * `outputFile`. Returns a result object with a human-readable status
 * message; the public `acceptChanges()` wrapper throws on failure.
 */
export async function acceptChangesResult(inputFile: string, outputFile: string): Promise<AcceptChangesResult> {
    const inputPath = path.resolve(inputFile);
    const outputPath = path.resolve(outputFile);

    if (!existsSync(inputPath)) {
        return { ok: false, message: `Error: Input file not found: ${inputFile}` };
    }

    if (path.extname(inputPath).toLowerCase() !== ".docx") {
        return {
            ok: false,
            message: `Error: Input file is not a DOCX file: ${inputFile}`,
        };
    }

    try {
        await fs.mkdir(path.dirname(outputPath), { recursive: true });
        await fs.copyFile(inputPath, outputPath);
    } catch (err) {
        return {
            ok: false,
            message: `Error: Failed to copy input file to output location: ${err instanceof Error ? err.message : String(err)}`,
        };
    }

    if (!(await setupLibreofficeMacro())) {
        return { ok: false, message: "Error: Failed to setup LibreOffice macro" };
    }

    const args = [
        "--headless",
        `-env:UserInstallation=file://${LIBREOFFICE_PROFILE}`,
        "--norestore",
        "vnd.sun.star.script:Standard.Module1.AcceptAllTrackedChanges?language=Basic&location=application",
        outputPath,
    ];

    let result: { stdout: string; stderr: string; exitCode: number };
    try {
        result = await runSofficeWithTimeout(args, 30000);
    } catch (err) {
        if (err instanceof TimeoutError) {
            return {
                ok: false,
                message: `Error: LibreOffice timed out after 30 seconds while processing ${inputFile}`,
            };
        }
        throw err;
    }

    if (result.exitCode !== 0) {
        return {
            ok: false,
            message: `Error: LibreOffice failed: ${result.stderr}`,
        };
    }

    return {
        ok: true,
        message: `Successfully accepted all tracked changes: ${inputFile} -> ${outputFile}`,
    };
}

/**
 * Public API: accept all tracked changes in `inputFile`, writing the
 * result to `outputFile`. Throws on failure.
 */
export async function acceptChanges(inputFile: string, outputFile: string): Promise<void> {
    const result = await acceptChangesResult(inputFile, outputFile);
    if (!result.ok) {
        throw new Error(result.message);
    }
}

class TimeoutError extends Error {}

async function runSofficeWithTimeout(args: string[], timeoutMs: number): Promise<{ stdout: string; stderr: string; exitCode: number }> {
    // Aborting the controller signals runSoffice() to SIGTERM (then SIGKILL
    // after 1s) the spawned soffice — without this the timer would only race
    // the JS promise, leaving a runaway LibreOffice process after the timeout.
    const controller = new AbortController();
    return new Promise((resolve, reject) => {
        let settled = false;
        const timer = setTimeout(() => {
            if (settled) return;
            settled = true;
            controller.abort();
            reject(new TimeoutError(`soffice timed out after ${timeoutMs}ms`));
        }, timeoutMs);
        runSoffice(args, { signal: controller.signal }).then(
            (res) => {
                if (settled) return;
                settled = true;
                clearTimeout(timer);
                resolve(res);
            },
            (err) => {
                if (settled) return;
                settled = true;
                clearTimeout(timer);
                reject(err);
            },
        );
    });
}

/**
 * Ensure the StarBasic macro that performs `AcceptAllTrackedChanges`
 * lives in the dedicated LibreOffice profile. Mirrors the Python
 * `_setup_libreoffice_macro` helper verbatim — including the initial
 * `--terminate_after_init` warm-up that creates the profile skeleton.
 */
export async function setupLibreofficeMacro(): Promise<boolean> {
    const macroFile = path.join(MACRO_DIR, "Module1.xba");

    if (existsSync(macroFile)) {
        try {
            const text = await fs.readFile(macroFile, "utf-8");
            if (text.includes("AcceptAllTrackedChanges")) {
                return true;
            }
        } catch {
            // fall through and rewrite
        }
    }

    if (!existsSync(MACRO_DIR)) {
        try {
            await runSofficeWithTimeout(
                ["--headless", `-env:UserInstallation=file://${LIBREOFFICE_PROFILE}`, "--terminate_after_init"],
                10000,
            );
        } catch {
            // Mirrors Python: warm-up failure is non-fatal; we just continue
            // and try to mkdir/write the macro file. If that also fails the
            // caller gets `false` below.
        }
        try {
            await fs.mkdir(MACRO_DIR, { recursive: true });
        } catch {
            /* tolerated — write below will surface the real error */
        }
    }

    try {
        await fs.writeFile(macroFile, ACCEPT_CHANGES_MACRO, "utf-8");
        return true;
    } catch (err) {
        process.stderr.write(`Failed to setup LibreOffice macro: ${err instanceof Error ? err.message : String(err)}\n`);
        return false;
    }
}

export function buildCommand(): Command {
    const cmd = new Command();
    cmd.name("accept-changes")
        .description("Accept all tracked changes in a DOCX file")
        .argument("<input_file>", "Input DOCX file with tracked changes")
        .argument("<output_file>", "Output DOCX file (clean, no tracked changes)");
    return cmd;
}

export async function runFromArgv(argv: readonly string[]): Promise<number> {
    const cmd = buildCommand();
    cmd.exitOverride();
    // Catch CommanderError (missing args, invalid options, --help) so the CLI
    // returns a clean exit code instead of bubbling to runCli().
    try {
        cmd.parse(argv as string[], { from: "user" });
    } catch (err) {
        return commanderExitCode(err);
    }
    const [inputFile, outputFile] = cmd.args;

    const result = await acceptChangesResult(inputFile, outputFile);
    process.stdout.write(`${result.message}\n`);
    return result.message.includes("Error") ? 1 : 0;
}

// Internal exports for tests.
export const __test = {
    getSofficeEnv,
    runSofficeWithTimeout,
    TimeoutError,
};

runCli(import.meta.url, () => runFromArgv(process.argv.slice(2)));
