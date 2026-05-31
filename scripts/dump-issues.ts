/**
 * Parity oracle — dump every ValidationIssue (code/severity/message/path) that
 * docx-validate produces for a file at a given profile, as JSON on stdout.
 *
 * Used by python-jubarte's pyvalidator parity tests to assert exact
 * code+message+severity parity. Run with the local tsx (NOT npx).
 *
 *   node_modules/.bin/tsx scripts/dump-issues.ts <file.docx> [profile]
 */
import { validate } from "../src/scripts/office/validate";
import type { Profile } from "../src/lib/types";

async function main(): Promise<void> {
    const [, , target, profileArg] = process.argv;
    if (!target) {
        process.stderr.write("usage: dump-issues.ts <file.docx> [profile]\n");
        process.exit(2);
    }
    const profile = (profileArg ?? "lenient") as Profile;
    const result = await validate(target, { profile });
    const issues = (result.issues ?? []).map((i) => ({
        severity: i.severity,
        message: i.message,
        path: i.path ?? null,
        code: i.code ?? null,
    }));
    process.stdout.write(JSON.stringify({ valid: result.valid, issues }));
}

main().catch((e: unknown) => {
    process.stderr.write(`${e instanceof Error ? e.stack : String(e)}\n`);
    process.exit(2);
});
