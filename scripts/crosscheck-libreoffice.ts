/*
 * Third-validator cross-check: LibreOffice (a real, independent Office engine).
 *
 * Microsoft Word is our ground-truth oracle (scripts/probe-word-fixtures.ts). This
 * script runs the SAME fixtures through LibreOffice headless to see whether a second
 * real Office consumer reproduces Word's accept/reject decision — i.e. whether the
 * Word-rejected files are universally broken or Word-uniquely-strict.
 *
 * LibreOffice "rejects" a file if a headless `--convert-to pdf` produces no PDF.
 *
 * Usage:
 *   tsx scripts/crosscheck-libreoffice.ts                # cross-check the 65 Word-rejected fixtures
 *   tsx scripts/crosscheck-libreoffice.ts --all          # cross-check every fixture in word-probe-all.jsonl
 */
import { promises as fs } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { execFile } from "node:child_process";
import { promisify } from "node:util";

const execFileAsync = promisify(execFile);
const HERE = path.dirname(fileURLToPath(import.meta.url));
const REPO = path.resolve(HERE, "..");
const SOFFICE = "/Applications/LibreOffice.app/Contents/MacOS/soffice";
const OUT = path.join(REPO, "validation-results", "libreoffice-crosscheck.jsonl");

/** Returns true if LibreOffice opened the file well enough to render a PDF. */
async function libreOfficeOpens(absFile: string, outDir: string): Promise<{ opens: boolean; detail: string }> {
    const pdf = path.join(outDir, `${path.basename(absFile, path.extname(absFile))}.pdf`);
    await fs.rm(pdf, { force: true });
    try {
        const { stderr } = await execFileAsync(
            SOFFICE,
            ["--headless", "--norestore", "-env:UserInstallation=file:///tmp/lo-xcheck-profile", "--convert-to", "pdf", "--outdir", outDir, absFile],
            { timeout: 60_000 },
        );
        const opened = await fs.stat(pdf).then(() => true, () => false);
        return { opens: opened, detail: opened ? "pdf-produced" : `no-pdf: ${stderr.trim().slice(0, 120)}` };
    } catch (e) {
        return { opens: false, detail: `error: ${(e as Error).message.slice(0, 120)}` };
    }
}

async function main(argv: readonly string[]): Promise<number> {
    const all = argv.includes("--all");
    const probe = (await fs.readFile(path.join(REPO, "validation-results", "word-probe-all.jsonl"), "utf8"))
        .trim()
        .split("\n")
        .filter(Boolean)
        .map((l) => JSON.parse(l) as { relativePath: string; word?: { clean: boolean; outcome: string } });

    // Default: just the files Word does NOT open cleanly. --all: everything.
    const targets = all ? probe : probe.filter((r) => r.word && !r.word.clean);
    const outDir = "/tmp/lo-xcheck-out";
    await fs.mkdir(outDir, { recursive: true });

    const records: unknown[] = [];
    let loOpensWordRejected = 0;
    let done = 0;
    for (const r of targets) {
        const abs = path.join(REPO, r.relativePath);
        if (!(await fs.stat(abs).then(() => true, () => false))) continue;
        const lo = await libreOfficeOpens(abs, outDir);
        const wordClean = r.word?.clean ?? null;
        if (wordClean === false && lo.opens) loOpensWordRejected += 1;
        records.push({ file: r.relativePath, word: r.word?.outcome ?? "?", libreOffice: lo.opens ? "opens" : "rejects", detail: lo.detail });
        if (++done % 10 === 0) process.stderr.write(`  ${done}/${targets.length}\n`);
    }

    await fs.writeFile(OUT, records.map((r) => JSON.stringify(r)).join("\n") + "\n");
    const wordRejectedCount = targets.filter((r) => r.word && !r.word.clean).length;
    process.stderr.write(`\nLibreOffice cross-check: ${done} files. Of ${wordRejectedCount} Word-rejected, LibreOffice OPENED ${loOpensWordRejected} and rejected ${wordRejectedCount - loOpensWordRejected}.\n`);
    process.stderr.write(`Wrote ${OUT}\n`);
    return 0;
}

main(process.argv.slice(2)).then((c) => process.exit(c));
