import { promises as fs } from "node:fs";
import path from "node:path";
import { performance } from "node:perf_hooks";
import { pathToFileURL } from "node:url";
import { withTempDir } from "../src/lib/run-cli";
import type { DOCXSchemaValidator } from "../src/scripts/office/validators/docx";

const ns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
async function main() {
    const [before, after] = process.argv.slice(2);
    if (!before || !after) throw new Error("Usage: tsx scripts/bench-redlining.ts <before docx.ts> <after docx.ts>");
    const Old = (await import(pathToFileURL(path.resolve(before)).href)).DOCXSchemaValidator as typeof DOCXSchemaValidator;
    const New = (await import(pathToFileURL(path.resolve(after)).href)).DOCXSchemaValidator as typeof DOCXSchemaValidator;
    for (const size of [200, 800])
        await withTempDir(async (dir) => {
            await fs.mkdir(path.join(dir, "word"));
            const text = "<w:r><w:t>bad</w:t><w:delText>bad</w:delText></w:r>";
            const body =
                "<w:ins>" +
                "<w:sdt>".repeat(size) +
                "<w:del>" +
                text.repeat(size) +
                "</w:del>" +
                text.repeat(size) +
                "</w:sdt>".repeat(size) +
                "</w:ins>";
            await fs.writeFile(path.join(dir, "word/document.xml"), `<w:document xmlns:w="${ns}"><w:body>${body}</w:body></w:document>`);
            for (const [name, Ctor] of [
                ["parent", Old],
                ["dfs", New],
            ] as const) {
                const v = new Ctor({ unpackedDir: dir });
                const samples: number[] = [];
                for (let i = 0; i < 6; i++) {
                    const start = performance.now();
                    const d = await v.validateDeletions();
                    const ins = await v.validateInsertions();
                    if (i) samples.push(performance.now() - start);
                    if (i === 5)
                        console.log(
                            JSON.stringify({
                                size,
                                name,
                                deletions: d.issues.length,
                                insertions: ins.issues.length,
                                ms: samples,
                                mean: samples.reduce((a, b) => a + b) / samples.length,
                            }),
                        );
                }
            }
        });
}
main();
