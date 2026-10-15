#!/usr/bin/env node

import { execFile } from "node:child_process";
import { promisify } from "node:util";
import { publicationPlan } from "./prerelease";

const [tag, ...extra] = process.argv.slice(2);
if (!tag || extra.length > 0) throw new Error("Usage: pnpm run release:publish v0.53.0-rc.0");
publicationPlan(tag.slice(1), tag);
await promisify(execFile)("gh", [
    "workflow",
    "run",
    "publish.yml",
    "--repo",
    "jandira-tech/docx-validate",
    "--ref",
    "main",
    "-f",
    `tag=${tag}`,
]);
process.stdout.write(
    `Dispatched publish.yml for ${tag}. Follow the run at https://github.com/jandira-tech/docx-validate/actions/workflows/publish.yml\n`,
);
