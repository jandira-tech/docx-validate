#!/usr/bin/env node
/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

import { execFile, spawn } from "node:child_process";
import { promises as fs } from "node:fs";
import path from "node:path";
import { fileURLToPath, pathToFileURL } from "node:url";
import { promisify } from "node:util";

const exec = promisify(execFile);
const REPO = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");

export function publicationPlan(
    version: string,
    tag?: string,
    releaseKind?: string,
): { version: string; tag: string; distTag: "next" | "latest" } {
    const match = /^(0|[1-9]\d*)\.(0|[1-9]\d*)\.(0|[1-9]\d*)(?:-([0-9A-Za-z-]+(?:\.[0-9A-Za-z-]+)*))?$/.exec(version);
    if (!match || match[4]?.split(".").some((part) => /^\d+$/.test(part) && part.length > 1 && part.startsWith("0"))) {
        throw new Error(`Invalid release version: ${version}`);
    }
    const expectedTag = `v${version}`;
    if (tag !== undefined && tag !== expectedTag) throw new Error(`Release tag ${tag} does not match package version ${version}`);
    const prerelease = match[4] !== undefined;
    if (releaseKind && releaseKind !== (prerelease ? "prerelease" : "stable")) {
        throw new Error("GitHub release prerelease status does not match the package version");
    }
    return { version, tag: expectedTag, distTag: prerelease ? "next" : "latest" };
}

export function verifyPackedFiles(files: readonly string[]): void {
    for (const required of ["dist/index.mjs", "dist/index.d.mts", "dist/schemas/ISO-IEC29500-4_2016/wml.xsd"]) {
        if (!files.includes(required)) throw new Error(`Package is missing ${required}`);
    }
    if (files.some((file) => /^(?:src|tests|node_modules)\//.test(file))) {
        throw new Error("Package unexpectedly contains source, tests, or node_modules");
    }
}

async function run(command: string, args: readonly string[]): Promise<void> {
    process.stdout.write(`\n> ${command} ${args.join(" ")}\n`);
    await new Promise<void>((resolve, reject) => {
        const child = spawn(command, [...args], { cwd: REPO, stdio: "inherit" });
        child.on("error", reject);
        child.on("close", (code) => (code === 0 ? resolve() : reject(new Error(`${command} exited with ${code}`))));
    });
}

export async function prerelease(): Promise<void> {
    const pkg = JSON.parse(await fs.readFile(path.join(REPO, "package.json"), "utf8")) as { name: string; version: string };
    if (pkg.name !== "docx-validate") throw new Error("Unexpected package name");
    const plan = publicationPlan(pkg.version, process.env.RELEASE_TAG || undefined, process.env.RELEASE_KIND || undefined);
    if (process.env.RELEASE_TAG) {
        const [head, tagged] = await Promise.all([
            exec("git", ["rev-parse", "HEAD"], { cwd: REPO }),
            exec("git", ["rev-parse", "--verify", `refs/tags/${plan.tag}^{commit}`], { cwd: REPO }),
        ]);
        if (head.stdout.trim() !== tagged.stdout.trim()) throw new Error("Checkout is not the release tag's commit");
    }
    await run("pnpm", ["exec", "tsc", "--noEmit"]);
    await run("pnpm", ["run", "lint"]);
    await run("pnpm", ["run", "check"]);
    await run("pnpm", ["run", "cspell"]);
    await run("pnpm", ["exec", "vitest", "run", "--coverage", "--maxWorkers=2"]);
    await run("pnpm", ["run", "build"]);

    const bundle = await fs.readFile(path.join(REPO, "dist/index.mjs"), "utf8");
    if (/LD_PRELOAD|SHIM_SOURCE|RTLD_NEXT|dlsym/.test(bundle))
        throw new Error("LibreOffice shim unexpectedly appears in the package bundle");
    const packed = await exec("npm", ["pack", "--dry-run", "--json", "--ignore-scripts"], { cwd: REPO, maxBuffer: 4 * 1024 * 1024 });
    const metadata = JSON.parse(packed.stdout) as { name: string; version: string; files: { path: string }[] }[];
    const tarball = metadata[0];
    if (metadata.length !== 1 || tarball?.name !== pkg.name || tarball.version !== pkg.version)
        throw new Error("Packed package identity does not match package.json");
    verifyPackedFiles(tarball.files.map((file) => file.path));
    if (process.env.GITHUB_OUTPUT) {
        await fs.appendFile(process.env.GITHUB_OUTPUT, `version=${plan.version}\ndist_tag=${plan.distTag}\n`);
    }
    process.stdout.write(`\nRelease checks passed: ${pkg.name}@${plan.version} → npm tag ${plan.distTag}\n`);
}

if (process.argv[1] && import.meta.url === pathToFileURL(path.resolve(process.argv[1])).href) {
    prerelease().catch((err: unknown) => {
        process.stderr.write(`${err instanceof Error ? err.message : String(err)}\n`);
        process.exitCode = 1;
    });
}
