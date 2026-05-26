/*
 * Copyright 2026 Jandira Technologies, LLC
 * Licensed under the Apache License, Version 2.0.
 */

// pathlite — a tiny, dependency-free POSIX path helper.
//
// The validators address OOXML parts with `/`-separated, package-relative paths
// (that's how they live inside the zip), so POSIX semantics are exactly right
// regardless of host OS. Replacing `node:path` with this lets the validator
// module graph bundle for the browser (no `node:` import). DiskPartFS — the only
// genuinely disk-bound code — keeps using `node:path` directly.

function normalizeArray(parts: string[], allowAboveRoot: boolean): string[] {
    const res: string[] = [];
    for (const p of parts) {
        if (!p || p === ".") continue;
        if (p === "..") {
            if (res.length && res[res.length - 1] !== "..") res.pop();
            else if (allowAboveRoot) res.push("..");
        } else {
            res.push(p);
        }
    }
    return res;
}

export const sep = "/";

export function normalize(p: string): string {
    const isAbsolute = p.startsWith("/");
    const trailing = p.length > 1 && p.endsWith("/");
    let out = normalizeArray(p.split("/"), !isAbsolute).join("/");
    if (!out && !isAbsolute) out = ".";
    if (out && trailing) out += "/";
    return (isAbsolute ? "/" : "") + out;
}

export function join(...parts: Array<string>): string {
    const joined = parts.filter((p) => p && p.length > 0).join("/");
    return joined ? normalize(joined) : ".";
}

export function dirname(p: string): string {
    if (p.length === 0) return ".";
    const hadLeading = p.startsWith("/");
    const segments = p.split("/").filter((s) => s.length > 0);
    segments.pop();
    if (segments.length === 0) return hadLeading ? "/" : ".";
    return (hadLeading ? "/" : "") + segments.join("/");
}

export function basename(p: string, ext?: string): string {
    const segments = p.split("/").filter((s) => s.length > 0);
    let base = segments.length ? segments[segments.length - 1] : "";
    if (ext && base.endsWith(ext) && base !== ext) base = base.slice(0, -ext.length);
    return base;
}

export function extname(p: string): string {
    const base = basename(p);
    const dot = base.lastIndexOf(".");
    return dot <= 0 ? "" : base.slice(dot);
}

export function resolve(...parts: Array<string>): string {
    let resolved = "";
    let isAbsolute = false;
    for (let i = parts.length - 1; i >= 0 && !isAbsolute; i--) {
        const p = parts[i];
        if (!p) continue;
        resolved = `${p}/${resolved}`;
        isAbsolute = p.startsWith("/");
    }
    const normalized = normalizeArray(resolved.split("/"), !isAbsolute).join("/");
    if (isAbsolute) return `/${normalized}` || "/";
    return normalized || ".";
}

export function relative(from: string, to: string): string {
    const f = resolve(from).split("/").filter(Boolean);
    const t = resolve(to).split("/").filter(Boolean);
    let i = 0;
    while (i < f.length && i < t.length && f[i] === t[i]) i++;
    const up = f.slice(i).map(() => "..");
    return [...up, ...t.slice(i)].join("/");
}

const path = { sep, normalize, join, dirname, basename, extname, resolve, relative };
export default path;
