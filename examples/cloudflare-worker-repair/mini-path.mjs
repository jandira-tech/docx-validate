// Browser-native minimal POSIX path — plain JS, no node:path, no polyfill pkg.
function clean(parts, keepUp) { const up = []; for (const p of parts) { if (!p || p === ".") continue; if (p === "..") { if (up.length && up[up.length - 1] !== "..") up.pop(); else if (keepUp) up.push(".."); } else up.push(p); } return up; }
export function normalize(p) { const abs = p.startsWith("/"); const r = clean(p.split("/"), !abs).join("/"); return abs ? "/" + r : (r || "."); }
export function join(...segs) { return normalize(segs.filter((s) => s && s.length).join("/")); }
export function dirname(p) { const s = p.replace(/\/+$/, ""); const i = s.lastIndexOf("/"); if (i < 0) return "."; if (i === 0) return "/"; return s.slice(0, i); }
export function basename(p, ext) { let b = p.replace(/\/+$/, "").split("/").pop() || ""; if (ext && b.endsWith(ext)) b = b.slice(0, -ext.length); return b; }
export function extname(p) { const b = basename(p); const i = b.lastIndexOf("."); return i > 0 ? b.slice(i) : ""; }
export function relative(from, to) { const f = normalize(from).split("/").filter(Boolean); const t = normalize(to).split("/").filter(Boolean); let i = 0; while (i < f.length && i < t.length && f[i] === t[i]) i++; return [...f.slice(i).map(() => ".."), ...t.slice(i)].join("/"); }
export function resolve(...segs) { let res = ""; let abs = false; for (let i = segs.length - 1; i >= 0 && !abs; i--) { const s = segs[i]; if (!s) continue; res = s + "/" + res; abs = s.startsWith("/"); } res = normalize(res); return abs && !res.startsWith("/") ? "/" + res : res; }
export const sep = "/";
export const posix = { normalize, join, dirname, basename, extname, relative, resolve, sep };
export default { normalize, join, dirname, basename, extname, relative, resolve, sep, posix };
