# Security

## Reporting a vulnerability

Email: `arthur@jandira.tech`. Please do not file a public GitHub issue
for vulnerabilities.

## About the LibreOffice shim (informational)

If you came here from an automated supply-chain alert about
`LD_PRELOAD` / runtime native code generation: this section explains the
pattern, where it lives, and why it is **not** included in the published
npm package.

### What the alert is flagging

`src/scripts/office/soffice.ts` defines a small C source string
(`SHIM_SOURCE`) and an `ensureShim()` helper that:

1. Writes the C source to `/tmp/lo_socket_shim.c`
2. Compiles it with `gcc -shared -fPIC` to `/tmp/lo_socket_shim.so`
3. Sets `LD_PRELOAD` to that `.so` when spawning `soffice`
   (LibreOffice headless)

The shim interposes `socket()`, `socketpair()`, `listen()`, `accept()`,
`close()`, and `read()` so LibreOffice can boot in sandboxed Linux VMs
where the kernel blocks `AF_UNIX` socket creation. The `close()` hook
calls `_exit(0)` once the listener FD is closed so soffice does not stay
alive waiting for a second client connection after the macro has
finished.

This pattern legitimately matches the shape of supply-chain malware
(runtime gcc + `LD_PRELOAD` + syscall interposition) — that is exactly
what generic AI scanners are trained to flag, and they are correct to
flag _unknown_ packages that do this.

### Why it is in the repo

LibreOffice's CLI macro flow (`acceptChanges()` in
`src/scripts/accept-changes.ts`) is the only part of `docx-validate`
that needs LibreOffice. In container images and sandboxed VMs (the
environments where this library tends to run), `AF_UNIX` is sometimes
restricted, and `soffice` aborts on startup. The shim is a stopgap that
keeps `acceptChanges()` working in those environments.

### Where it does NOT live

- **The published npm bundle (`dist/index.mjs`)** — `src/index.ts` does
  not re-export `soffice.ts` or `accept-changes.ts`, so the bundler
  tree-shakes both files (and `SHIM_SOURCE`, `ensureShim`,
  `getSofficeEnv`, `runSoffice`, `acceptChanges`, etc.) out of the
  published artifact. `grep -E 'LD_PRELOAD|SHIM_SOURCE|RTLD_NEXT|dlsym'
dist/index.mjs` returns zero matches.
- **The default programmatic API** — npm consumers doing
  `import { ... } from "docx-validate"` cannot reach the shim or
  anything that depends on it.

### Where it does live

- The source files in `src/scripts/office/soffice.ts` and
  `src/scripts/accept-changes.ts` ship in the GitHub repo for
  developers and CI users who run the CLI scripts directly via
  `bunx tsx src/scripts/...`.
- They are also exercised by `tests/soffice.test.ts` and
  `tests/accept-changes.test.ts` under
  `SOFFICE_AVAILABLE=1 bun run test`.

### If your scanner still flags the package

Verify with:

```bash
npm pack docx-validate
tar -xf docx-validate-*.tgz
grep -RE "LD_PRELOAD|SHIM_SOURCE|RTLD_NEXT|dlsym|VCLPLUGIN" package/
# (should produce no output)
```

If your scanner reports a hit on the published artifact (not the
GitHub source), please open an issue with the diagnostic output — that
would be a regression we want to fix.
