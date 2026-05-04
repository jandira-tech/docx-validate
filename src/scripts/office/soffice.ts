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
 * Helper for running LibreOffice (soffice) in environments where AF_UNIX
 * sockets may be blocked (e.g., sandboxed VMs). Detects the restriction
 * at runtime and applies an LD_PRELOAD shim if needed.
 *
 * Port of src/docx-validate/scripts/office/soffice.py (task #6).
 *
 * On non-Linux platforms the shim is skipped — matching the Python
 * behaviour of relying on `gcc` + `LD_PRELOAD`, both Linux-only.
 */

import { spawn, spawnSync } from "node:child_process";
import { existsSync, unlinkSync, writeFileSync } from "node:fs";
import * as os from "node:os";
import * as path from "node:path";

import { runCli } from "../../lib/run-cli.ts";

export interface RunSofficeOptions {
    env?: NodeJS.ProcessEnv;
    cwd?: string;
    /**
     * If provided, aborting this signal terminates the spawned soffice
     * process (SIGTERM, then SIGKILL after 1s if it has not exited).
     * Used by `runSofficeWithTimeout()` so the caller does not leak a
     * zombie process when the timer wins the race.
     */
    signal?: AbortSignal;
}

export interface RunSofficeResult {
    stdout: string;
    stderr: string;
    exitCode: number;
}

const SHIM_SO = path.join(os.tmpdir(), "lo_socket_shim.so");
const SHIM_C = path.join(os.tmpdir(), "lo_socket_shim.c");

const SHIM_SOURCE = String.raw`
#define _GNU_SOURCE
#include <dlfcn.h>
#include <errno.h>
#include <signal.h>
#include <stdio.h>
#include <stdlib.h>
#include <sys/socket.h>
#include <unistd.h>

static int (*real_socket)(int, int, int);
static int (*real_socketpair)(int, int, int, int[2]);
static int (*real_listen)(int, int);
static int (*real_accept)(int, struct sockaddr *, socklen_t *);
static int (*real_close)(int);
static int (*real_read)(int, void *, size_t);

/* Per-FD bookkeeping (FDs >= 1024 are passed through unshimmed). */
static int is_shimmed[1024];
static int peer_of[1024];
static int wake_r[1024];            /* accept() blocks reading this */
static int wake_w[1024];            /* close()  writes to this      */
static int listener_fd = -1;        /* FD that received listen()    */

__attribute__((constructor))
static void init(void) {
    real_socket     = dlsym(RTLD_NEXT, "socket");
    real_socketpair = dlsym(RTLD_NEXT, "socketpair");
    real_listen     = dlsym(RTLD_NEXT, "listen");
    real_accept     = dlsym(RTLD_NEXT, "accept");
    real_close      = dlsym(RTLD_NEXT, "close");
    real_read       = dlsym(RTLD_NEXT, "read");
    for (int i = 0; i < 1024; i++) {
        peer_of[i] = -1;
        wake_r[i]  = -1;
        wake_w[i]  = -1;
    }
}

/* ---- socket ---------------------------------------------------------- */
int socket(int domain, int type, int protocol) {
    if (domain == AF_UNIX) {
        int fd = real_socket(domain, type, protocol);
        if (fd >= 0) return fd;
        /* socket(AF_UNIX) blocked – fall back to socketpair(). */
        int sv[2];
        if (real_socketpair(domain, type, protocol, sv) == 0) {
            if (sv[0] >= 0 && sv[0] < 1024) {
                is_shimmed[sv[0]] = 1;
                peer_of[sv[0]]    = sv[1];
                int wp[2];
                if (pipe(wp) == 0) {
                    wake_r[sv[0]] = wp[0];
                    wake_w[sv[0]] = wp[1];
                }
            }
            return sv[0];
        }
        errno = EPERM;
        return -1;
    }
    return real_socket(domain, type, protocol);
}

/* ---- listen ---------------------------------------------------------- */
int listen(int sockfd, int backlog) {
    if (sockfd >= 0 && sockfd < 1024 && is_shimmed[sockfd]) {
        listener_fd = sockfd;
        return 0;
    }
    return real_listen(sockfd, backlog);
}

/* ---- accept ---------------------------------------------------------- */
int accept(int sockfd, struct sockaddr *addr, socklen_t *addrlen) {
    if (sockfd >= 0 && sockfd < 1024 && is_shimmed[sockfd]) {
        /* Block until close() writes to the wake pipe. */
        if (wake_r[sockfd] >= 0) {
            char buf;
            real_read(wake_r[sockfd], &buf, 1);
        }
        errno = ECONNABORTED;
        return -1;
    }
    return real_accept(sockfd, addr, addrlen);
}

/* ---- close ----------------------------------------------------------- */
int close(int fd) {
    if (fd >= 0 && fd < 1024 && is_shimmed[fd]) {
        int was_listener = (fd == listener_fd);
        is_shimmed[fd] = 0;

        if (wake_w[fd] >= 0) {              /* unblock accept() */
            char c = 0;
            write(wake_w[fd], &c, 1);
            real_close(wake_w[fd]);
            wake_w[fd] = -1;
        }
        if (wake_r[fd] >= 0) { real_close(wake_r[fd]); wake_r[fd]  = -1; }
        if (peer_of[fd] >= 0) { real_close(peer_of[fd]); peer_of[fd] = -1; }

        if (was_listener)
            _exit(0);                        /* conversion done – exit */
    }
    return real_close(fd);
}
`;

/**
 * Probe whether AF_UNIX sockets can be created at all. Mirrors the
 * Python `socket.socket(AF_UNIX, SOCK_STREAM)` check; returns true when
 * creation fails (i.e. the LD_PRELOAD shim is required).
 *
 * Node has no synchronous `socket()` binding, so we shell out to a tiny
 * `node -e` probe — the child either exits 0 (AF_UNIX works) or non-zero
 * (kernel blocked the syscall, shim needed). The probe binds an abstract
 * Unix path so it leaves no fs trace.
 */
export function needsShim(): boolean {
    if (process.platform !== "linux") {
        return false;
    }
    const probeScript = [
        "const net = require('node:net');",
        "const s = net.createServer();",
        "s.once('error', () => process.exit(1));",
        "s.once('listening', () => { s.close(); process.exit(0); });",
        `s.listen({ path: '\\0lo-shim-probe-' + process.pid });`,
    ].join("");
    const result = spawnSync(process.execPath, ["-e", probeScript], {
        stdio: "ignore",
        timeout: 5000,
    });
    return result.status !== 0;
}

/**
 * Compile (if necessary) the LD_PRELOAD shim and return the path to the
 * cached `.so`. Throws if `gcc` is unavailable or the build fails.
 *
 * Exported for tests; production callers should prefer `getSofficeEnv`.
 */
export function ensureShim(): string {
    if (existsSync(SHIM_SO)) {
        return SHIM_SO;
    }
    writeFileSync(SHIM_C, SHIM_SOURCE);
    const result = spawnSync("gcc", ["-shared", "-fPIC", "-o", SHIM_SO, SHIM_C, "-ldl"], { encoding: "utf8" });
    if (result.error) {
        try {
            unlinkSync(SHIM_C);
        } catch {
            /* ignore */
        }
        throw result.error;
    }
    if (result.status !== 0) {
        try {
            unlinkSync(SHIM_C);
        } catch {
            /* ignore */
        }
        throw new Error(`gcc failed (exit ${result.status}): ${result.stderr ?? ""}`.trim());
    }
    try {
        unlinkSync(SHIM_C);
    } catch {
        /* ignore */
    }
    return SHIM_SO;
}

/**
 * Build the environment for invoking soffice. Always sets
 * `SAL_USE_VCLPLUGIN=svp`; appends `LD_PRELOAD` only when the shim is
 * required and we are on Linux.
 */
export function getSofficeEnv(base?: NodeJS.ProcessEnv): NodeJS.ProcessEnv {
    const env: NodeJS.ProcessEnv = { ...(base ?? process.env) };
    env.SAL_USE_VCLPLUGIN = "svp";

    if (needsShim()) {
        const shim = ensureShim();
        env.LD_PRELOAD = shim;
    }

    return env;
}

/**
 * Run `soffice` with the given args, capturing stdout/stderr and returning
 * the exit code. Mirrors the Python `run_soffice` helper.
 */
export function runSoffice(args: string[], options?: RunSofficeOptions): Promise<RunSofficeResult> {
    const env = getSofficeEnv(options?.env);
    return new Promise((resolve, reject) => {
        const child = spawn("soffice", args, {
            env,
            cwd: options?.cwd,
            stdio: ["ignore", "pipe", "pipe"],
        });

        // Collect raw chunks so that multi-byte UTF-8 sequences split across
        // chunks are decoded once at the end (string concatenation per-chunk
        // would corrupt characters straddling a chunk boundary).
        const stdoutChunks: Buffer[] = [];
        const stderrChunks: Buffer[] = [];

        child.stdout.on("data", (chunk: Buffer) => {
            stdoutChunks.push(chunk);
        });
        child.stderr.on("data", (chunk: Buffer) => {
            stderrChunks.push(chunk);
        });

        let killTimer: NodeJS.Timeout | undefined;
        const onAbort = (): void => {
            // First SIGTERM, then escalate to SIGKILL if the process has not
            // exited within 1s — soffice occasionally ignores SIGTERM during
            // startup.
            try {
                child.kill("SIGTERM");
            } catch {
                /* already exited */
            }
            killTimer = setTimeout(() => {
                try {
                    child.kill("SIGKILL");
                } catch {
                    /* already exited */
                }
            }, 1000);
        };
        if (options?.signal) {
            if (options.signal.aborted) {
                onAbort();
            } else {
                options.signal.addEventListener("abort", onAbort, { once: true });
            }
        }

        child.on("error", (err) => {
            if (killTimer) clearTimeout(killTimer);
            options?.signal?.removeEventListener("abort", onAbort);
            reject(err);
        });
        child.on("close", (code) => {
            if (killTimer) clearTimeout(killTimer);
            options?.signal?.removeEventListener("abort", onAbort);
            resolve({
                stdout: Buffer.concat(stdoutChunks).toString("utf8"),
                stderr: Buffer.concat(stderrChunks).toString("utf8"),
                exitCode: code ?? 0,
            });
        });
    });
}

// Internal exports for tests.
export const __test = {
    SHIM_SO,
    SHIM_C,
    SHIM_SOURCE,
};

runCli(import.meta.url, async () => {
    const result = await runSoffice(process.argv.slice(2));
    process.stdout.write(result.stdout);
    process.stderr.write(result.stderr);
    return result.exitCode;
});
