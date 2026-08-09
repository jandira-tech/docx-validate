/*
 * Copyright 2026 Jandira Technologies, LLC
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *     http://www.apache.org/licenses/LICENSE-2.0
 */

/**
 * `Normalize` — bytes-in, canonical-bytes-out top-level class.
 *
 * PR C Task C.3 of the four-class architecture refactor — see
 * `docs/superpowers/specs/2026-05-29-four-class-architecture-design.md` §2.
 *
 * Unifies the scattered canonical-form passes (currently `mergeRuns`)
 * into a single transform. Distinct from `Repair`: this class assumes
 * valid input and its goal is a canonical form, not fixup.
 *
 * **`changed` semantics:** reflects whether any canonical-form PASS
 * (e.g. `mergeRuns`) reported a modification to the document content —
 * NOT whether the output bytes differ from the input. JSZip's pack step
 * is not byte-deterministic (per CLAUDE.md note 3: "outer ZIP envelope
 * can differ"), so byte-hash comparison would falsely report `changed:
 * true` even on already-canonical input.
 */

import { writeFile, readFile } from "node:fs/promises";
import path from "node:path";
import { unpack } from "./scripts/office/unpack";
import { pack } from "./scripts/office/pack";
import { mergeRuns } from "./scripts/office/helpers/merge-runs";
import { withTempDir } from "./lib/run-cli";

export type NormalizeResult = {
  bytes: Uint8Array;
  changed: boolean;
  /**
   * Per-pass change counts. Today only `mergeRuns`; future canonical-form
   * passes (e.g. `normalize-sectpr-order`) get their own entry here.
   */
  passes: { mergeRuns: number };
};

export class Normalize {
  public run(bytes: Uint8Array): Promise<NormalizeResult> {
    return withTempDir(async (dir) => {
      const inputPath = path.join(dir, "input.docx");
      await writeFile(inputPath, bytes);

      const unpackedDir = path.join(dir, "unpacked");
      await unpack(inputPath, unpackedDir);

      // Canonical-form passes. `mergeRuns` is the one we always apply.
      // `simplifyRedlines` requires an author identity and isn't
      // appropriate as a default canonical pass — consumers compose
      // it directly via the existing helper when needed.
      const mergeResult = await mergeRuns(unpackedDir);

      const outPath = path.join(dir, "out.docx");
      await pack(unpackedDir, outPath);
      const outBuffer = await readFile(outPath);
      const outBytes = new Uint8Array(outBuffer);

      return {
        bytes: outBytes,
        changed: mergeResult.count > 0,
        passes: { mergeRuns: mergeResult.count },
      };
    });
  }
}
