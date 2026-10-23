#!/usr/bin/env node
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
 * Packaged validator CLI entry (#314).
 *
 * This module is the `docx-validate` bin: `vp pack` emits it as
 * `dist/cli.mjs`, and package.json's `bin` field points there. It is NOT
 * re-exported from `src/index.ts`, so importing the library never triggers
 * argv parsing.
 *
 * The per-script `runCli(import.meta.url, …)` shims stay inside
 * `src/scripts/**` for source execution (`tsx src/scripts/office/validate.ts`),
 * but the pack build rewrites them to an inert URL — inside the bundle every
 * module shares one `import.meta.url`, so the entry-point guard cannot tell
 * `node dist/index.mjs` apart from a library import (see vite.config.ts).
 * This file is the one place the guard is allowed to fire in the bundle:
 * `import.meta.url` here IS `dist/cli.mjs`.
 */

import { runCli } from "./lib/run-cli";
import { runValidateFromArgv } from "./scripts/office/validate";

runCli(import.meta.url, () => runValidateFromArgv(process.argv.slice(2)));
