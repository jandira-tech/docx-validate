## 2026-05-19 - Insecure Randomness in ID Generation

**Vulnerability:** ID generation functions (`generateHexId` in `comment.ts` and `repairDurableId` in `docx.ts`) used `Math.random()` to generate durable tokens and object IDs. Since `Math.random()` is not cryptographically secure and is highly predictable (seeded based on the JS engine's current state), a local or network attacker could potentially predict subsequent ID generations or brute-force collisions to tamper with document structures.
**Learning:** Using `Math.random()` for any non-trivial identification logic, especially when tracking revisions, comments, or durable document objects across users, breaks security boundary assumptions and can lead to token collision or ID prediction vulnerabilities.
**Prevention:** Always use cryptographically secure pseudo-random number generators (CSPRNG), such as `randomInt` or `randomBytes` from `node:crypto`, when generating durable, unique IDs for document metadata and internal tokens.

## 2026-05-19 - Predictable Path Vulnerability in LD_PRELOAD Shim

**Vulnerability:** The LibreOffice shim compiled to a predictable path `/tmp/lo_socket_shim.so`. Since `/tmp` is world-writable, a local attacker could pre-create this file or overwrite it with a malicious shared object. When `soffice` executed with `LD_PRELOAD=/tmp/lo_socket_shim.so`, it would execute arbitrary code as the user running the process.
**Learning:** Hardcoding paths in `/tmp` for temporary files, especially executable code like shared objects used in `LD_PRELOAD`, creates critical security risks (e.g., race conditions, local privilege escalation). This is a unique pattern since this library generates and runs native code at runtime.
**Prevention:** Always use secure temporary directories (e.g., `fs.mkdtempSync` with a specific prefix) for generating dynamic content, especially native code. Ensure file permissions restrict access to the user creating the file.

## 2026-05-19 - Predictable LibreOffice Profile Temporary Directory

**Vulnerability:** The LibreOffice macros were written into a globally shared directory: `/tmp/libreoffice_docx_profile`. Because `/tmp/` is world-writable and the path was completely static across users, a local attacker (on a shared machine or sandbox) could pre-create this directory or overwrite the `Module1.xba` macro file with malicious payloads. When LibreOffice is invoked by a higher-privileged user or automated service, it would load and execute the injected malicious StarBasic macros, leading to Local Privilege Escalation (CWE-377).
**Learning:** Hardcoded temporary paths in `/tmp/` for executing code (like scripts or Office macros) break user isolation and open pathways to injection attacks.
**Prevention:** Never use hardcoded strings under `/tmp/`. Always derive secure temporary paths dynamically using the OS's safe tempdir utilities (`os.tmpdir()`) combined with secure random tokens (`fs.mkdtempSync`), ensuring uniqueness and proper permissions to prevent concurrent users or attackers from meddling with execution profiles.

## 2026-05-19 - Side Effects from File I/O on Module Import

**Vulnerability:** While fixing the Insecure Temporary File vulnerability with `mkdtempSync`, assigning the result of `mkdtempSync` to an exported constant executed the synchronous I/O operations directly at module load time.
**Learning:** Performing side effects like file I/O (e.g., creating temporary directories) directly inside the module scope introduces architectural flaws. It means importing the file anywhere (like in test suites or other tools) inadvertently triggers directory creation, leading to orphaned files and unintended side effects, even if the target CLI function is never run.
**Prevention:** Always encapsulate file system interactions, including the generation of temporary directories or profiles, inside functions (e.g., lazy getters) rather than static module-level initialization.
