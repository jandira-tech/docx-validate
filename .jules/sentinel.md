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
## 2025-02-13 - Replace insecure Math.random() with crypto.randomInt()
**Vulnerability:** Predictable PRNG (`Math.random()`) used for generating security-sensitive or collision-sensitive durable IDs (`paraId`, `durableId`).
**Learning:** `Math.random()` does not provide cryptographic security and its values can be predicted or easily collide in certain JavaScript environments. This could lead to predictable IDs which can be an issue if those IDs are expected to be somewhat unguessable, or if collisions matter.
**Prevention:** Always use `node:crypto` functions like `randomInt` or `randomUUID` for generating unique IDs or tokens when executing in a Node.js context.
