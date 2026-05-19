## 2026-05-19 - Predictable Path Vulnerability in LD_PRELOAD Shim

**Vulnerability:** The LibreOffice shim compiled to a predictable path `/tmp/lo_socket_shim.so`. Since `/tmp` is world-writable, a local attacker could pre-create this file or overwrite it with a malicious shared object. When `soffice` executed with `LD_PRELOAD=/tmp/lo_socket_shim.so`, it would execute arbitrary code as the user running the process.
**Learning:** Hardcoding paths in `/tmp` for temporary files, especially executable code like shared objects used in `LD_PRELOAD`, creates critical security risks (e.g., race conditions, local privilege escalation). This is a unique pattern since this library generates and runs native code at runtime.
**Prevention:** Always use secure temporary directories (e.g., `fs.mkdtempSync` with a specific prefix) for generating dynamic content, especially native code. Ensure file permissions restrict access to the user creating the file.
