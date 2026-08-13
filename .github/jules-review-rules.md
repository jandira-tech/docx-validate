# Jules PR Review Rules

- Enforce that all changes are free of hardcoded secrets or API keys.
- Security fixes should fail securely and include comments explaining the security concern.
- PR titles for security fixes must follow the format '🛡️ Sentinel: [CRITICAL/HIGH] Fix [vulnerability type]' or '🛡️ Sentinel: [security improvement]'.
- Ensure the codebase uses cryptographically secure pseudo-random number generators (like `crypto.randomInt()`) instead of `Math.random()` for IDs or tokens.
- All new files or modified sections should have corresponding unit tests where applicable.
