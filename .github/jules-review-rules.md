# Jules PR Review Rules

This document outlines the rules for the automated code review system.

## Groundedness
All claims about the codebase must be grounded in the code. Reviewers must verify any claims about missing variables or unused imports against the codebase using tools like `grep` before raising an issue, to avoid hallucinations.

## Security
1. **No hardcoded secrets:** Never commit secrets, API keys, or passwords.
2. **Secure PRNGs:** Do not use `Math.random()` for generating IDs or durable tokens, as it is not cryptographically secure. Use `randomInt` from `node:crypto` instead.
3. **Fail securely:** Security fixes must fail securely and not leak sensitive information in error messages or stack traces.
4. **Least privilege:** Use established security libraries and follow the principle of least privilege.
