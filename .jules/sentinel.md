## 2025-02-18 - Math.random() usage for Security IDs
**Vulnerability:** The code was using `Math.floor(Math.random() * X)` to generate supposedly unique, durable identifiers for comments and paragraph structures.
**Learning:** `Math.random()` is not a cryptographically secure random number generator, leading to predictable values and collision risks. These identifiers need to be unique and unpredictable.
**Prevention:** Use `randomInt` from `node:crypto` instead of `Math.random()` to generate secure, unpredictable IDs.
