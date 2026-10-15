# Jules PR Review Rules

This document specifies instructions and criteria for the automated PR review bot.

1. **Performance over style**: Look for algorithmic inefficiencies (O(N^2) DOM traversals, multiple tree climbs) rather than formatting preferences.
2. **Strict types**: Flag any introduced `any` types; prefer `unknown`.
3. **No native Node dependencies**: Check that PRs do not add C++ / native modules to `package.json` that cannot be built consistently across platforms.
4. **No missing `r:id` rules**: Verify that relationship checks remain robust and catch missing `r:id` references.
