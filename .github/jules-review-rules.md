# Jules Review Rules

This project validates and processes OOXML (.docx/.pptx) files using `xmldom` and native DOM APIs.

## Performance
- Do not use `xpath` descendant queries (`.//`) on large XML subtrees as they cause extreme performance bottlenecks in `xmldom`. Use `getElementsByTagNameNS` and deduplicate with Sets when dealing with potentially nested identical tags (like `<w:del>`).
- All code must run smoothly for extremely large document profiles.

## Formatting & Types
- Ensure arrays are typed as `T[]`, NOT `Array<T>`.
- Object definitions must use `interface`, NOT `type`.

## Security
- Do not use `Math.random()` for token or ID generation; use `crypto.randomInt`.
