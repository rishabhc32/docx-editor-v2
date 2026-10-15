---
'@docx-editor.dev/core': minor
---

`insertHyperlink` accepts a `tooltip`, a new link gets Word's Hyperlink style even in a document that does not define it yet, and in suggesting mode a plain link's new target, new tooltip, or removal is recorded as a suggestion, while text that a copy cannot reproduce is refused instead of losing formatting.
