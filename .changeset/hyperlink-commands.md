---
'@docx-editor.dev/core': minor
---

`insertHyperlink` accepts a `tooltip`, a new link gets Word's Hyperlink style even in a document that does not define it yet, and in suggesting mode a plain link's new target, new tooltip, or removal is recorded as a suggestion, while text that a copy cannot reproduce is refused instead of losing formatting. Removing a link also removes its Hyperlink style, as Word does, and `can` reports the same link refusals as `exec`.
