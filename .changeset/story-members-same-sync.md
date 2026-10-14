---
'@docx-editor.dev/editor-api': patch
---

Members of a header, footer, or note body, such as `revisions`, `lists`, `search()`, and `insertParagraph()`, now work in the same `context.sync()` that requests the body instead of failing with `InvalidObjectPath`.
