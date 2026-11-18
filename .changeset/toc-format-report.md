---
'@docx-editor.dev/core': patch
---

Formatting a selection that includes a generated table of contents still formats the rest of the selection, and now reports the skipped table through `lastRejection` instead of leaving it out silently.
