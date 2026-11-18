---
'@docx-editor.dev/core': patch
---

Accepting text suggested over a content control's prompt now leaves the control holding that text without the prompt's state or style, also in files that an earlier version saved. A content control with no declared type shows its prompt when it is empty, like a rich text control.
