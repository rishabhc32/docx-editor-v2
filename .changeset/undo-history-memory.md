---
'@docx-editor.dev/core': patch
---

An editor in a collaboration session no longer keeps old versions of the document in memory through its local undo history and remote caret positions, so a long session no longer grows without limit.
