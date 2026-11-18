---
'@docx-editor.dev/core': patch
---

A long editing session no longer keeps a header or footer layout and a numbering definition copy for each older revision in the undo history, so memory stays bounded when list and header edits repeat, for example in a collaboration room.
