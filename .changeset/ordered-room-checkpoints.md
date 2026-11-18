---
'@docx-editor.dev/core': minor
---

Ordered collaboration rooms can save their file as a checkpoint that names media by reference instead of holding its bytes, with `writeOrderedCheckpoint`, `openOrderedCheckpoint`, and `orderedPackageMedia`. Pages read a checkpoint's media through their `media` reader, and a copy of it needs no upload.
