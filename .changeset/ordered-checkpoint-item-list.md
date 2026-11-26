---
'@docx-editor.dev/core': minor
---

An ordered room's close saves the room's whole item list as checkpoint version 6, and `closeOrderedRoom` cleans the room, checks that the new file opens to it, and writes a Word backup before a host deletes the old file. A room refuses a change made before its file as `other-file`, and a checkpoint of version 5 still opens and writes version 6 at its next close.
