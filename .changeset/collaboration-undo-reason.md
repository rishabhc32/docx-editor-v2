---
'@docx-editor.dev/core': minor
---

Undo and redo in a collaboration session that cannot write, for example while offline with offline editing turned off, now return the session's refusal reason instead of "nothing to undo". A collaboration session reports that reason through the new optional `EditorCollaborationSession.writeRefusal()`.
