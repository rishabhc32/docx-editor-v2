---
'@docx-editor.dev/core': minor
---

In ordered collaboration rooms, a pending change that fails with an unexpected error when the client applies it again is now taken back and refused by name, so the participant can still rejoin the room and keeps their other unsent changes. The new `onApplyError` option of `createOrderedCollaboration` reports the error.
