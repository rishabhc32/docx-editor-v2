---
'@docx-editor.dev/core': minor
---

An edit that a collaboration room refuses now fails an editing API sync with `debugInfo.innerError.code` set to the new diagnostic `edit.collaboration-refused`, instead of a `GeneralException` with no reason.
