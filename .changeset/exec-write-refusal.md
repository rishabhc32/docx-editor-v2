---
'@docx-editor.dev/core': patch
---

`exec` write commands such as `deleteText`, `cut`, `paste`, `insertBreak`, `toggleMark`, and paragraph formatting now return `ok: false` with the reason when the document or a collaboration room refuses the write, instead of reporting success with no change.
