---
'@docx-editor.dev/core': patch
---

`exec({ type: 'insertText' })` now returns `ok: false` with the reason whenever the editor or a collaboration session refuses the text, instead of reporting success with no change.
