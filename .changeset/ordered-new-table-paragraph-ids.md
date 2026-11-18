---
'@docx-editor.dev/core': patch
---

In an ordered collaboration room, each paragraph of a new table, row, or other block an edit makes gets its own `w14:paraId`, so two new tables never share paragraph ids and a saved file reopens to the same document.
