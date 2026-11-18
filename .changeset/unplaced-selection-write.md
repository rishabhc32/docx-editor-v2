---
'@docx-editor.dev/core': patch
---

Formatting, indenting, and deleting a selection that ends on a page still being laid out now apply to the whole selection, and a selection that ends in hidden text is refused with a reason instead of changing nothing.
