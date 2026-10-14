---
'@docx-editor.dev/editor-api': patch
---

`Revision.accept()` and `Revision.reject()` name the reason in `debugInfo.innerError` when they refuse a change that cannot be decided alone, such as `review.incomplete-group` for a deletion inside another author's insertion.
