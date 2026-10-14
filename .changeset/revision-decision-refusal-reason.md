---
'@docx-editor.dev/editor-api': patch
---

`Revision.accept()`, `Revision.reject()`, and decisions on a review item key name the reason in `debugInfo.innerError` when they refuse a change that cannot be decided alone, such as `review.incomplete-group` for a formatting change inside another author's insertion.
