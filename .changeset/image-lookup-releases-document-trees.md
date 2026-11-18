---
'@docx-editor.dev/core': patch
---

A long editing session no longer keeps the document as it was first opened in memory through its picture lookups. `ImageResourceLookup.liveReferenceCount` returns zero once nothing else holds the lookup's package.
