---
'@docx-editor.dev/core': patch
---

The surface's `setSelection` now refuses a selection in a header, footer, or other part that is not the open story, and reports the reason in `lastRejection`, so typing can no longer change a part that the pages do not show.
