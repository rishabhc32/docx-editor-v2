---
'@docx-editor.dev/core': minor
---

A content lock now keeps only the control's own content, so a control nested in a content-locked control or a group, such as a field in a Word-saved group, can be filled, as in Word. Accept All and Reject All now leave the changes in locked content and decide the others, as Word does, and report them as skipped with the reason `locked`.
