---
'@docx-editor.dev/core': minor
---

In an ordered collaboration room, a key press lays out the pages once instead of twice, and reads the review queue again only for the paragraph it changed. A collaboration session can implement the new optional `EditorCollaborationSession.settleLocalCommit()` to give the editor its version of a local edit before the layout.
