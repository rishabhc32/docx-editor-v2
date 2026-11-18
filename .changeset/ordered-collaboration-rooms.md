---
'@docx-editor.dev/core': minor
---

Add experimental ordered collaboration rooms in `@docx-editor.dev/core/collaboration/ordered`, where a room server puts every change in one order, checks each change against a policy such as suggestions only, and refuses an edit by name. Connect an editor with `createOrderedCollaboration`, and run the room with `OrderedRoom`.
