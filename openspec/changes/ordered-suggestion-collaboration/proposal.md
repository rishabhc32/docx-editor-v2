## Why

Contract review rooms need every party change tracked exactly and decided by the server. Today a room runs three heavy layers on every edit: the engine store with tracked changes inside the Word tree, a Yjs copy of the whole package in every browser, and a server check that copies the document and compares it in full. On a 521-page contract this costs 65 to 75 ms for a local edit with sync, 3.5 s for the server check of one party edit (8.1 s for the first), about 880 MB of memory per person, and 203 ms for one accept or reject. Decisions also depend on their order in some cases, and Reject All loses content-control prompts.

Spike 5 on branch `spike/suggestion-log` (commit `bfd70f4a`, `spikes/suggestion-seq/REPORT.md`) tested a different model: one ordered list of items per story, held in a summed B+ tree, where a suggestion is a label on items and a decision only flips a status. Over 65 Word item kinds and 4,080 scenarios, every set of decisions gave one result in every order. Against today's stack on 521 pages, a local edit with sync took 0.25 to 0.33 ms, the server step 0.06 ms, a decision 0.004 ms, and one person's copy about 120 MB.

## What Changes

- Add an ordered suggestion model to the engine's `collaboration` lane: the item list, its tree, Word load and Word compose, the edits, decisions, and an undo record. It is DOM-free and runs in browsers and on Node.
- Add a versioned JSON change format with strict validation. A change names places by item addresses that never change during a room session, and new items take ids made from the change id, so every copy applies a change the same way.
- Add an ordered collaboration client that implements `EditorCollaborationSession`. The editor applies local edits as today; the client turns each admitted `TreeDocOp` batch into changes, refuses what it cannot express by name, sends changes to the server, and publishes the list's Word form for remote changes through `applyRemotePackage`.
- Add an ordered room core: one order for all changes, a pluggable per-change policy (a contract host's suggestions-only rule as the first policy), duplicate detection, and decisions as messages.
- In the host server (separate repository, TypeScript on Node 24), replace Hocuspocus and Yjs with ordered rooms: a zstd write-ahead change log on local disk, uploads to S3 with conditional writes, a room close that folds the log into a new Word file, restart from the Word file and the log, and fencing for two processes.
- Fix four engine defects that the spike found, as separate changes on `main`.

## Capabilities

### New Capabilities

- `ordered-suggestion-model`: the item list, Word load and compose, edits, decisions, rules.
- `ordered-collaboration-client`: the browser and headless session over the editor.
- `ordered-collaboration-room`: ordering, validation, policy, persistence contract, restart.

### Modified Capabilities

- None. Yjs collaboration (`full-document-yjs-collaboration`) stays available in the engine for other hosts. The first host uses ordered rooms only.

## Impact

- `packages/core/src/collaboration/ordered/` (new), a new export subpath `@docx-editor.dev/core/collaboration/ordered`, API Extractor snapshot, lane graph entry.
- The host server (separate repository): new ordered room transport and storage, TypeScript, Node 24, AWS SDK v3 conditional writes, `node:zlib` zstd.
- The host's web app: create the ordered client in place of the Hocuspocus provider for new rooms.
- No migration: the host deploys ordered rooms as a new system, with no earlier rooms or files to carry over. The host server's Hocuspocus and Yjs path is removed after the multi-party test passes.
- Every edit the editor can make works in ordered rooms in the first release, including pictures, links, fields, controls, columns, headers and footers, sections, notes, equations, text boxes, paste with formatting, new list definitions, and the editor's own accept, reject and comment actions. The only refusals are the host's party rules that its earlier rooms also apply.
