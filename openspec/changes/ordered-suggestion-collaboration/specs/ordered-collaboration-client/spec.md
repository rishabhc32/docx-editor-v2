## ADDED Requirements

### Requirement: The client implements the editor session contract

The ordered client SHALL implement `EditorCollaborationSession` and SHALL attach through `CollaborationDocumentPort`, with no change to the editor.

#### Scenario: Attach

- **WHEN** a host creates an ordered client and passes it to the collaboration module
- **THEN** the editor SHALL attach its port, and `attached` SHALL become true

### Requirement: Local edits are admitted or refused before they commit

`gateOperations` SHALL translate each supported `TreeDocOp` into changes and apply them to the local list inside an undo record. It SHALL return a refusal code, and change nothing, when an op is not supported, cannot be translated, or the list refuses it. It SHALL refuse an empty batch from an opaque write with `ordered-opaque-write`.

#### Scenario: Typing

- **WHEN** a person types in suggesting mode
- **THEN** the client SHALL admit the ops, queue one change, and send it

#### Scenario: Unsupported op

- **WHEN** an op such as `insertDrawing` with a new image file arrives
- **THEN** the client SHALL return `ordered-unsupported-op`, and the editor SHALL not commit

#### Scenario: A key lands while the editor takes the room's change

- **WHEN** a key waits to be typed as the room's change arrives, and a host that reads the editor on each change lands it while the editor takes the list's version
- **THEN** the client SHALL send the key's change, and the key SHALL stay on both editors

### Requirement: The screen follows the list

After a local commit, the client SHALL publish the list's version of each touched paragraph that differs from the editor's. For a remote change, the client SHALL publish only the touched paragraphs through `applyRemotePackage`.

#### Scenario: Placement difference

- **WHEN** the engine places typed text on the other side of a bookmark end than the list
- **THEN** the editor SHALL show the list's version after the commit

### Requirement: Pending changes are applied again on top of server changes

When a server change arrives, the client SHALL take back its pending changes, apply the server change, and apply its pending changes again. A pending change that no longer applies SHALL be refused and reported to the host.

#### Scenario: Concurrent typing at one place

- **WHEN** two people type at the same place at the same time
- **THEN** both editors SHALL end with the same document as the server

### Requirement: Undo touches only the person's own suggestions

Undo SHALL emit a new change that takes back the person's last own change. It SHALL NOT change other people's text. Undo of a change without an inverse SHALL be refused by name.

#### Scenario: Undo after another person's change

- **WHEN** Ana types, Ben types after it, and Ana presses Undo
- **THEN** Ana's text SHALL go and Ben's text SHALL stay on both editors

#### Scenario: The organiser undoes a direct edit

- **WHEN** the organiser deletes text directly that holds Ana's pending insertion, and presses Undo
- **THEN** the text SHALL show again, Ana's insertion SHALL show again as her pending insertion, and both editors SHALL end with the same document as the server

#### Scenario: A party cannot take back a direct edit

- **WHEN** a party sends a `revert` change
- **THEN** the server SHALL refuse it by name (`policy-direct`) and change nothing

### Requirement: Reconnect and restart

The client SHALL keep pending changes and the undo records of unsaved changes until the server sends `saved`. After a reconnect it SHALL go back to the server's order number and send its own changes again. After a `reload`, it SHALL keep a copy of unsent changes for the host.

#### Scenario: Server restart

- **WHEN** the server restarts at order number 57 and the client is at 61
- **THEN** the client SHALL take back 58 to 61, send its own changes among them again, and end equal to the server

### Requirement: Presence

The client SHALL send its selection as paragraph id and offset, and SHALL show other people's selections. Presence SHALL need no order.

#### Scenario: Remote caret

- **WHEN** Ben moves his caret
- **THEN** Ana's editor SHALL show Ben's caret at the same characters
