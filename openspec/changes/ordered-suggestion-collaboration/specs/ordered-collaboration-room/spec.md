## ADDED Requirements

### Requirement: One order for all changes

The room SHALL give each accepted change the next order number and SHALL send changes to every client in that order.

#### Scenario: Two clients

- **WHEN** two clients send changes at the same time
- **THEN** every client SHALL receive the same changes in the same order

### Requirement: Untrusted input is validated

The room SHALL validate every client message against the change format, with size, count and rate limits. The room SHALL set the author from the login and the time from its clock, and SHALL ignore any author, time or order number that a client sends.

#### Scenario: Forged author

- **WHEN** a client sends a change that names another author
- **THEN** the room SHALL record the change under the logged-in person, or refuse it

#### Scenario: Oversized batch

- **WHEN** a batch exceeds the size cap
- **THEN** the room SHALL refuse it with a named code and apply nothing from it

#### Scenario: Locked content control

- **WHEN** a client sends a change that edits the content of a `contentLocked` or `sdtContentLocked` content control, or removes an `sdtLocked` or `sdtContentLocked` content control
- **THEN** the room SHALL refuse it by name (`locked`) and apply nothing, for the organiser and a party alike

### Requirement: Per-change policy

The room SHALL check each change with a policy that reads only the change, the items it touches, and what applying the change did. The suggestions-only policy SHALL give a party's change the verdict of the host's earlier check, except for the documented faults of that check's Word form.

#### Scenario: Same verdicts as today's check

- **WHEN** the same party changes run through the host's earlier check and the policy
- **THEN** both SHALL give the same verdict for each change, or the difference SHALL be documented

### Requirement: What a party may change

The suggestions-only policy SHALL allow a party to add suggestions, change or take back their own suggestions, and mark their own discussions through the marker request. It SHALL refuse decisions, header and footer changes and direct edits that change something, changes nested in another person's suggestion as Word nests them, a format or property change that takes over another person's pending record, and a party's discussion marker that comes in a batch of changes.

#### Scenario: Party accepts

- **WHEN** a party sends a decision
- **THEN** the room SHALL refuse it and apply nothing

#### Scenario: Party marks a discussion with their own comment

- **WHEN** a party sends a `marker` change in a batch of changes, also for a discussion they opened
- **THEN** the room SHALL refuse it by name (`policy-marker`, reason `marker`) and apply nothing
- **AND** the same marker sent as a marker request SHALL get the request's checks: a discussion the party opened, not marked yet

#### Scenario: Party formats over another person's pending format

- **WHEN** a party's format or property change would take over another person's pending record
- **THEN** the room SHALL refuse it by name (`policy-not-own`, reason `others_suggestions`) and take it back
- **AND** the organiser's same change SHALL take the record over (policy 2)

#### Scenario: Party's direct edit changes nothing

- **WHEN** a party's direct edit, or a party's change in a header or footer, makes, extends, or takes back no suggestion
- **THEN** the room SHALL accept it, as today's check finds the document the same

### Requirement: Duplicates are applied once

The room SHALL apply a change id once. A repeated id SHALL get the order number of the first copy.

#### Scenario: Resend after a crash

- **WHEN** a client sends a change again that the room already applied
- **THEN** the room SHALL answer with the existing order number and apply nothing

### Requirement: Durable log before broadcast

The server SHALL write each batch to its local log and sync it before it sends the batch to anyone. It SHALL send `saved` for an order number only after S3 holds that number. The S3 copy SHALL never go back to an earlier order number.

#### Scenario: Overlapping uploads

- **WHEN** an older upload would finish after a newer one
- **THEN** the conditional write SHALL refuse the older upload

### Requirement: Restart rebuilds the room

On restart the server SHALL rebuild the room from the Word file and the change log. If an order number is missing, it SHALL refuse to open the room and report the error.

#### Scenario: Crash after sync, before send

- **WHEN** the server crashes after it synced a batch and before it sent the batch
- **THEN** after restart the batch SHALL be in the room, and a resent change SHALL be recognized as a duplicate

### Requirement: Close folds the log into a Word file

After the last person leaves and a wait, the server SHALL compose the Word file, write it as the next file, point the room record to it with a conditional write, and delete the old log, in this order.

#### Scenario: Crash during close

- **WHEN** the server crashes between any two close steps
- **THEN** the next open SHALL give the same document as before the close

### Requirement: One writer per room

A room SHALL have one writer. A process that takes a room SHALL increase the writer number, so the old process's next upload fails and it stops serving the room.

#### Scenario: Two processes

- **WHEN** a second process opens a room that another process serves
- **THEN** only one process SHALL keep writing, and no confirmed change SHALL be lost
