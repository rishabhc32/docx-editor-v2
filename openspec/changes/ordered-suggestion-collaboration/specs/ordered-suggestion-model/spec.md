## ADDED Requirements

### Requirement: Item addresses never change during a room session

The ordered model SHALL give every character, object, marker, paragraph mark and container token an address that stays valid for the whole room session. No edit and no decision SHALL remove, move or rename an item.

#### Scenario: Accept keeps addresses

- **WHEN** a suggestion that inserted "45" is accepted
- **THEN** each unit of "45" SHALL keep the address it had before the accept

#### Scenario: Rejected insertion

- **WHEN** a suggestion that inserted text is rejected
- **THEN** its units SHALL stay in the list, hidden, with their addresses

#### Scenario: Split and join

- **WHEN** a paragraph is split or joined
- **THEN** the addresses of every existing unit SHALL stay valid

### Requirement: Decisions are order-independent

A decision SHALL change only the status of one suggestion. Any set of decisions on the same suggestions SHALL give the same document in every order.

#### Scenario: Two decisions in both orders

- **WHEN** suggestion A is accepted and suggestion B is rejected
- **THEN** the composed document SHALL be the same whichever decision comes first

#### Scenario: Item-kind matrix

- **WHEN** the item-kind matrix runs every decision set in up to six orders
- **THEN** every order SHALL give one result

### Requirement: Word form round trip

Load SHALL turn each story part into items, and compose SHALL turn items back into the part. A file without suggestions SHALL compose to a part with the same canonical fingerprint. Pending suggestions SHALL compose to Word revision markup that Word and the engine read as the same suggestions.

#### Scenario: Save and reopen

- **WHEN** a document with pending suggestions is composed, saved and loaded again
- **THEN** the loaded view SHALL equal the composed view, and each decision SHALL give the same result as before the save, except for the documented Word limits

#### Scenario: One paragraph

- **WHEN** a change touches one paragraph
- **THEN** compose SHALL be able to build only that paragraph, with its paragraph identity kept

### Requirement: Placement follows the engine

The model SHALL place typed text, breaks and markers by the engine's placement rules, in one function.

#### Scenario: Matrix against the engine

- **WHEN** the item-kind matrix runs in CI
- **THEN** every view SHALL equal the engine's view, or the difference SHALL be on the documented list of engine defects

### Requirement: A second deletion keeps the first

A deletion of text that another suggestion already struck SHALL do nothing on the struck part.

#### Scenario: Second deletion

- **WHEN** Ben deletes text that Ana already deleted
- **THEN** the text SHALL keep only Ana's deletion

### Requirement: A second format replaces the first

A format on text that another suggestion already formatted SHALL replace the earlier format record and SHALL report the replaced suggestion to the host.

#### Scenario: Second format

- **WHEN** Ben formats text that Ana already formatted
- **THEN** Ben's format SHALL show and Ana's suggestion SHALL be reported as replaced

### Requirement: One card is one decision

A deletion next to an insertion by the same author SHALL form one card. Changes inside another person's suggestion SHALL join that card. A decision SHALL apply to a whole card.

#### Scenario: Replacement

- **WHEN** Ana deletes "30" and types "45" in its place
- **THEN** the review SHALL show one card, and one accept SHALL apply both parts

### Requirement: Enter inside a control prompt

Enter inside a control's prompt SHALL remove the prompt and then break. When the control allows one line only, Enter SHALL be refused by name.

#### Scenario: Single-line control

- **WHEN** a person presses Enter in the prompt of a plain text control without multiple lines
- **THEN** the model SHALL refuse the edit by name and change nothing

### Requirement: Unsupported edits are refused by name

The model SHALL refuse an edit it cannot express, with a stable reason code. It SHALL NOT drop any part of an edit without a refusal.

#### Scenario: Not covered kind

- **WHEN** an edit types inside a text box
- **THEN** the model SHALL refuse it with a named reason and change nothing

### Requirement: Undo record

Every edit SHALL be able to run inside an undo record that takes it back exactly, without copying the document.

#### Scenario: Take back and apply again

- **WHEN** a client takes back two pending changes, applies one server change, and applies its two changes again
- **THEN** the result SHALL equal applying the server change first and the two changes after it
