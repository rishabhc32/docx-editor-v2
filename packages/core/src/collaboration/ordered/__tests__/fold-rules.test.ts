// How a checkpoint keeps each fact of the room's model (design D14): the field table of
// `model/snapshot-fields.ts` and `model/snapshot-side.ts` gives every field of every type stored,
// with its codec, or rebuilt by the open, with how. Each table covers every key of its type, so a
// new field fails the type check there until it has an entry. `fold-check.ts` checks after each
// step of the shape sessions that an open gives back every stored field.
//
// What a fold of version 5 dropped is stored now: a reopened room holds what the live room held
// (`retracted`, `gone`, `cancelledFmts`, `takenBackLendings`, `taken`, each record's `change`,
// `acceptedOnly`, a direct edit's object records, `direct`, `composeIds`, `reattributed`, the
// `insertion` counter, each resource's key and suggestion, `added`). The cleanup of a close decides
// what goes (design D14).

import { describe, expect, test } from 'bun:test';
import type { SnapshotFields } from '../model/snapshot-codec.ts';
import {
  ADDRESS,
  ATTRIBUTION,
  COMMENT,
  COMMENT_RANGE,
  COUNTERS,
  DOCUMENT,
  ELEMENT_RECORD,
  FACE_FORMAT,
  FILE_MEDIA,
  FILE_STORE,
  ITEM,
  LIST_BULLET,
  LIST_LEVEL_VALUE,
  NOTE_PROPERTIES,
  OBJECT_RECORD,
  PARAGRAPH_OP_RECORD,
  PART_REVISION,
  PIECE,
  PROPERTIES_FORMAT,
  RESOURCES,
  RESOURCES_BY_KIND,
  REVISION_IDS,
  SECTION_RECORD,
  STORY,
  SUGGESTION,
  UNIT_RANGE,
  WITHIN,
} from '../model/snapshot-fields.ts';
import {
  NAMED_PROMPT_STATE,
  PROMPTED_BY_EDIT,
  SIDE,
  SWAP,
  TEXT_BOX_HOST,
  TYPED_RUN,
  UNTRACKED,
} from '../model/snapshot-side.ts';

/** Every table, by the type it covers. */
const TABLES: Readonly<Record<string, SnapshotFields<object>>> = {
  OrderedItem: ITEM,
  OrderedWithin: WITHIN,
  OrderedFieldPiece: PIECE,
  'OrderedFormatRecord properties': PROPERTIES_FORMAT,
  'OrderedFormatRecord face': FACE_FORMAT,
  OrderedElementRecord: ELEMENT_RECORD,
  OrderedParagraphOpRecord: PARAGRAPH_OP_RECORD,
  OrderedSectionRecord: SECTION_RECORD,
  OrderedObjectRecord: OBJECT_RECORD,
  OrderedNoteProperties: NOTE_PROPERTIES,
  OrderedSuggestion: SUGGESTION,
  OrderedAttribution: ATTRIBUTION,
  OrderedStory: STORY,
  OrderedItemAddress: ADDRESS,
  OrderedUnitRange: UNIT_RANGE,
  OrderedDocument: DOCUMENT,
  OrderedCounters: COUNTERS,
  OrderedRevisionIds: REVISION_IDS,
  OrderedPartRevision: PART_REVISION,
  OrderedComment: COMMENT,
  OrderedCommentRange: COMMENT_RANGE,
  OrderedResources: RESOURCES,
  OrderedFileMedia: FILE_MEDIA,
  OrderedFileStore: FILE_STORE,
  OrderedListLevel: LIST_LEVEL_VALUE,
  OrderedListBullet: LIST_BULLET,
  ...Object.fromEntries(
    Object.entries(RESOURCES_BY_KIND).map(([kind, table]) => [`OrderedResource ${kind}`, table])
  ),
  OrderedTextBoxHost: TEXT_BOX_HOST,
  TypedRun: TYPED_RUN,
  Untracked: UNTRACKED,
  PromptedByEdit: PROMPTED_BY_EDIT,
  Swap: SWAP,
  NamedPromptState: NAMED_PROMPT_STATE,
};

/** The fields an open makes again: each one a reviewed choice. */
const REBUILT = [
  'OrderedItem.leaf',
  'OrderedStory.spaces',
  'OrderedDocument.file',
  'OrderedDocument.labelled',
  'OrderedDocument.ownEdited',
  'OrderedDocument.undo',
  'OrderedResources.fileMedia',
];

describe('every fact of the model has a checkpoint rule', () => {
  test('the open makes again only the reviewed fields, each with how', () => {
    const rebuilt = Object.entries(TABLES).flatMap(([type, table]) =>
      Object.entries(table).flatMap(([key, field]) =>
        'rebuilt' in field ? [{ field: `${type}.${key}`, how: field.rebuilt }] : []
      )
    );
    expect(rebuilt.map(({ field }) => field)).toEqual(REBUILT);
    expect(rebuilt.filter(({ how }) => how.trim().length < 10)).toEqual([]);
  });

  test('each side state module keeps beside the document is saved', () => {
    expect(Object.keys(SIDE)).toEqual([
      'textBoxes',
      'typers',
      'droppedMarks',
      'filledPrompts',
      'firstFills',
      'emptiedControls',
      'namedPrompts',
      'splitParts',
    ]);
  });
});
