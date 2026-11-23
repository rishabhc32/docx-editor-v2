// A checkpoint's takeovers back in the document (design D5, policy 2).

import type { OoxmlElement } from '../../../store/package/ooxml-tree.ts';
import { createOrderedItem } from '../model/document.ts';
import {
  holdsTakeover,
  recordedProperties,
  type CheckpointTakeover,
} from '../model/checkpoint-takeovers.ts';
import { relabel, replaceItems } from '../model/mutations.ts';
import type {
  OrderedDocument,
  OrderedElementRecord,
  OrderedFormatRecord,
  OrderedItem,
  OrderedPropertyRecord,
  OrderedStory,
} from '../model/types.ts';

/** A file's format record. */
type FaceRecord = Extract<OrderedFormatRecord, { readonly kind: 'face' }>;

/** A loaded record that holds a file's change record, and the item it is on. */
interface Holder {
  readonly story: OrderedStory;
  readonly item: OrderedItem;
  readonly record: FaceRecord | OrderedElementRecord;
  /** A field's run whose own change record it is, by its piece's place. */
  readonly piece?: number;
}

/** What one item gets back: the taken records in front of each of its own, by that record. */
interface ItemTakeovers {
  readonly story: OrderedStory;
  readonly fmts: Map<OrderedFormatRecord, OrderedFormatRecord[]>;
  readonly props: Map<OrderedPropertyRecord, OrderedPropertyRecord[]>;
  /** A field's runs: each one's own face under each taken record, by its piece's place. */
  readonly faces: Map<number, Map<string, OoxmlElement | null>>;
}

/**
 * Put each takeover a checkpoint names back in the document loaded from it: each item whose
 * change record holds it has the taken record again, with `replaced` (a suggestion of the
 * manifest's list, by `listedIds`), in front of the holder's own. `labels` gives the elements
 * that carry each label. False when a takeover does not fit the file: its label on no change
 * record a loaded record of the properties it names holds, or on two (a text box's other copy
 * holds none), or on another element.
 */
export function withCheckpointTakeovers(
  doc: OrderedDocument,
  labels: ReadonlyMap<string, readonly OoxmlElement[]>,
  takeovers: readonly CheckpointTakeover[],
  listedIds: readonly string[]
): boolean {
  if (takeovers.length === 0) {
    return true;
  }

  const holders = holdersOf(doc);
  const items = new Map<OrderedItem, ItemTakeovers>();
  for (const takeover of takeovers) {
    const held = (labels.get(String(takeover.label)) ?? []).filter((e) => holders.has(e));
    const element = held[0];
    const holdersAt = element === undefined ? [] : (holders.get(element) ?? []);
    if (
      held.length !== 1 ||
      element === undefined ||
      !holdsTakeover(element, takeover) ||
      !holdersAt.every((holder) => noteTaken(items, holder, element.localName, takeover, listedIds))
    ) {
      return false;
    }
  }

  for (const [item, taken] of items) {
    putBack(doc, item, taken);
  }

  return true;
}

/** Every loaded record that holds a file's change record, by that record. */
function holdersOf(doc: OrderedDocument): Map<OoxmlElement, Holder[]> {
  const out = new Map<OoxmlElement, Holder[]>();
  const add = (element: OoxmlElement, holder: Holder) => {
    const list = out.get(element);
    if (list === undefined) {
      out.set(element, [holder]);
    } else {
      list.push(holder);
    }
  };
  for (const story of doc.stories.values()) {
    for (let item = story.tree.first(); item !== null; item = story.tree.next(item)) {
      for (const record of item.props) {
        if (record.kind === 'element' && record.fileRecord) {
          add(record.fileRecord, { story, item, record });
        }
      }

      // A field's runs each keep their own change record of the field's one record
      for (const [piece, { recordFiles }] of (item.pieces ?? []).entries()) {
        for (const [id, element] of recordFiles ?? []) {
          const record = item.fmts.find((f) => f.id === id);
          if (record?.kind === 'face') {
            add(element, { story, item, record, piece });
          }
        }
      }

      for (const record of item.pieces === undefined ? item.fmts : []) {
        if (record.kind === 'face' && record.fileRecord) {
          add(record.fileRecord, { story, item, record });
        }
      }
    }
  }

  return out;
}

/**
 * Note in `items` the records a holder took back for its item, each held by the holder's
 * suggestion: a format as the face it gave, a property record as the properties. A field's other
 * runs give only their own faces. False when the holder is not a record of the properties its
 * change record (`changeRecord`, a local name) names, or a record would be taken from the holder's
 * own suggestion.
 */
function noteTaken(
  items: Map<OrderedItem, ItemTakeovers>,
  { story, item, record, piece }: Holder,
  changeRecord: string,
  takeover: CheckpointTakeover,
  listedIds: readonly string[]
): boolean {
  const target = recordedProperties(changeRecord);
  if (record.kind === 'face' ? target !== 'rPr' : record.target !== target) {
    return false;
  }

  const replaced: { readonly id: string; readonly properties: OoxmlElement }[] = [];
  for (const { suggestion, properties } of takeover.taken) {
    const id = listedIds[suggestion];
    if (id === undefined || id === record.suggestion) {
      return false;
    }

    replaced.push({ id, properties });
  }

  let entry = items.get(item);
  if (entry === undefined) {
    entry = { story, fmts: new Map(), props: new Map(), faces: new Map() };
    items.set(item, entry);
  }

  if (record.kind === 'element') {
    const taken = replaced.map(
      ({ id, properties }): OrderedPropertyRecord => ({
        kind: 'element',
        suggestion: record.suggestion,
        seq: record.seq,
        replaced: id,
        target: record.target,
        element: properties,
        recordUnchanged: true,
      })
    );
    entry.props.set(record, taken);
    return true;
  }

  const taken = replaced.map(
    ({ id, properties }, k): FaceRecord => ({
      kind: 'face',
      id: `${record.id}<${k}`,
      suggestion: record.suggestion,
      seq: record.seq,
      replaced: id,
      face: properties,
    })
  );
  if (!entry.fmts.has(record)) {
    entry.fmts.set(record, taken);
  }

  if (piece !== undefined) {
    entry.faces.set(piece, new Map(taken.map(({ id, face }) => [id, face])));
  }

  return true;
}

/**
 * An item with the taken records in front of its own, which keep their seq: the taken records
 * take it too, so they draw first, as they were made. A field's runs get their own faces.
 */
function putBack(doc: OrderedDocument, item: OrderedItem, taken: ItemTakeovers): void {
  const { story, faces } = taken;
  const fmts = item.fmts.flatMap((record) => [...(taken.fmts.get(record) ?? []), record]);
  const props = item.props.flatMap((record) => [...(taken.props.get(record) ?? []), record]);
  if (item.pieces === undefined || faces.size === 0) {
    relabel(doc, story, item, { fmts, props });
    return;
  }

  const pieces = item.pieces.map((piece, k) => {
    const more = faces.get(k);
    if (more === undefined) {
      return piece;
    }

    return { ...piece, recordFaces: new Map([...(piece.recordFaces ?? []), ...more]) };
  });
  replaceItems(doc, story, [item], [createOrderedItem({ ...item, fmts, props, pieces })]);
}
