// A checkpoint's takeovers back in the document (design D5, policy 2).

import type { OoxmlPackage } from '../../../store/package/ooxml-package.ts';
import type { OoxmlElement } from '../../../store/package/ooxml-tree.ts';
import { createOrderedItem } from '../model/document.ts';
import { orderedStoryParts } from '../load/load.ts';
import {
  changeRecordsOf,
  recordedProperties,
  takeoverKey,
  type CheckpointSuggestion,
  type CheckpointTakeover,
  type CheckpointTakeovers,
} from '../model/checkpoint-takeovers.ts';
import { newSuggestion, relabel, replaceItems } from '../model/mutations.ts';
import { sameMaker } from '../model/person.ts';
import type {
  OrderedDocument,
  OrderedElementRecord,
  OrderedFormatRecord,
  OrderedItem,
  OrderedPropertyRecord,
  OrderedStory,
} from '../model/types.ts';
import { TripleIndex } from './ops/review.ts';

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
 * Put each takeover a checkpoint names back in the document loaded from it: the suggestion it
 * was taken from exists once, pending (the one loaded from its own marks, if the file writes
 * any), and each item whose change record holds it has the taken record again, with `replaced`,
 * in front of the holder's own. False when a takeover does not fit the file: no change record
 * at its place, or none loaded as a record of the properties it names.
 */
export function withCheckpointTakeovers(
  doc: OrderedDocument,
  base: OoxmlPackage,
  { takeovers, suggestions }: CheckpointTakeovers
): boolean {
  if (takeovers.length === 0) {
    return true;
  }

  const wanted = wantedTakeovers(takeovers);
  const holders = holdersOf(doc);
  const suggestionOf = replacedSuggestions(doc, suggestions);
  const items = new Map<OrderedItem, ItemTakeovers>();
  let found = 0;
  for (const [name, part] of orderedStoryParts(base)) {
    for (const { element, id, occurrence } of changeRecordsOf(part)) {
      const takeover = wanted.get(takeoverKey(name, element.localName, id, occurrence));
      if (takeover === undefined) {
        continue;
      }

      found += 1;
      const holdersAt = holders.get(element) ?? [];
      if (
        holdersAt.length === 0 ||
        !holdersAt.every((holder) => noteTaken(items, holder, takeover, suggestionOf))
      ) {
        return false;
      }
    }
  }

  if (found !== wanted.size) {
    return false;
  }

  for (const [item, taken] of items) {
    putBack(doc, item, taken);
  }

  return true;
}

/** The takeovers a checkpoint names, by the key of the change record that holds each. */
function wantedTakeovers(
  takeovers: readonly CheckpointTakeover[]
): Map<string, CheckpointTakeover> {
  return new Map(
    takeovers.map((takeover) => {
      const { part, record, id, occurrence } = takeover;
      return [takeoverKey(part, record, id, occurrence), takeover];
    })
  );
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
 * The suggestion of each place of a checkpoint's list, found or made the first time it is asked
 * for: null when one loaded from its marks is another maker's.
 */
function replacedSuggestions(
  doc: OrderedDocument,
  suggestions: readonly CheckpointSuggestion[]
): (k: number) => string | null {
  const index = new TripleIndex(doc);
  const ids = new Map<number, string | null>();
  return (k) => {
    let id = ids.get(k);
    if (id === undefined) {
      const listed = suggestions[k];
      id = listed === undefined ? null : replacedSuggestion(doc, index, listed);
      ids.set(k, id);
    }

    return id;
  };
}

/**
 * The suggestion a takeover's records were taken from: the one loaded from its own marks, when
 * the file writes some and it is the same maker, else a new pending one. Null when one loaded
 * from its marks is another maker's.
 */
function replacedSuggestion(
  doc: OrderedDocument,
  index: TripleIndex,
  listed: CheckpointSuggestion
): string | null {
  const { author, date } = listed.attribution;
  for (const id of listed.revisions) {
    const [loaded] = index.suggestions({ id, author, ...(date === undefined ? {} : { date }) });
    const suggestion = loaded === undefined ? undefined : doc.suggestions.get(loaded);
    if (suggestion !== undefined) {
      return sameMaker(suggestion.attribution, listed.attribution) ? suggestion.id : null;
    }
  }

  return newSuggestion(doc, listed.attribution).id;
}

/**
 * Note in `items` the records a holder took back for its item, each held by the holder's
 * suggestion: a format as the face it gave, a property record as the properties. A field's other
 * runs give only their own faces. False when the holder is not a record of the properties the
 * takeover names, or a record would be taken from the holder's own suggestion.
 */
function noteTaken(
  items: Map<OrderedItem, ItemTakeovers>,
  { story, item, record, piece }: Holder,
  takeover: CheckpointTakeover,
  suggestionOf: (k: number) => string | null
): boolean {
  const target = recordedProperties(takeover.record);
  if (record.kind === 'face' ? target !== 'rPr' : record.target !== target) {
    return false;
  }

  const replaced: { readonly id: string; readonly properties: OoxmlElement }[] = [];
  for (const { suggestion, properties } of takeover.taken) {
    const id = suggestionOf(suggestion);
    if (id === null || id === record.suggestion) {
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
