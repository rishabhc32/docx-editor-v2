// A checkpoint's records made over a change record's pending ones back in the document (design
// D14).

import type { OoxmlElement } from '../../../store/package/ooxml-tree.ts';
import type { CheckpointMadeOver, CheckpointOverRecord } from '../model/checkpoint-over.ts';
import { createOrderedItem } from '../model/document.ts';
import { nextSeq, relabel, replaceItems } from '../model/mutations.ts';
import type {
  OrderedDocument,
  OrderedFormatRecord,
  OrderedItem,
  OrderedPropertyRecord,
  OrderedStory,
} from '../model/types.ts';
import { isWml } from '../ooxml-nodes.ts';
import { holdersOf, type Holder } from './takeovers.ts';

/** A made-over record with its suggestions' ids and its order, the item's last. */
type Made = Omit<CheckpointOverRecord, 'suggestion' | 'over'> & {
  readonly suggestion: string;
  readonly seq: number;
  readonly over: readonly string[];
};

/** A file's format record. */
type FaceRecord = Extract<OrderedFormatRecord, { readonly kind: 'face' }>;

/** What one item gets back: each holder with its properties, and the records made over it. */
interface ItemOver {
  readonly story: OrderedStory;
  readonly fmts: Map<OrderedFormatRecord, readonly OrderedFormatRecord[]>;
  readonly props: Map<OrderedPropertyRecord, readonly OrderedPropertyRecord[]>;
  /** A field's runs: each one's own face under the holder, by its piece's place. */
  readonly faces: Map<number, Map<string, OoxmlElement>>;
}

/**
 * Put back in the document loaded from a checkpoint the records made over the pending records of
 * each change record the manifest names: each loaded record that holds it gives the properties
 * the manifest gives, without those records, and the records follow it, each in an order after
 * every record of the document, so compose draws them on top as the room did. `labels` gives the
 * elements that carry each label. False when an entry does not fit the file: its label on no
 * change record of its properties that a loaded record holds, or on two; a suggestion not listed,
 * or of another status than its record's (a record on its author's own new words is pending, a
 * direct one accepted).
 */
export function withCheckpointMadeOver(
  doc: OrderedDocument,
  labels: ReadonlyMap<string, readonly OoxmlElement[]>,
  madeOver: readonly CheckpointMadeOver[],
  listedIds: readonly string[]
): boolean {
  if (madeOver.length === 0) {
    return true;
  }

  const holders = holdersOf(doc);
  const items = new Map<OrderedItem, ItemOver>();
  for (const { label, properties, records } of madeOver) {
    const held = (labels.get(String(label)) ?? []).filter((e) => holders.has(e));
    const element = held[0];
    const made = madeRecords(doc, records, listedIds);
    if (
      held.length !== 1 ||
      element === undefined ||
      !isWml(element, `${properties.localName}Change`) ||
      made === null ||
      !(holders.get(element) ?? []).every((holder) => noteOver(items, holder, properties, made))
    ) {
      return false;
    }
  }

  for (const [item, over] of items) {
    putBack(doc, item, over);
  }

  return true;
}

/** The listed records with their suggestions' ids, each a new order after every record. */
function madeRecords(
  doc: OrderedDocument,
  records: readonly CheckpointOverRecord[],
  listedIds: readonly string[]
): Made[] | null {
  const out: Made[] = [];
  for (const { suggestion: k, over: madeOver, untracked, write } of records) {
    const suggestion = listedIds[k];
    const over = madeOver.flatMap((j) => listedIds[j] ?? []);
    const status = suggestion === undefined ? undefined : doc.suggestions.get(suggestion)?.status;
    if (
      suggestion === undefined ||
      status !== (untracked === true ? 'pending' : 'accepted') ||
      over.length !== madeOver.length
    ) {
      return null;
    }

    out.push({ suggestion, seq: nextSeq(doc), over, untracked, write });
  }

  return out;
}

/**
 * Note in `items` what a holder gets back: the properties, and the records made over it, as the
 * document's records of its kind. A field's runs each give their own face, and the field's one
 * record takes the records once. False when the holder is not a record of those properties.
 */
function noteOver(
  items: Map<OrderedItem, ItemOver>,
  { story, item, record, piece }: Holder,
  properties: OoxmlElement,
  made: readonly Made[]
): boolean {
  let entry = items.get(item);
  if (entry === undefined) {
    entry = { story, fmts: new Map(), props: new Map(), faces: new Map() };
    items.set(item, entry);
  }

  if (record.kind === 'face') {
    const fmts = formatRecords(record, made);
    if (properties.localName !== 'rPr' || fmts === null) {
      return false;
    }

    entry.fmts.set(record, [
      piece === undefined ? { ...record, face: properties } : record,
      ...fmts,
    ]);
    if (piece !== undefined) {
      entry.faces.set(piece, new Map([[record.id, properties]]));
    }

    return true;
  }

  const props = propertyRecords(made);
  if (record.target !== properties.localName || props === null) {
    return false;
  }

  entry.props.set(record, [{ ...record, element: properties }, ...props]);
  return true;
}

/** The made records as format records on a run, each with an id of the holder's; null for another write. */
function formatRecords(holder: FaceRecord, made: readonly Made[]): OrderedFormatRecord[] | null {
  const out: OrderedFormatRecord[] = [];
  for (const [k, { suggestion, seq, over, untracked, write }] of made.entries()) {
    if (!('format' in write)) {
      return null;
    }

    const id = `${holder.id}>${k}`;
    const flag = untracked === true ? { untracked } : {};
    out.push({ kind: 'properties', id, suggestion, seq, ...write.format, over, ...flag });
  }

  return out;
}

/** The made records as a paragraph's or a section's records; null for a format. */
function propertyRecords(made: readonly Made[]): OrderedPropertyRecord[] | null {
  const out: OrderedPropertyRecord[] = [];
  for (const { suggestion, seq, over, write } of made) {
    if ('paragraphOp' in write) {
      out.push({
        kind: 'paragraphOp',
        suggestion,
        seq,
        target: 'pPr',
        op: write.paragraphOp,
        over,
      });
    } else if ('sectionOp' in write) {
      out.push({ kind: 'sectionOp', suggestion, seq, target: 'sectPr', op: write.sectionOp, over });
    } else {
      return null;
    }
  }

  return out;
}

/**
 * An item with each holder giving its properties and followed by the records made over it. A
 * field's runs get their own faces.
 */
function putBack(doc: OrderedDocument, item: OrderedItem, over: ItemOver): void {
  const { story, faces } = over;
  const fmts = item.fmts.flatMap((record) => over.fmts.get(record) ?? [record]);
  const props = item.props.flatMap((record) => over.props.get(record) ?? [record]);
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
