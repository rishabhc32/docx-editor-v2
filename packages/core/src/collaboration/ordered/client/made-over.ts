// A checkpoint's records of the change records with a record made over them, back in the
// document (design D14).

import type { OoxmlElement } from '../../../store/package/ooxml-tree.ts';
import type {
  CheckpointMadeOver,
  CheckpointOverRecord,
  CheckpointOverTarget,
} from '../model/checkpoint-over.ts';
import type { CheckpointSuggestion } from '../model/checkpoint-suggestions.ts';
import { createOrderedItem } from '../model/document.ts';
import { nextSeq, relabel, replaceItems } from '../model/mutations.ts';
import type {
  OrderedDocument,
  OrderedFormatRecord,
  OrderedPropertyRecord,
  OrderedPropertyTarget,
} from '../model/types.ts';
import { isWml } from '../ooxml-nodes.ts';
import { holdersOf, type Holder } from './takeovers.ts';

/** A listed record with its suggestions' ids and its new order. */
type Restored = Omit<
  CheckpointOverRecord,
  'suggestion' | 'replaced' | 'replacedEarlier' | 'over'
> & {
  readonly suggestion: string;
  readonly seq: number;
  readonly replaced?: string;
  readonly replacedEarlier?: readonly string[];
  readonly over?: readonly string[];
};

/**
 * Put back in the document loaded from a checkpoint the records of each change record the
 * manifest names: on each item whose loaded record holds it, the manifest's records take the
 * place of the item's records of those properties (the loaded one and those a takeover put
 * back), in the order they apply, each with a new order after every record of the document.
 * `labels` gives the elements that carry each label, `suggestions` the manifest's list. False
 * when an entry does not fit the file: its label on no change record of its properties that a
 * loaded record holds, or on two; a suggestion not listed; a record on its author's own new words
 * of a direct edit or a rejected suggestion; a record made over a pending one with no change
 * record of its own of no accepted direct edit.
 */
export function withCheckpointMadeOver(
  doc: OrderedDocument,
  labels: ReadonlyMap<string, readonly OoxmlElement[]>,
  madeOver: readonly CheckpointMadeOver[],
  listedIds: readonly string[],
  suggestions: readonly CheckpointSuggestion[]
): boolean {
  if (madeOver.length === 0) {
    return true;
  }

  const holders = holdersOf(doc);
  for (const { label, target, records } of madeOver) {
    const held = (labels.get(String(label)) ?? []).filter((e) => holders.has(e));
    const element = held[0];
    const restored = restoredRecords(doc, records, listedIds, suggestions);
    if (
      held.length !== 1 ||
      element === undefined ||
      !isWml(element, `${target}Change`) ||
      restored === null ||
      !(holders.get(element) ?? []).every((holder) => putBack(doc, holder, target, restored))
    ) {
      return false;
    }
  }

  return true;
}

/** The listed records with their suggestions' ids, each a new order; null when one does not fit. */
function restoredRecords(
  doc: OrderedDocument,
  records: readonly CheckpointOverRecord[],
  listedIds: readonly string[],
  suggestions: readonly CheckpointSuggestion[]
): Restored[] | null {
  const out: Restored[] = [];
  for (const record of records) {
    const suggestion = listedIds[record.suggestion];
    const replaced = record.replaced === undefined ? undefined : listedIds[record.replaced];
    const replacedEarlier = record.replacedEarlier?.flatMap((k) => listedIds[k] ?? []);
    const over = record.over?.flatMap((k) => listedIds[k] ?? []);
    const made = suggestion === undefined ? undefined : doc.suggestions.get(suggestion);
    const direct = suggestions[record.suggestion]?.direct === true;

    // A record on its author's own new words is decided with them, never a direct edit's, and a
    // rejected one is not drawn; one made over a pending record with no change record of its own
    // is an accepted direct edit's
    if (
      suggestion === undefined ||
      made === undefined ||
      (record.replaced !== undefined && replaced === undefined) ||
      replacedEarlier?.length !== record.replacedEarlier?.length ||
      over?.length !== record.over?.length ||
      (record.untracked === true && (direct || made.status === 'rejected')) ||
      (record.untracked !== true &&
        record.over !== undefined &&
        (!direct || made.status !== 'accepted'))
    ) {
      return null;
    }

    out.push({ ...record, suggestion, replaced, replacedEarlier, over, seq: nextSeq(doc) });
  }

  return out;
}

/**
 * A holder's item with the restored records in the place of its records of `target`. A field's
 * run gives its own face to each record that gives one, by the record's place. False when the
 * holder is not a record of those properties.
 */
function putBack(
  doc: OrderedDocument,
  { story, item, record, piece }: Holder,
  target: CheckpointOverTarget,
  restored: readonly Restored[]
): boolean {
  // A record that gives properties, as a load reads one, keeps the file's change record, which
  // states the properties before it, as the room's does: also one another suggestion took over
  const fileRecordOf = (r: Restored) => (r.write === undefined ? record.fileRecord : undefined);

  if (record.kind !== 'face') {
    if (record.target !== target) {
      return false;
    }

    const props = restored.map((r) => propertyRecord(r, record.target, fileRecordOf(r)));
    relabel(doc, story, item, {
      props: [...item.props.filter((r) => r.target !== target), ...props],
    });
    return true;
  }

  if (target !== 'rPr') {
    return false;
  }

  const fmts = restored.map((r, k) => formatRecord(r, k, fileRecordOf(r)));
  if (piece === undefined || item.pieces === undefined) {
    relabel(doc, story, item, { fmts });
    return true;
  }

  const faces = new Map(item.pieces[piece]?.recordFaces);
  for (const restoredRecord of fmts) {
    if (restoredRecord.kind === 'face') {
      faces.set(restoredRecord.id, restoredRecord.face);
    }
  }

  const pieces = item.pieces.map((p, k) => (k === piece ? { ...p, recordFaces: faces } : p));
  replaceItems(doc, story, [item], [createOrderedItem({ ...item, fmts, pieces })]);
  return true;
}

/** A restored record of a run's properties: its format, or the face it gives with `fileRecord`. */
function formatRecord(
  restored: Restored,
  k: number,
  fileRecord: OoxmlElement | undefined
): OrderedFormatRecord {
  const { suggestion, seq, replaced, replacedEarlier, over, untracked, element, write } = restored;
  const fields = { id: `${suggestion}>${k}`, suggestion, seq, replaced, replacedEarlier };
  if (write !== undefined && 'format' in write) {
    return { kind: 'properties', ...fields, ...write.format, over, untracked };
  }

  return { kind: 'face', ...fields, face: element, fileRecord };
}

/**
 * A restored record of a paragraph mark's or a section's properties: its write, or the element it
 * gives with `fileRecord`.
 */
function propertyRecord(
  restored: Restored,
  target: OrderedPropertyTarget,
  fileRecord: OoxmlElement | undefined
): OrderedPropertyRecord {
  const { suggestion, seq, replaced, replacedEarlier, over, untracked, write } = restored;
  const fields = { suggestion, seq, replaced, replacedEarlier, over, untracked };
  if (write !== undefined && 'paragraphOp' in write) {
    return { kind: 'paragraphOp', ...fields, target: 'pPr', op: write.paragraphOp };
  }

  if (write !== undefined && 'sectionOp' in write) {
    const { keptOnReject } = restored;
    return { kind: 'sectionOp', ...fields, target: 'sectPr', op: write.sectionOp, keptOnReject };
  }

  const { element, merge, recordUnchanged } = restored;
  return { kind: 'element', ...fields, target, element, fileRecord, merge, recordUnchanged };
}
