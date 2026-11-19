// Word revision ids (`w:id`). Compose writes one id for each suggestion and role, the same one
// every time and on every copy, whenever the copy composes.

import {
  WML_NAMESPACE_URI,
  type OoxmlAttribute,
  type OoxmlElement,
} from '../../../store/package/ooxml-tree.ts';
import { fnv1a32 } from '../../../store/package/para-id.ts';
import { attributeOf } from '../ooxml-nodes.ts';
import { isHidden, storyOfItem } from './document.ts';
import type { OrderedDocument, OrderedItem, OrderedStory } from './types.ts';

/** The place a revision mark takes, which decides whether two marks share one `w:id`. */
export const REVISION_ROLE = {
  /** Inserted runs: one id across paragraphs. */
  insert: 'ins',
  /** Struck runs: one id for each paragraph (the mark item's space names it). */
  deleteIn: (paragraphKey: string) => `del@${paragraphKey}`,
  moveFromIn: (paragraphKey: string) => `move-from@${paragraphKey}`,
  moveTo: 'move-to',
  moveFromRange: 'moveFrom-range',
  moveToRange: 'moveTo-range',
  markInsert: 'mark-ins',
  markDelete: 'mark-del',
  /**
   * A notes part's paragraph marks: a struck note's body takes its reference's suggestion (D20),
   * and Word gives its marks their own `w:id`s, apart from the reference paragraph's mark.
   */
  noteMarkInsert: 'note-mark-ins',
  noteMarkDelete: 'note-mark-del',
  /**
   * A notes part's inserted runs: a new note's text takes its reference's suggestion (D20), and
   * Word gives its `w:ins` a `w:id` of its own, apart from the reference's.
   */
  noteInsert: 'note-ins',
  format: 'fmt',
  paragraphProperties: 'ppr',
  rowInsert: 'row',
  rowDelete: 'row-del',
  tableProperties: 'table-props',
  sectionProperties: 'sect-props',
  /** A file's own numbering record in a paragraph's `w:numPr`: `w:numberingChange` or `w:ins`. */
  numbering: 'numbering',
} as const;

const key = (suggestion: string, role: string) => `${suggestion}/${role}`;

/**
 * Minted ids: from 10^9 up to 2^31 - 2, inside a signed 32-bit integer as Word reads it, where a
 * file's dense ids do not reach. Two minted ids meet with a chance of about one in 1.1 billion
 * for each pair; they then share a `w:id`, which Word reads as two marks still, and the engine
 * joins only when author and date also agree.
 */
const MINTED_FROM = 1_000_000_000;
const MINTED_SPAN = 1_147_483_647;

/** The minted id a hash gives, in the range above. */
export const mintedNumber = (hash: number): number => MINTED_FROM + (hash % MINTED_SPAN);

/**
 * The `w:id` for a suggestion's role: the file's own for the role it was read in (load seeds the
 * table), or the one its suggestion id and role give, past every id the file holds. A suggestion
 * a change made has its id minted from the change id, so its marks get the same `w:id` on every
 * copy. A pure function of the document's suggestion ids and the file: no counter, so the order
 * in which a copy composes does not change it.
 */
export function revisionIdFor(doc: OrderedDocument, suggestion: string, role: string): string {
  const table = doc.revisionIds;
  const k = key(suggestion, role);
  const found = table.byRole.get(k);
  if (found !== undefined) return found;

  // A file holds fewer ids than the span, so the probe ends
  for (let attempt = 0; ; attempt += 1) {
    const hash = fnv1a32(
      attempt === 0 ? `ordered revision ${k}` : `ordered revision ${k} ${attempt}`
    );
    const id = String(mintedNumber(hash));
    if (!table.fileIds.has(id)) return id;
  }
}

/**
 * Keep a file's revision mark for a suggestion's role: its `w:id`, and its other attributes
 * past the id and a move range's name (the first mark seen wins).
 */
export function seedRevisionId(
  doc: OrderedDocument,
  suggestion: string,
  role: string,
  mark: OoxmlElement
): void {
  const id = mark.attributes.find(
    (a) => a.namespaceUri === WML_NAMESPACE_URI && a.localName === 'id'
  );
  if (id === undefined) return;
  const k = key(suggestion, role);
  if (!doc.revisionIds.byRole.has(k)) {
    doc.revisionIds.byRole.set(k, id.value);
    doc.revisionIds.attributesByRole.set(
      k,
      mark.attributes.filter(
        (a) =>
          !(
            a.namespaceUri === WML_NAMESPACE_URI &&
            (a.localName === 'id' || a.localName === 'name')
          )
      )
    );
  }

  reserveFileRevisionId(doc, id.value);
}

/**
 * Give `to` the file's marks `from` was read with, for each role `to` has none of: the marks a
 * load joins into another suggestion keep their `w:id` and attributes.
 */
export function adoptRevisionRoles(doc: OrderedDocument, from: string, to: string): void {
  const table = doc.revisionIds;
  const prefix = key(from, '');
  for (const [k, id] of [...table.byRole]) {
    if (!k.startsWith(prefix)) continue;
    const target = key(to, k.slice(prefix.length));
    if (table.byRole.has(target)) continue;
    table.byRole.set(target, id);
    const attributes = table.attributesByRole.get(k);
    if (attributes !== undefined) table.attributesByRole.set(target, attributes);
  }
}

/** The file's own attributes of a suggestion's mark for a role, past its id, if it had one. */
export function seededAttributes(
  doc: OrderedDocument,
  suggestion: string,
  role: string
): readonly OoxmlAttribute[] | undefined {
  return doc.revisionIds.attributesByRole.get(key(suggestion, role));
}

/** A `w:id` the loaded parts hold: minted ids pass over it. */
export function reserveFileRevisionId(doc: OrderedDocument, id: string): void {
  doc.revisionIds.fileIds.add(id);
}

/** Roles whose `w:id` is one for the whole suggestion. */
const WHOLE_ROLES: readonly string[] = [
  REVISION_ROLE.insert,
  REVISION_ROLE.moveTo,
  REVISION_ROLE.moveFromRange,
  REVISION_ROLE.moveToRange,
  REVISION_ROLE.markInsert,
  REVISION_ROLE.markDelete,
  REVISION_ROLE.noteMarkInsert,
  REVISION_ROLE.noteMarkDelete,
  REVISION_ROLE.noteInsert,
  REVISION_ROLE.format,
  REVISION_ROLE.paragraphProperties,
  REVISION_ROLE.rowInsert,
  REVISION_ROLE.rowDelete,
  REVISION_ROLE.tableProperties,
  REVISION_ROLE.sectionProperties,
  REVISION_ROLE.numbering,
];

/** The mark that ends the paragraph compose writes an item in: the next mark that shows. */
function paragraphKeyOf(
  doc: OrderedDocument,
  story: OrderedStory,
  item: OrderedItem,
  known: Map<OrderedItem, string>
): string | null {
  const walked: OrderedItem[] = [];
  let key: string | null = null;
  for (let at: OrderedItem | null = item; at; at = story.tree.next(at)) {
    const cached = known.get(at);
    if (cached !== undefined) {
      key = cached;
      break;
    }

    walked.push(at);
    if (at.kind === 'mark' && !isHidden(doc, at)) {
      key = at.space;
      break;
    }
  }

  if (key !== null) {
    for (const at of walked) {
      known.set(at, key);
    }
  }

  return key;
}

/**
 * The `w:id`s of the file's own change records a suggestion holds. Word gives each record of
 * one change an id of its own (a row height: the table's, the grid's and each cell's record),
 * and the editor that opened the file names them by those ids.
 */
export function fileRecordIdsOf(doc: OrderedDocument, suggestion: string): string[] {
  const ids: string[] = [];
  for (const item of doc.labelled.get(suggestion) ?? []) {
    for (const record of item.props) {
      if (record.suggestion !== suggestion || record.kind !== 'element') {
        continue;
      }

      const id = record.fileRecord && attributeOf(record.fileRecord, 'id');
      if (id !== undefined) {
        ids.push(id);
      }
    }
  }

  return ids;
}

/**
 * Every `w:id` compose writes for a suggestion's marks, its file records' own, and those a room
 * checkpoint listed for it.
 */
export function revisionIdsOf(doc: OrderedDocument, suggestion: string): Set<string> {
  const ids = new Set(WHOLE_ROLES.map((role) => revisionIdFor(doc, suggestion, role)));
  const listed = doc.revisionIds.listed.get(suggestion) ?? [];
  for (const id of [...fileRecordIdsOf(doc, suggestion), ...listed]) {
    ids.add(id);
  }

  const known = new Map<OrderedItem, string>();
  for (const item of doc.labelled.get(suggestion) ?? []) {
    if (!item.dels.includes(suggestion)) {
      continue;
    }

    const story = storyOfItem(doc, item);
    const key = story === null ? null : paragraphKeyOf(doc, story, item, known);
    if (key === null) {
      continue;
    }

    ids.add(revisionIdFor(doc, suggestion, REVISION_ROLE.deleteIn(key)));
    ids.add(revisionIdFor(doc, suggestion, REVISION_ROLE.moveFromIn(key)));
  }

  return ids;
}
