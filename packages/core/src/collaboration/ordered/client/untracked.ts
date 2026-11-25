// A checkpoint's records with no change record, and what the room holds of a record with one,
// back in the document (design D14).

import type { OoxmlElement, OoxmlNode } from '../../../store/package/ooxml-tree.ts';
import {
  UNTRACKED_CONTAINERS,
  type CheckpointChangeRecords,
  type CheckpointUntracked,
} from '../model/checkpoint-untracked.ts';
import { createOrderedItem, storyOfAddress } from '../model/document.ts';
import { replaceItems } from '../model/mutations.ts';
import {
  tokenSpace,
  type OrderedDocument,
  type OrderedElementRecord,
  type OrderedItem,
  type OrderedPropertyRecord,
  type OrderedPropertyTarget,
  type OrderedStory,
} from '../model/types.ts';
import { withMergedChildren } from '../compose/properties.ts';
import { changedKeys } from '../edits/table-nodes.ts';
import { childWhere, isWml, withChildren } from '../ooxml-nodes.ts';

/**
 * Put back in the document loaded from a checkpoint each container's properties that a record
 * changes with no change record: the container's own properties of the target are the manifest's
 * again, and the manifest's records, each with its suggestion's order, take the place of those
 * loaded from the file (but the records a takeover put back), so compose draws them over those
 * properties as the room did. `labels` gives the elements that carry each label. False when a
 * label is on no container the load holds whose properties the target is (a cell for a table's
 * cells), or on two (a text box's other copy holds none).
 */
export function withCheckpointUntracked(
  doc: OrderedDocument,
  labels: ReadonlyMap<string, readonly OoxmlElement[]>,
  untracked: readonly CheckpointUntracked[],
  listedIds: readonly string[]
): boolean {
  for (const entry of untracked) {
    const container = UNTRACKED_CONTAINERS.get(entry.target);
    const found =
      container === undefined ? undefined : labelledOpen(doc, labels, entry.label, container);
    const shell = found?.open.node;
    if (
      found === undefined ||
      shell === null ||
      shell === undefined ||
      shell.kind === 'textValue'
    ) {
      return false;
    }

    const records = restoredRecords(doc, entry, listedIds, labels);
    if (records === null) {
      return false;
    }

    const item = restoredOpen(found.open, shell, entry, records);
    replaceItems(doc, found.story, [found.open], [item]);
  }

  return true;
}

/**
 * Each listed change record's facts back on the record the file loaded: the children it sets and
 * the cells it holds. The loaded record is the one of the target with the suggestion and the one
 * it was taken from, in their order on the container; so this runs after the takeovers. False
 * when the container, a cell or a record is not there.
 */
export function withCheckpointChangeRecords(
  doc: OrderedDocument,
  labels: ReadonlyMap<string, readonly OoxmlElement[]>,
  changeRecords: readonly CheckpointChangeRecords[],
  listedIds: readonly string[]
): boolean {
  for (const { label, target, records } of changeRecords) {
    const container = UNTRACKED_CONTAINERS.get(target);
    const found = container === undefined ? undefined : labelledOpen(doc, labels, label, container);
    if (found === undefined) {
      return false;
    }

    const props = [...found.open.props];
    const given = new Set<number>();
    for (const record of records) {
      const suggestion = listedIds[record.suggestion];
      const replaced = record.replaced === undefined ? undefined : listedIds[record.replaced];
      const at = props.findIndex(
        (r, k) =>
          !given.has(k) &&
          r.kind === 'element' &&
          r.target === target &&
          r.suggestion === suggestion &&
          r.replaced === replaced
      );
      const loaded = props[at];
      const tableCells =
        record.tableCells === undefined ? undefined : cellSpaces(doc, labels, record.tableCells);
      if (
        (record.replaced !== undefined && replaced === undefined) ||
        loaded?.kind !== 'element' ||
        tableCells === null
      ) {
        return false;
      }

      // A direct change made over the record shows in the file's record and in no room record
      // after a fold: the file's record gives back what differs from the container's own
      const own = found.open.node;
      const was =
        own && own.kind !== 'textValue' ? childWhere(own, (c) => isWml(c, target)) : undefined;
      const merge =
        record.merge === undefined || loaded.fileRecord === undefined || loaded.element === null
          ? record.merge
          : [...new Set([...record.merge, ...changedKeys(was, loaded.element)])];
      given.add(at);
      props[at] = { ...loaded, merge, tableCells };
    }

    replaceItems(doc, found.story, [found.open], [createOrderedItem({ ...found.open, props })]);
  }

  return true;
}

/** The open item of the container of this local name that a label is on, and its story. */
function labelledOpen(
  doc: OrderedDocument,
  labels: ReadonlyMap<string, readonly OoxmlElement[]>,
  label: number,
  container: string
): { readonly story: OrderedStory; readonly open: OrderedItem } | undefined {
  const found = (labels.get(String(label)) ?? []).flatMap((element) => {
    const space = tokenSpace(doc.file, element.id);
    const story = storyOfAddress(doc, { space, offset: 0 });
    const open = story?.spaces.get(space)?.[0];
    return story === null || open?.kind !== 'open' || !isWml(element, container)
      ? []
      : [{ story, open }];
  });
  return found.length === 1 ? found[0] : undefined;
}

/** The spaces of the cells the labels name; null when one is not there. */
function cellSpaces(
  doc: OrderedDocument,
  labels: ReadonlyMap<string, readonly OoxmlElement[]>,
  cells: readonly number[]
): string[] | null {
  const spaces: string[] = [];
  for (const cell of cells) {
    const space = labelledOpen(doc, labels, cell, 'tc')?.open.space;
    if (space === undefined) {
      return null;
    }

    spaces.push(space);
  }

  return spaces;
}

/**
 * The manifest's records of a container's properties as the document's, each with its
 * suggestion's order and its table's cells; null when a suggestion or a cell is not there.
 */
function restoredRecords(
  doc: OrderedDocument,
  { target, records }: CheckpointUntracked,
  listedIds: readonly string[],
  labels: ReadonlyMap<string, readonly OoxmlElement[]>
): OrderedElementRecord[] | null {
  const restored: OrderedElementRecord[] = [];
  for (const record of records) {
    const id = listedIds[record.suggestion];
    const seq = id === undefined ? undefined : doc.suggestions.get(id)?.seq;
    const replaced = record.replaced === undefined ? undefined : listedIds[record.replaced];
    const over = record.over?.flatMap((k) => listedIds[k] ?? []);
    if (
      id === undefined ||
      seq === undefined ||
      (record.replaced !== undefined && replaced === undefined) ||
      over?.length !== record.over?.length
    ) {
      return null;
    }

    const cells = cellSpaces(doc, labels, record.tableCells ?? []);
    if (cells === null) {
      return null;
    }

    restored.push({
      ...record,
      kind: 'element',
      suggestion: id,
      seq,
      target,
      tableCells: record.tableCells === undefined ? undefined : cells,
      replaced,
      over,
    });
  }

  return restored;
}

/**
 * A container's open item with its properties of the target as the manifest gives them, and the
 * restored records in the place of those loaded from the file (but the records a takeover put
 * back), in their order.
 */
function restoredOpen(
  open: OrderedItem,
  shell: OoxmlElement,
  { target, properties, records }: CheckpointUntracked,
  restored: readonly OrderedElementRecord[]
): OrderedItem {
  // Records that set some children only give back their own: the others are the file's
  const keys = records.every((r) => r.merge !== undefined && r.untracked === true)
    ? records.flatMap((r) => r.merge ?? [])
    : undefined;
  const node = withChildren(shell, withProperties(shell.children, target, properties, keys));
  const taken = (r: OrderedPropertyRecord) => r.kind === 'element' && r.replaced !== undefined;
  const props = [...open.props.filter((r) => r.target !== target || taken(r)), ...restored].sort(
    (a, b) => a.seq - b.seq
  );
  return createOrderedItem({ ...open, node, props });
}

/**
 * A shell's children with its properties of `target` in their place (first when it had none):
 * `properties`, or only their children `keys` names over the shell's own.
 */
function withProperties(
  children: readonly OoxmlNode[],
  target: OrderedPropertyTarget,
  properties: OoxmlElement | undefined,
  keys: readonly string[] | undefined
): OoxmlNode[] {
  const at = children.findIndex((c) => isWml(c, target));
  const loaded = children[at];
  let restored = properties;
  if (keys !== undefined && loaded !== undefined && loaded.kind !== 'textValue') {
    const kept = withMergedChildren(target, properties ?? null, keys, loaded.children);

    // With no properties of its own, the container's held only what the records give
    restored =
      properties === undefined && kept.length === 0 ? undefined : withChildren(loaded, kept);
  }

  if (restored === undefined) {
    return children.filter((c) => !isWml(c, target));
  }

  return at < 0 ? [restored, ...children] : children.map((c, k) => (k === at ? restored : c));
}
