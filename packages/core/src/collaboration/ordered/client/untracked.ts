// A checkpoint's records with no change record back in the document (design D14).

import type { OoxmlElement, OoxmlNode } from '../../../store/package/ooxml-tree.ts';
import { UNTRACKED_CONTAINERS, type CheckpointUntracked } from '../model/checkpoint-untracked.ts';
import { createOrderedItem, storyOfAddress } from '../model/document.ts';
import { replaceItems } from '../model/mutations.ts';
import {
  tokenSpace,
  type OrderedDocument,
  type OrderedElementRecord,
  type OrderedItem,
  type OrderedPropertyRecord,
  type OrderedPropertyTarget,
} from '../model/types.ts';
import { withMergedChildren } from '../compose/properties.ts';
import { isWml, withChildren } from '../ooxml-nodes.ts';

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
  // The open item of the container of this local name that a label is on, and its story
  const opened = (label: number, container: string) => {
    const found = (labels.get(String(label)) ?? []).flatMap((element) => {
      const space = tokenSpace(doc.file, element.id);
      const story = storyOfAddress(doc, { space, offset: 0 });
      const open = story?.spaces.get(space)?.[0];
      return story === null || open?.kind !== 'open' || !isWml(element, container)
        ? []
        : [{ story, open }];
    });
    return found.length === 1 ? found[0] : undefined;
  };
  for (const entry of untracked) {
    const container = UNTRACKED_CONTAINERS.get(entry.target);
    const found = container === undefined ? undefined : opened(entry.label, container);
    const shell = found?.open.node;
    if (
      found === undefined ||
      shell === null ||
      shell === undefined ||
      shell.kind === 'textValue'
    ) {
      return false;
    }

    const cellOf = (cell: number) => opened(cell, 'tc')?.open.space;
    const records = restoredRecords(doc, entry, listedIds, cellOf);
    if (records === null) {
      return false;
    }

    const item = restoredOpen(found.open, shell, entry, records);
    replaceItems(doc, found.story, [found.open], [item]);
  }

  return true;
}

/**
 * The manifest's records of a container's properties as the document's, each with its
 * suggestion's order and its table's cells; null when a suggestion or a cell is not there.
 */
function restoredRecords(
  doc: OrderedDocument,
  { target, records }: CheckpointUntracked,
  listedIds: readonly string[],
  cellOf: (cell: number) => string | undefined
): OrderedElementRecord[] | null {
  const restored: OrderedElementRecord[] = [];
  for (const record of records) {
    const id = listedIds[record.suggestion];
    const seq = id === undefined ? undefined : doc.suggestions.get(id)?.seq;
    const replaced = record.replaced === undefined ? undefined : listedIds[record.replaced];
    if (
      id === undefined ||
      seq === undefined ||
      (record.replaced !== undefined && replaced === undefined)
    ) {
      return null;
    }

    const cells: string[] = [];
    for (const cell of record.tableCells ?? []) {
      const space = cellOf(cell);
      if (space === undefined) {
        return null;
      }

      cells.push(space);
    }

    restored.push({
      ...record,
      kind: 'element',
      suggestion: id,
      seq,
      target,
      tableCells: record.tableCells === undefined ? undefined : cells,
      replaced,
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
    restored = withChildren(
      loaded,
      withMergedChildren(target, properties ?? null, keys, loaded.children)
    );
  }

  if (restored === undefined) {
    return children.filter((c) => !isWml(c, target));
  }

  return at < 0 ? [restored, ...children] : children.map((c, k) => (k === at ? restored : c));
}
