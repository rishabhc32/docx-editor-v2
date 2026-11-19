// A checkpoint's records with no change record back in the document (design D14).

import type { OoxmlElement, OoxmlNode } from '../../../store/package/ooxml-tree.ts';
import { containerKey } from '../model/checkpoint-takeovers.ts';
import type { CheckpointUntracked } from '../model/checkpoint-untracked.ts';
import { createOrderedItem } from '../model/document.ts';
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
import { isWml, withChildren } from '../ooxml-nodes.ts';

/**
 * Put back in the document loaded from a checkpoint each container's properties that a record
 * changes with no change record: the container's own properties of the target are the manifest's
 * again, and the manifest's records, each with its suggestion's order, take the place of those
 * loaded from the file (but the records a takeover put back), so compose draws them over those
 * properties as the room did. False when a container or a cell is not at its place.
 */
export function withCheckpointUntracked(
  doc: OrderedDocument,
  places: ReadonlyMap<string, OoxmlElement>,
  untracked: readonly CheckpointUntracked[],
  listed: readonly string[]
): boolean {
  const opened = (part: string, container: string, occurrence: number) => {
    const element = places.get(containerKey(part, container, occurrence));
    return element === undefined ? null : openItemOf(doc, tokenSpace(doc.file, element.id));
  };
  for (const { part, container, occurrence, target, properties, records } of untracked) {
    const found = opened(part, container, occurrence);
    const shell = found?.open.node;
    if (found === null || shell === null || shell === undefined || shell.kind === 'textValue') {
      return false;
    }

    const restored: OrderedElementRecord[] = [];
    for (const { suggestion, element, untracked, merge, recordUnchanged, tableCells } of records) {
      const id = listed[suggestion];
      const cells = (tableCells ?? []).map((cell) => opened(part, 'tc', cell)?.open.space);
      const seq = id === undefined ? undefined : doc.suggestions.get(id)?.seq;
      if (id === undefined || seq === undefined || cells.includes(undefined)) {
        return false;
      }

      restored.push({
        kind: 'element',
        suggestion: id,
        seq,
        target,
        element,
        ...(untracked === undefined ? {} : { untracked }),
        ...(merge === undefined ? {} : { merge }),
        ...(recordUnchanged === undefined ? {} : { recordUnchanged }),
        ...(tableCells === undefined ? {} : { tableCells: cells.filter((c) => c !== undefined) }),
      });
    }

    // Records that set some children only give back their own: the others are the file's
    const keys = records.every((r) => r.merge !== undefined && r.untracked === true)
      ? records.flatMap((r) => r.merge ?? [])
      : undefined;
    const { story, open } = found;
    const node = withChildren(shell, withProperties(shell.children, target, properties, keys));
    const taken = (r: OrderedPropertyRecord) => r.kind === 'element' && r.replaced !== undefined;
    const props = [...open.props.filter((r) => r.target !== target || taken(r)), ...restored].sort(
      (a, b) => a.seq - b.seq
    );
    replaceItems(doc, story, [open], [createOrderedItem({ ...open, node, props })]);
  }

  return true;
}

/** The open item of a loaded container's space, and its story. */
function openItemOf(
  doc: OrderedDocument,
  space: string
): { readonly story: OrderedStory; readonly open: OrderedItem } | null {
  for (const story of doc.stories.values()) {
    const open = story.spaces.get(space)?.[0];
    if (open?.kind === 'open') {
      return { story, open };
    }
  }

  return null;
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
