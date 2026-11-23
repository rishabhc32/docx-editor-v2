// A checkpoint's list definitions back to their pending suggestions (design D14).

import { withPart, type OoxmlPackage } from '../../../store/package/ooxml-package.ts';
import type { OoxmlNode } from '../../../store/package/ooxml-tree.ts';
import { withoutPart } from '../../../store/package/package-edit.ts';
import type { CheckpointList } from '../model/checkpoint-lists.ts';
import { orderedListNumber, type OrderedListResource } from '../model/resources.ts';
import type { OrderedDocument } from '../model/types.ts';
import { attributeOf, isWml, withChildren } from '../ooxml-nodes.ts';

/** A list definition taken out of a checkpoint's file, for its suggestion. */
export interface TakenList {
  /** The suggestion, by its place in the manifest's list. */
  readonly suggestion: number;
  readonly key: string;
  readonly definition: NonNullable<OrderedListResource['definition']>;
}

/** The place of the child `w:<name>` whose `w:<id>` is `value`, or -1. */
const placeOf = (children: readonly OoxmlNode[], name: string, id: string, value: string) =>
  children.findIndex(
    (child) => child.kind !== 'textValue' && isWml(child, name) && attributeOf(child, id) === value
  );

/**
 * A checkpoint's package without the list definitions the manifest lists, and each definition as
 * the room wrote it: its `w:abstractNum`, and the overrides of its `w:num`. A numbering part that
 * holds nothing else goes too, as the room wrote it only for them. Null when the file does not
 * hold one of them.
 */
export function withoutCheckpointLists(
  pkg: OoxmlPackage,
  lists: readonly CheckpointList[]
): { readonly pkg: OoxmlPackage; readonly taken: readonly TakenList[] } | null {
  if (lists.length === 0) {
    return { pkg, taken: [] };
  }

  const part = [...pkg.parts.values()].find((candidate) => isWml(candidate.root, 'numbering'));
  if (part === undefined) {
    return null;
  }

  const children = [...part.root.children];
  const taken: TakenList[] = [];
  for (const { suggestion, key } of lists) {
    const abstractAt = placeOf(
      children,
      'abstractNum',
      'abstractNumId',
      orderedListNumber(key, 'abstract')
    );
    const abstract = children[abstractAt];
    const numAt = placeOf(children, 'num', 'numId', orderedListNumber(key, 'num'));
    const num = children[numAt];
    if (
      abstract === undefined ||
      abstract.kind === 'textValue' ||
      num === undefined ||
      num.kind === 'textValue'
    ) {
      return null;
    }

    const overrides = num.children.flatMap((child) =>
      child.kind !== 'textValue' && isWml(child, 'lvlOverride') ? [child] : []
    );
    taken.push({ suggestion, key, definition: { abstract, overrides } });
    children.splice(Math.max(abstractAt, numAt), 1);
    children.splice(Math.min(abstractAt, numAt), 1);
  }

  const out = withPart(pkg, { ...part, root: withChildren(part.root, children) });
  if (children.some((child) => child.kind !== 'textValue')) {
    return { pkg: out, taken };
  }

  const removed = withoutPart(out, part.name);
  return removed.ok ? { pkg: removed.pkg, taken } : null;
}

/**
 * Give each taken list definition back to its suggestion, by the ids of the suggestions the
 * manifest lists (`listedIds`), as the resource the room held. False when a suggestion is not
 * listed.
 */
export function withCheckpointLists(
  doc: OrderedDocument,
  taken: readonly TakenList[],
  listedIds: readonly string[]
): boolean {
  for (const { suggestion, key, definition } of taken) {
    const id = listedIds[suggestion];
    if (id === undefined) {
      return false;
    }

    doc.resources.added.set(key, {
      kind: 'list',
      key,
      suggestion: id,
      numId: orderedListNumber(key, 'num'),
      abstractNumId: orderedListNumber(key, 'abstract'),
      levels: [],
      definition,
    });
  }

  return true;
}
