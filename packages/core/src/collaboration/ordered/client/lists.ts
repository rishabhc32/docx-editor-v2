// A checkpoint's list definitions and styles back to their suggestions (design D14).

import { withPart, type OoxmlPackage } from '../../../store/package/ooxml-package.ts';
import type { OoxmlElement, OoxmlNode, OoxmlPart } from '../../../store/package/ooxml-tree.ts';
import { withoutPart } from '../../../store/package/package-edit.ts';
import { styleElementResource } from '../edits/resources.ts';
import type { CheckpointList, CheckpointStyle } from '../model/checkpoint-lists.ts';
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

  const out = withPartChildren(pkg, part, children);
  return out === null ? null : { pkg: out, taken };
}

/** A style taken out of a checkpoint's file, for its suggestion. */
export interface TakenStyle {
  /** The suggestion, by its place in the manifest's list. */
  readonly suggestion: number;
  readonly key: string;
  readonly element: OoxmlElement;
}

/**
 * A checkpoint's package without the styles the manifest lists, and each `w:style` as the room
 * wrote it. A styles part that holds nothing else goes too. Null when the file does not hold one
 * of them.
 */
export function withoutCheckpointStyles(
  pkg: OoxmlPackage,
  styles: readonly CheckpointStyle[]
): { readonly pkg: OoxmlPackage; readonly taken: readonly TakenStyle[] } | null {
  if (styles.length === 0) {
    return { pkg, taken: [] };
  }

  const part = [...pkg.parts.values()].find((candidate) => isWml(candidate.root, 'styles'));
  if (part === undefined) {
    return null;
  }

  const children = [...part.root.children];
  const taken: TakenStyle[] = [];
  for (const { suggestion, key, styleId } of styles) {
    const at = placeOf(children, 'style', 'styleId', styleId);
    const element = children[at];
    if (element === undefined || element.kind === 'textValue') {
      return null;
    }

    taken.push({ suggestion, key, element });
    children.splice(at, 1);
  }

  const out = withPartChildren(pkg, part, children);
  return out === null ? null : { pkg: out, taken };
}

/** A package with a part's children; a part left with no element goes. Null when it cannot. */
function withPartChildren(
  pkg: OoxmlPackage,
  part: OoxmlPart,
  children: readonly OoxmlNode[]
): OoxmlPackage | null {
  const out = withPart(pkg, { ...part, root: withChildren(part.root, children) });
  if (children.some((child) => child.kind !== 'textValue')) {
    return out;
  }

  const removed = withoutPart(out, part.name);
  return removed.ok ? removed.pkg : null;
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

/**
 * Give each taken style back to its suggestion, by the ids of the suggestions the manifest lists
 * (`listedIds`), as a pasted style is held. False when a suggestion is not listed or a style has
 * no id or no known type.
 */
export function withCheckpointStyles(
  doc: OrderedDocument,
  taken: readonly TakenStyle[],
  listedIds: readonly string[]
): boolean {
  for (const { suggestion, key, element } of taken) {
    const id = listedIds[suggestion];
    const resource = id === undefined ? null : styleElementResource(key, id, element);
    if (resource === null) {
      return false;
    }

    doc.resources.added.set(key, resource);
  }

  return true;
}
