// A checkpoint's list definitions and styles back to their suggestions (design D14).

import { isRelationshipReference } from '../../../store/package/canonical-media-references.ts';
import { withPart, type OoxmlPackage } from '../../../store/package/ooxml-package.ts';
import type { OoxmlElement, OoxmlNode, OoxmlPart } from '../../../store/package/ooxml-tree.ts';
import { partNameKey } from '../../../store/package/opc-names.ts';
import { withoutPart } from '../../../store/package/package-edit.ts';
import { resolveRelationship } from '../../../store/package/relationships.ts';
import { RELATIONSHIP_TYPES } from '../compose/resources.ts';
import { styleElementResource } from '../edits/resources.ts';
import type { CheckpointList, CheckpointStyle } from '../model/checkpoint-lists.ts';
import {
  orderedListNumber,
  type OrderedListBullet,
  type OrderedListResource,
} from '../model/resources.ts';
import type { OrderedDocument } from '../model/types.ts';
import { attributeOf, eachElement, isWml, withChildren } from '../ooxml-nodes.ts';
import {
  numberingPartOf,
  placeOf,
  withCheckpointGiven,
  withoutRelationship,
  type ListedMedia,
} from './checkpoint-resources.ts';

/** A list definition taken out of a checkpoint's file, for its suggestion. */
export interface TakenList {
  /** The suggestion, by its place in the manifest's list. */
  readonly suggestion: number;
  readonly key: string;
  readonly definition: NonNullable<OrderedListResource['definition']>;
  readonly bullets: readonly OrderedListBullet[];
}

/**
 * A checkpoint's package without the list definitions the manifest lists, and each definition as
 * the room wrote it: its `w:abstractNum`, the overrides of its `w:num`, and the picture bullets its
 * levels name, each with its images, the numbering part's relationships to the media `media`
 * lists. A numbering part that holds nothing else goes too, as the room wrote it only for them.
 * Null when the file does not hold one of them.
 */
export function withoutCheckpointLists(
  pkg: OoxmlPackage,
  lists: readonly CheckpointList[],
  media: ReadonlyMap<string, ListedMedia>
): { readonly pkg: OoxmlPackage; readonly taken: readonly TakenList[] } | null {
  if (lists.length === 0) {
    return { pkg, taken: [] };
  }

  const part = numberingPartOf(pkg);
  if (part === undefined) {
    return null;
  }

  let next = pkg;
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
    const bullets: OrderedListBullet[] = [];
    for (const id of bulletIdsOf(abstract)) {
      const bullet = children[placeOf(children, 'numPicBullet', 'numPicBulletId', id)];
      const out = bullet === undefined ? null : withoutBullet(next, part.name, bullet, media);
      if (out === null) {
        return null;
      }

      next = out.pkg;
      bullets.push({ id, element: out.element, images: out.images });
      children.splice(children.indexOf(out.element), 1);
    }

    taken.push({ suggestion, key, definition: { abstract, overrides }, bullets });
    children.splice(children.indexOf(abstract), 1);
    children.splice(children.indexOf(num), 1);
  }

  const out = withPartChildren(next, part, children);
  return out === null ? null : { pkg: out, taken };
}

/** The picture bullet ids a list definition's levels name, each once, in order. */
function bulletIdsOf(abstract: OoxmlElement): string[] {
  const ids = new Set<string>();
  eachElement(abstract, (node) => {
    const id = isWml(node, 'lvlPicBulletId') ? attributeOf(node, 'val') : undefined;
    if (id !== undefined) {
      ids.add(id);
    }
  });
  return [...ids];
}

/**
 * The package without a picture bullet's images, the numbering part's image relationships to the
 * media `media` lists, and the bullet with each image's relationship id and media key. Null when
 * an image is not one of them.
 */
function withoutBullet(
  pkg: OoxmlPackage,
  numbering: string,
  bullet: OoxmlNode,
  media: ReadonlyMap<string, ListedMedia>
): {
  readonly pkg: OoxmlPackage;
  readonly element: OoxmlElement;
  readonly images: OrderedListBullet['images'];
} | null {
  if (bullet.kind === 'textValue') {
    return null;
  }

  const ids: string[] = [];
  eachElement(bullet, (node) => {
    for (const entry of node.attributes) {
      if (isRelationshipReference(entry)) {
        ids.push(entry.value);
      }
    }
  });
  let next = pkg;
  const images: { id: string; media: string }[] = [];
  for (const id of ids) {
    const record = (next.relationships.get(numbering) ?? []).find((entry) => entry.id === id);
    const resolved = record === undefined ? null : resolveRelationship(record);
    const target = resolved?.mode === 'Internal' && resolved.target.ok ? resolved.target : null;
    const listed = target === null ? undefined : media.get(partNameKey(target.partName));
    const without = listed && withoutRelationship(next, numbering, id);
    if (record?.type !== RELATIONSHIP_TYPES.image || !listed || !without) {
      return null;
    }

    next = without;
    images.push({ id, media: listed.key });
  }

  return { pkg: next, element: bullet, images };
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
export const withCheckpointLists = (
  doc: OrderedDocument,
  taken: readonly TakenList[],
  listedIds: readonly string[]
): boolean =>
  withCheckpointGiven(doc, taken, listedIds, ({ key, definition, bullets }, suggestion) => ({
    kind: 'list',
    key,
    suggestion,
    numId: orderedListNumber(key, 'num'),
    abstractNumId: orderedListNumber(key, 'abstract'),
    levels: [],
    definition,
    ...(bullets.length > 0 ? { bullets } : {}),
  }));

/**
 * Give each taken style back to its suggestion, by the ids of the suggestions the manifest lists
 * (`listedIds`), as a pasted style is held. False when a suggestion is not listed or a style has
 * no id or no known type.
 */
export const withCheckpointStyles = (
  doc: OrderedDocument,
  taken: readonly TakenStyle[],
  listedIds: readonly string[]
): boolean =>
  withCheckpointGiven(doc, taken, listedIds, ({ key, element }, suggestion) =>
    styleElementResource(key, suggestion, element)
  );
