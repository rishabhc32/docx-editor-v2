// A checkpoint's list levels, relationships, parts and media back to their suggestions (design
// D14), and the take-out and give-back rules every section of resources shares.

import { ensureNumberingLevel } from '../../../store/package/numbering-part.ts';
import { withPart, type OoxmlPackage } from '../../../store/package/ooxml-package.ts';
import { ooxmlTreesEqual } from '../../../store/package/ooxml-serialize.ts';
import type { OoxmlNode, OoxmlPart } from '../../../store/package/ooxml-tree.ts';
import { partNameKey } from '../../../store/package/opc-names.ts';
import {
  relsPartNameFor,
  resolveContentTypeOf,
  withoutPart,
} from '../../../store/package/package-edit.ts';
import { resolveRelationship } from '../../../store/package/relationships.ts';
import { FRAGMENT_PART_CONTENT_TYPES } from '../change/fragment-names-objects.ts';
import { RELATIONSHIP_TYPES } from '../compose/resources.ts';
import { isSafeLinkAddress } from '../edits/resources.ts';
import type { CheckpointResource } from '../model/checkpoint-lists.ts';
import {
  ORDERED_INK_TYPE,
  orderedMediaPart,
  orderedObjectPart,
  orderedRelationshipId,
  type OrderedFileMedia,
  type OrderedMediaReference,
  type OrderedObjectPartResource,
  type OrderedRelationshipResource,
  type OrderedRelationshipType,
  type OrderedResource,
} from '../model/resources.ts';
import type { OrderedDocument } from '../model/types.ts';
import { orderedBlobHash } from '../protocol/blobs.ts';
import { attributeOf, isWml, withChildren } from '../ooxml-nodes.ts';

/** A resource without its suggestion: what a checkpoint's file holds of it. */
type Unowned<R = OrderedResource> = R extends OrderedResource ? Omit<R, 'suggestion'> : never;

/** A resource taken out of a checkpoint's file, for its suggestion. */
export interface TakenResource {
  /** The suggestion, by its place in the manifest's list. */
  readonly suggestion: number;
  readonly resource: Unowned;
}

/**
 * Give each taken entry back to its suggestion, by the ids of the suggestions the manifest lists
 * (`listedIds`), as the resource `resourceOf` makes of it. False when a suggestion is not listed,
 * an entry makes no resource, or a key is given twice.
 */
export function withCheckpointGiven<T extends { readonly suggestion: number }>(
  doc: OrderedDocument,
  taken: readonly T[],
  listedIds: readonly string[],
  resourceOf: (entry: T, suggestion: string) => OrderedResource | null
): boolean {
  const { added } = doc.resources;
  for (const entry of taken) {
    const id = listedIds[entry.suggestion];
    const resource = id === undefined ? null : resourceOf(entry, id);
    if (resource === null || added.has(resource.key)) {
      return false;
    }

    added.set(resource.key, resource);
  }

  return true;
}

/** Give each taken list level, relationship, part and media back to its suggestion. */
export const withCheckpointResources = (
  doc: OrderedDocument,
  taken: readonly TakenResource[],
  listedIds: readonly string[]
): boolean =>
  withCheckpointGiven(doc, taken, listedIds, ({ resource }, suggestion) => ({
    ...resource,
    suggestion,
  }));

/** Media the manifest's `resources` lists: its suggestion, its key and its reference. */
export interface ListedMedia {
  readonly suggestion: number;
  readonly key: string;
  readonly reference: OrderedMediaReference;
}

/**
 * The media the manifest's `resources` lists, by part name key, each with the reference the
 * manifest's `media` gives its part (the name its key and content type give), or for ink, which
 * stays inside a checkpoint, the reference its bytes in the package give. Null when one has none.
 */
export function listedMedia(
  pkg: OoxmlPackage,
  media: readonly OrderedFileMedia[],
  resources: readonly CheckpointResource[]
): Map<string, ListedMedia> | null {
  const out = new Map<string, ListedMedia>();
  for (const { kind, key, suggestion } of resources) {
    if (kind !== 'media') {
      continue;
    }

    const reference =
      media.find(
        (item) => partNameKey(item.part) === partNameKey(orderedMediaPart(key, item.contentType))
      ) ?? inkReference(pkg, key);
    if (reference === undefined) {
      return null;
    }

    out.set(partNameKey(reference.part), { suggestion, key, reference });
  }

  return out;
}

/** The ink part a key names, by the bytes the package holds of it, or undefined. */
function inkReference(pkg: OoxmlPackage, key: string): OrderedMediaReference | undefined {
  const part = orderedMediaPart(key, ORDERED_INK_TYPE);
  const bytes = [...pkg.partBytes].find(([name]) => partNameKey(name) === partNameKey(part))?.[1];
  if (bytes === undefined || resolveContentTypeOf(pkg, part) !== ORDERED_INK_TYPE) {
    return undefined;
  }

  return {
    part,
    contentType: ORDERED_INK_TYPE,
    sha256: orderedBlobHash(bytes),
    size: bytes.length,
  };
}

/**
 * A package without one relationship of `owner`, to a part or to an address outside: its element
 * in the owner's relationships part, and the package's index of it. A relationships part left with
 * none goes, as the room wrote it only for them. Null when the part does not hold it.
 */
export function withoutRelationship(
  pkg: OoxmlPackage,
  owner: string,
  id: string
): OoxmlPackage | null {
  const relsName = relsPartNameFor(owner);
  const rels = pkg.parts.get(relsName);
  if (rels === undefined) {
    return null;
  }

  const children = rels.root.children.filter(
    (child) => child.kind === 'textValue' || attributeOf(child, 'Id') !== id
  );
  if (children.length === rels.root.children.length) {
    return null;
  }

  const relationships = new Map(pkg.relationships);
  relationships.set(
    owner,
    (pkg.relationships.get(owner) ?? []).filter((record) => record.id !== id)
  );
  const next: OoxmlPackage = Object.freeze({
    ...withPart(pkg, { ...rels, root: withChildren(rels.root, children) }),
    relationships,
    externalTargets: pkg.externalTargets.filter(
      (target) => target.ownerPart !== owner || target.id !== id
    ),
  });
  if (children.some((child) => child.kind !== 'textValue')) {
    return next;
  }

  relationships.delete(owner);
  const removed = withoutPart(next, relsName);
  return removed.ok ? removed.pkg : null;
}

/** Whether a relationship of any part of the package points at the part `name`. */
function namesPart(pkg: OoxmlPackage, name: string): boolean {
  const key = partNameKey(name);
  for (const records of pkg.relationships.values()) {
    for (const record of records) {
      const resolved = resolveRelationship(record);
      if (resolved.mode === 'Internal' && resolved.target.ok) {
        if (partNameKey(resolved.target.partName) === key) {
          return true;
        }
      }
    }
  }

  return false;
}

/**
 * A checkpoint's package without the list levels the manifest lists, each taken: the `w:lvl` of
 * the list's definition the room declared, which must be the level the engine declares there
 * (`ensureNumberingLevel`). Two suggestions that declared one level take it out once. Null when
 * the file does not hold one of them as the engine declares it.
 */
export function withoutCheckpointLevels(
  pkg: OoxmlPackage,
  resources: readonly CheckpointResource[]
): { readonly pkg: OoxmlPackage; readonly taken: readonly TakenResource[] } | null {
  let next = pkg;
  const taken: TakenResource[] = [];
  const done = new Set<string>();
  for (const entry of resources) {
    if (entry.kind !== 'list-level') {
      continue;
    }

    const { suggestion, key, numId, level, format } = entry;
    taken.push({ suggestion, resource: { kind: 'list-level', key, numId, level, format } });
    if (done.has(`${numId} ${level}`)) {
      continue;
    }

    // The engine declares the level as the file's definition of the list has it, in its place
    done.add(`${numId} ${level}`);
    const without = withoutLevel(next, numId, level);
    if (without === null) {
      return null;
    }

    const declared = ensureNumberingLevel(without.pkg, numId, level, format);
    const after = declared && numberingPartOf(declared)?.root.children[without.at];
    if (!after || !ooxmlTreesEqual(without.definition, after)) {
      return null;
    }

    next = without.pkg;
  }

  return { pkg: next, taken };
}

/** The place of the child `w:<name>` whose `w:<id>` is `value`, or -1. */
export const placeOf = (children: readonly OoxmlNode[], name: string, id: string, value: string) =>
  children.findIndex(
    (child) => child.kind !== 'textValue' && isWml(child, name) && attributeOf(child, id) === value
  );

/** A package's numbering part, if any. */
export const numberingPartOf = (pkg: OoxmlPackage): OoxmlPart | undefined =>
  [...pkg.parts.values()].find((candidate) => isWml(candidate.root, 'numbering'));

/**
 * The package without level `level` of the definition list `numId` names, with that definition
 * as it was and its place in the numbering part; or null.
 */
function withoutLevel(
  pkg: OoxmlPackage,
  numId: string,
  level: number
): { readonly pkg: OoxmlPackage; readonly definition: OoxmlNode; readonly at: number } | null {
  const part = numberingPartOf(pkg);
  const children = part?.root.children ?? [];
  const num = children[placeOf(children, 'num', 'numId', numId)];
  const reference =
    num === undefined || num.kind === 'textValue'
      ? undefined
      : num.children.find((child) => isWml(child, 'abstractNumId'));
  const abstractId =
    reference === undefined || reference.kind === 'textValue'
      ? undefined
      : attributeOf(reference, 'val');
  const at =
    abstractId === undefined ? -1 : placeOf(children, 'abstractNum', 'abstractNumId', abstractId);
  const abstract = children[at];
  if (part === undefined || abstract === undefined || abstract.kind === 'textValue') {
    return null;
  }

  const levels = abstract.children.filter(
    (child) => placeOf([child], 'lvl', 'ilvl', String(level)) === -1
  );
  if (levels.length === abstract.children.length) {
    return null;
  }

  const rest = [...children];
  rest[at] = withChildren(abstract, levels);
  const without = withPart(pkg, { ...part, root: withChildren(part.root, rest) });
  return { pkg: without, definition: abstract, at };
}

/** The kind of relationship a change adds of a full type, or undefined. */
function relationshipType(type: string): OrderedRelationshipType | undefined {
  const kinds = Object.keys(RELATIONSHIP_TYPES) as OrderedRelationshipType[];
  return kinds.find((kind) => RELATIONSHIP_TYPES[kind] === type);
}

/**
 * A checkpoint's package and media without the relationships, chart and diagram parts and media
 * the manifest lists, each taken: a relationship of its owner under the id its key gives, to a
 * listed part or media, or a link's safe address; a part under the name its key and type give, of
 * that type, and media `listed` gives, each named by no other relationship and holding none of
 * its own once the listed ones are out. Null when one does not fit.
 */
export function withoutCheckpointObjects(
  pkg: OoxmlPackage,
  media: readonly OrderedFileMedia[],
  resources: readonly CheckpointResource[],
  listed: ReadonlyMap<string, ListedMedia>
): {
  readonly pkg: OoxmlPackage;
  readonly media: readonly OrderedFileMedia[];
  readonly taken: readonly TakenResource[];
} | null {
  const taken: TakenResource[] = [];
  const targets = new Map<string, string>();
  for (const [part, { suggestion, key, reference }] of listed) {
    targets.set(part, key);
    const { part: name, contentType, sha256, size } = reference;
    taken.push({
      suggestion,
      resource: { kind: 'media', key, part: name, contentType, sha256, size },
    });
  }

  const parts: string[] = [];
  for (const entry of resources) {
    if (entry.kind !== 'part') {
      continue;
    }

    const name = orderedObjectPart(entry.key, entry.type);
    const part = pkg.parts.get(name);
    if (
      part === undefined ||
      resolveContentTypeOf(pkg, name) !== FRAGMENT_PART_CONTENT_TYPES[entry.type]
    ) {
      return null;
    }

    parts.push(name);
    targets.set(partNameKey(name), entry.key);
    const resource: Unowned<OrderedObjectPartResource> = {
      kind: 'part',
      key: entry.key,
      part: name,
      type: entry.type,
      root: part.root,
    };
    taken.push({ suggestion: entry.suggestion, resource });
  }

  let next = pkg;
  for (const entry of resources) {
    if (entry.kind !== 'relationship') {
      continue;
    }

    const id = orderedRelationshipId(entry.key);
    const found = relationshipOf(next, entry.owner, id, targets);
    const without = found && withoutRelationship(next, entry.owner, id);
    if (!found || !without) {
      return null;
    }

    next = without;
    const resource: Unowned<OrderedRelationshipResource> = {
      kind: 'relationship',
      key: entry.key,
      owner: entry.owner,
      id,
      ...found,
    };
    taken.push({ suggestion: entry.suggestion, resource });
  }

  // Most media is a reference the package holds no bytes for: its content type goes all the same
  for (const name of [...parts, ...[...listed.values()].map(({ reference }) => reference.part)]) {
    const owns = (next.relationships.get(name) ?? []).length > 0;
    const removed = owns || namesPart(next, name) ? null : withoutPart(next, name);
    if (!removed?.ok) {
      return null;
    }

    next = removed.pkg;
  }

  const rest = media.filter((item) => !listed.has(partNameKey(item.part)));
  return { pkg: next, media: rest, taken };
}

/**
 * The type and target of the relationship `id` of `owner`: a kind a change adds, to a part
 * `targets` names (by part name key) or a link's safe address. Undefined when it does not fit.
 */
function relationshipOf(
  pkg: OoxmlPackage,
  owner: string,
  id: string,
  targets: ReadonlyMap<string, string>
):
  | {
      readonly type: OrderedRelationshipType;
      readonly target: { readonly resource: string } | { readonly external: string };
    }
  | undefined {
  if (!pkg.parts.has(owner)) {
    return undefined;
  }

  // An address outside: only a link's, which `sanitizeHref` keeps as it is
  const external = pkg.externalTargets.find(
    (target) => target.ownerPart === owner && target.id === id
  );
  if (external !== undefined) {
    const link =
      relationshipType(external.type) === 'hyperlink' && isSafeLinkAddress(external.rawTarget);
    return link ? { type: 'hyperlink', target: { external: external.rawTarget } } : undefined;
  }

  const record = (pkg.relationships.get(owner) ?? []).find((entry) => entry.id === id);
  const type = record === undefined ? undefined : relationshipType(record.type);
  const resolved = record === undefined ? null : resolveRelationship(record);
  const part = resolved?.mode === 'Internal' && resolved.target.ok ? resolved.target : null;
  const resource = part === null ? undefined : targets.get(partNameKey(part.partName));
  return type === undefined || resource === undefined ? undefined : { type, target: { resource } };
}
