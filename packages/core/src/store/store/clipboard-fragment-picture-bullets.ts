// Picture bullets on copy and paste: Word copies a list's bullets and their images with it.

import { isRelationshipReference } from '../package/canonical-media-references.ts';
import type { OoxmlPackage } from '../package/ooxml-package.ts';
import {
  WML_NAMESPACE_URI,
  type OoxmlElement,
  type OoxmlNode,
  type OoxmlPart,
} from '../package/ooxml-tree.ts';
import { resolveInternalTarget } from '../package/opc-names.ts';
import { relationshipsOf } from '../package/package-edit.ts';
import type { RelationshipRecord } from '../package/relationships.ts';
import { sha256FontBytes } from '../package/sha256.ts';
import { relationshipIdsIn } from './clipboard-fragment-closure.ts';
import { isElementNode, isWml, nodeSignature } from './clipboard-fragment-defaults.ts';
import { withRewrittenAttribute } from './clipboard-fragment-identifiers.ts';
import { attributeValueOf } from './tree-op-nodes.ts';

/** The picture bullets of a part compared by look: the number the layout reads. */
const MAX_PICTURE_BULLETS = 64;

/** The picture bullet ids list levels name (`w:lvlPicBulletId/@w:val`), under nodes. */
export function pictureBulletIdsIn(nodes: readonly OoxmlNode[]): Set<string> {
  const ids = new Set<string>();
  const stack = [...nodes];
  for (let node = stack.pop(); node !== undefined; node = stack.pop()) {
    if (!isElementNode(node)) continue;
    if (isWml(node, 'lvlPicBulletId')) {
      const id = attributeValueOf(node, 'val');
      if (id !== undefined) ids.add(id);
    }

    for (const child of node.children) {
      stack.push(child);
    }
  }

  return ids;
}

/**
 * The `w:numPicBullet` elements of a numbering part, by id: the first of each id, as layout
 * reads.
 */
function bulletsOf(part: OoxmlPart | null): Map<string, OoxmlElement> {
  const out = new Map<string, OoxmlElement>();
  for (const child of part?.root.children ?? []) {
    if (!isElementNode(child) || !isWml(child, 'numPicBullet')) continue;
    const id = attributeValueOf(child, 'numPicBulletId');
    if (id !== undefined && !out.has(id)) out.set(id, child);
  }

  return out;
}

/** Levels' picture bullet ids rewritten through `map`; an id mapped to null loses its bullet. */
export function withPictureBulletIds<T extends OoxmlNode>(
  node: T,
  map: ReadonlyMap<string, string | null>
): T {
  if (!isElementNode(node)) return node;
  let changed = false;
  const children: OoxmlNode[] = [];
  for (const child of node.children) {
    if (isElementNode(child) && isWml(child, 'lvlPicBulletId')) {
      const id = attributeValueOf(child, 'val');
      const mapped = id === undefined ? undefined : map.get(id);
      if (mapped === null) {
        changed = true;
        continue;
      }

      if (mapped !== undefined && mapped !== id) {
        changed = true;
        children.push(withRewrittenAttribute(child, WML_NAMESPACE_URI, 'val', mapped));
        continue;
      }
    }

    const next = withPictureBulletIds(child, map);
    if (next !== child) changed = true;
    children.push(next);
  }

  return changed ? ({ ...node, children } as T) : node;
}

/**
 * What a picture bullet looks like, whatever its id: its shape, with each image relationship
 * replaced by the image's SHA-256. Two bullets of one look are one bullet, in any document.
 * Only the first `MAX_PICTURE_BULLETS` of a part are read, the number layout reads.
 */
export function pictureBulletIdentities(
  pkg: OoxmlPackage,
  part: OoxmlPart | null
): Map<string, string> {
  const out = new Map<string, string>();
  if (!part) return out;
  const images = new Map<string, string>();
  for (const record of relationshipsOf(pkg, part.name)) {
    if (record.targetMode === 'External') continue;
    const resolved = resolveInternalTarget(record.ownerPart, record.rawTarget);
    const bytes = resolved.ok ? pkg.partBytes.get(resolved.partName) : undefined;
    if (bytes) images.set(record.id, sha256FontBytes(bytes));
  }

  const hashed = (node: OoxmlNode): OoxmlNode => {
    if (!isElementNode(node)) return node;
    return {
      ...node,
      attributes: node.attributes
        .filter((a) => !(a.namespaceUri === WML_NAMESPACE_URI && a.localName === 'numPicBulletId'))
        .map((a) =>
          isRelationshipReference(a) ? { ...a, value: `sha256:${images.get(a.value) ?? ''}` } : a
        ),
      children: node.children.map(hashed),
    } as OoxmlElement;
  };
  for (const [id, bullet] of [...bulletsOf(part)].slice(0, MAX_PICTURE_BULLETS)) {
    out.set(id, nodeSignature(hashed(bullet)));
  }

  return out;
}

/** The bullet identities of a paste's two numbering parts. */
export interface PictureBulletContext {
  readonly fragment: ReadonlyMap<string, string>;
  readonly target: ReadonlyMap<string, string>;
}

/**
 * The bullet identities of a paste, or none when no level of the fragment names a bullet. A
 * level that names a bullet the fragment lacks (its image did not travel) looks like no bullet,
 * never like the target's bullet of the same id.
 */
export function pictureBulletContext(
  fragment: OoxmlPackage,
  fragmentNumbering: OoxmlPart | null,
  target: OoxmlPackage,
  targetNumbering: OoxmlPart | null
): PictureBulletContext | undefined {
  if (!fragmentNumbering || pictureBulletIdsIn([fragmentNumbering.root]).size === 0) {
    return undefined;
  }

  return {
    fragment: pictureBulletIdentities(fragment, fragmentNumbering),
    target: pictureBulletIdentities(target, targetNumbering),
  };
}

/** A definition as its signature sees it: each picture bullet by its look, not its id. */
export function bySignatureLook(
  node: OoxmlElement,
  identities: ReadonlyMap<string, string> | undefined
): OoxmlElement {
  if (identities === undefined) return node;
  const map = new Map<string, string>();
  for (const id of pictureBulletIdsIn([node])) {
    map.set(id, `pic:${identities.get(id) ?? ''}`);
  }

  return withPictureBulletIds(node, map);
}

export interface PictureBulletImportPlan {
  /** Fragment bullet id to the target's: an equal bullet it holds, a fresh one, or null (none). */
  readonly bulletIdMap: Map<string, string | null>;
  /** Renumbered `w:numPicBullet`s to insert before the first definition. */
  readonly bulletsToImport: OoxmlElement[];
}

/**
 * The picture bullets that imported definitions name: one the target holds with the same look
 * is reused, any other imports under a fresh id. `null` when no fresh id is left.
 */
export function planPictureBulletImport(
  fragmentNumbering: OoxmlPart | null,
  imported: readonly OoxmlElement[],
  context: PictureBulletContext | undefined,
  freshId: () => string | null
): PictureBulletImportPlan | null {
  const plan: PictureBulletImportPlan = { bulletIdMap: new Map(), bulletsToImport: [] };
  const named = pictureBulletIdsIn(imported);
  if (named.size === 0) return plan;
  const bullets = bulletsOf(fragmentNumbering);
  const targetByLook = new Map<string, string>();
  for (const [id, look] of context?.target ?? []) {
    if (!targetByLook.has(look)) targetByLook.set(look, id);
  }

  for (const id of named) {
    const bullet = bullets.get(id);
    const look = context?.fragment.get(id);
    if (bullet === undefined || look === undefined) {
      plan.bulletIdMap.set(id, null);
      continue;
    }

    const reused = targetByLook.get(look);
    if (reused !== undefined) {
      plan.bulletIdMap.set(id, reused);
      continue;
    }

    const fresh = freshId();
    if (fresh === null) return null;
    plan.bulletIdMap.set(id, fresh);
    targetByLook.set(look, fresh);
    plan.bulletsToImport.push(
      withRewrittenAttribute(bullet, WML_NAMESPACE_URI, 'numPicBulletId', fresh)
    );
  }

  return plan;
}

/**
 * The picture bullets a copy carries for its list definitions, with the numbering part's image
 * relationships they name. A level names its `w:numPicBullet` (ECMA-376 Part 1, §17.9.20) by its
 * `w:lvlPicBulletId` (§17.9.12), and the bullet's image is a relationship of the NUMBERING part.
 * `resolveMedia` copies the images and answers the relationships that cannot travel; a bullet that
 * names one drops, and the level that named it keeps its `w:lvlText`.
 */
export function copiedPictureBullets(
  pkg: OoxmlPackage,
  numbering: OoxmlPart | null,
  definitions: readonly OoxmlElement[],
  resolveMedia: (records: readonly RelationshipRecord[]) => ReadonlySet<string>
): { readonly bullets: OoxmlElement[]; readonly records: RelationshipRecord[] } {
  const named = pictureBulletIdsIn(definitions);
  if (!numbering || named.size === 0) return { bullets: [], records: [] };
  const bullets = [...bulletsOf(numbering)].filter(([id]) => named.has(id)).map(([, b]) => b);
  const used = relationshipIdsIn(bullets);
  const records = relationshipsOf(pkg, numbering.name).filter((record) => used.has(record.id));
  const dropped = resolveMedia(records);
  const kept = bullets.filter((bullet) => {
    for (const id of relationshipIdsIn([bullet])) {
      if (dropped.has(id)) return false;
    }

    return true;
  });
  return { bullets: kept, records: records.filter((record) => !dropped.has(record.id)) };
}

/**
 * Pasted picture bullets with their image relationships rewritten to the target's. The images
 * merge as every pasted picture's does (one part per content hash, the bytes sniffed). A bullet
 * whose image did not merge drops, and its id maps to null, so the levels that named it keep
 * their `w:lvlText`.
 */
export function mergedPictureBullets(
  bullets: readonly OoxmlElement[],
  relIdMap: ReadonlyMap<string, string>,
  dropRelIds: ReadonlySet<string>
): { readonly bullets: OoxmlElement[]; readonly dropped: Map<string, null> } {
  const kept: OoxmlElement[] = [];
  const dropped = new Map<string, null>();
  for (const bullet of bullets) {
    const ids = relationshipIdsIn([bullet]);
    if ([...ids].some((id) => dropRelIds.has(id) || !relIdMap.has(id))) {
      const id = attributeValueOf(bullet, 'numPicBulletId');
      if (id !== undefined) dropped.set(id, null);
      continue;
    }

    kept.push(withRelationshipIds(bullet, relIdMap) as OoxmlElement);
  }

  return { bullets: kept, dropped };
}

function withRelationshipIds(node: OoxmlNode, map: ReadonlyMap<string, string>): OoxmlNode {
  if (!isElementNode(node)) return node;
  return {
    ...node,
    attributes: node.attributes.map((a) =>
      isRelationshipReference(a) && map.has(a.value) ? { ...a, value: map.get(a.value)! } : a
    ),
    children: node.children.map((child) => withRelationshipIds(child, map)),
  } as OoxmlElement;
}
