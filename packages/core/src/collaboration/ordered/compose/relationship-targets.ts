// The relationships a change adds: their full types, and their targets as their owners write
// them.

import type { LaterRelationships } from '../../../store/store/content-control-prompt-glossary.ts';
import {
  FRAGMENT_OBJECT_RELATIONSHIP_TYPES,
  FRAGMENT_PART_RELATIONSHIP_TYPES,
} from '../change/fragment-names-objects.ts';
import { OrderedRefusal } from '../model/refusal.ts';
import type { OrderedRelationshipResource } from '../model/resources.ts';
import type { OrderedDocument } from '../model/types.ts';

export const OFFICE_RELATIONSHIPS =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

/** The full type of each kind of relationship a change adds. */
export const RELATIONSHIP_TYPES: Readonly<Record<OrderedRelationshipResource['type'], string>> = {
  image: `${OFFICE_RELATIONSHIPS}/image`,
  hyperlink: `${OFFICE_RELATIONSHIPS}/hyperlink`,
  header: `${OFFICE_RELATIONSHIPS}/header`,
  footer: `${OFFICE_RELATIONSHIPS}/footer`,
  footnotes: `${OFFICE_RELATIONSHIPS}/footnotes`,
  endnotes: `${OFFICE_RELATIONSHIPS}/endnotes`,
  oleObject: FRAGMENT_OBJECT_RELATIONSHIP_TYPES.oleObject,
  ink: FRAGMENT_OBJECT_RELATIONSHIP_TYPES.ink,
  ...FRAGMENT_PART_RELATIONSHIP_TYPES,
};

/** A relationship target from its owner: relative within the owner's folder, else absolute. */
export function targetFrom(owner: string, part: string): string {
  const folder = owner.slice(0, owner.lastIndexOf('/') + 1);
  return part.startsWith(folder) ? part.slice(folder.length) : part;
}

/** A relationship's target as its owner writes it; refused when it names no part. */
export function targetOf(
  doc: OrderedDocument,
  resource: OrderedRelationshipResource,
  owner: string
) {
  const target = rawTargetOf(doc, resource, owner);
  if (target === null) {
    throw new OrderedRefusal('unknown-address', `no part for ${resource.key}`);
  }

  return target;
}

/** A relationship's target as its owner writes it, or null when it names no part. */
function rawTargetOf(
  doc: OrderedDocument,
  resource: OrderedRelationshipResource,
  owner: string
): string | null {
  if ('external' in resource.target) {
    return resource.target.external;
  }

  const target = doc.resources.added.get(resource.target.resource);
  if (!target || (target.kind !== 'media' && target.kind !== 'story' && target.kind !== 'part')) {
    return null;
  }

  return targetFrom(owner, target.part);
}

/**
 * The relationships a change added, by owner and id, as the package will hold them once
 * `withOrderedResources` writes them: a kept prompt's ids are read before that
 * (`withRetainedPrompts`).
 */
export function laterRelationshipsOf(doc: OrderedDocument): LaterRelationships {
  return (owner, id) => {
    for (const resource of doc.resources.added.values()) {
      if (resource.kind !== 'relationship' || resource.owner !== owner || resource.id !== id) {
        continue;
      }

      const rawTarget = rawTargetOf(doc, resource, owner);
      return rawTarget === null
        ? undefined
        : {
            ownerPart: owner,
            id,
            type: RELATIONSHIP_TYPES[resource.type],
            rawTarget,
            targetMode: 'external' in resource.target ? 'External' : 'Internal',
          };
    }

    return undefined;
  };
}
