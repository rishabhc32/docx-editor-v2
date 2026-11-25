// The list definitions, styles and other resources a checkpoint's manifest gives back to their
// suggestions (design D14).

import { normalizePartName } from '../../../store/package/opc-names.ts';
import {
  FRAGMENT_PART_CONTENT_TYPES,
  type OrderedFragmentPartType,
} from '../change/fragment-names-objects.ts';
import { NUM_ID } from '../change/apply-properties.ts';
import { at, int, Invalid, oneOf, plainFields, required, type Fields } from '../change/read.ts';
import { manifestString, readNamedOnce } from './checkpoint-suggestions.ts';
import { orderedResourceToken } from './resources.ts';

/**
 * A list definition a change added, pending or accepted, which the file holds in its numbering
 * part as the room wrote it, under the numbers its key gives (`orderedListNumber`). An open takes
 * it out of the file and gives it back to the suggestion, so a decision writes it or leaves it out
 * as the room does, in the room's order. Its picture bullets come back with it, their images as
 * the media `resources` lists.
 */
export interface CheckpointList {
  /** The suggestion, by its place in the manifest's list. */
  readonly suggestion: number;
  readonly key: string;
}

/**
 * The listed list definitions, each checked: every value comes from storage. Each names a
 * suggestion of the manifest's list of `suggestions`, and a resource key once.
 *
 * @throws Invalid when one does not read.
 */
export function readLists(value: unknown, suggestions: number): CheckpointList[] {
  return readNamedOnce(
    value,
    'lists',
    (entry, path) => readList(plainFields(entry, path), path, suggestions),
    ({ key }) => key
  );
}

/** A listed resource's suggestion and key, checked. */
function readList(f: Fields, path: string, suggestions: number): CheckpointList {
  const key = manifestString(required(f, 'key'), at(path, 'key'));
  try {
    orderedResourceToken(key);
  } catch {
    throw new Invalid('invalid-value', at(path, 'key'), 'not a resource key');
  }

  return {
    suggestion: int(required(f, 'suggestion'), at(path, 'suggestion'), 0, suggestions - 1),
    key,
  };
}

/**
 * A style a change added, pending or accepted, which the file holds in its styles part as the
 * room wrote it, under its id. An open takes it out of the file and gives it back to the
 * suggestion, so a decision writes it or leaves it out as the room does.
 */
export interface CheckpointStyle extends CheckpointList {
  readonly styleId: string;
}

/**
 * The listed styles, each checked: every value comes from storage. Each names a suggestion of the
 * manifest's list of `suggestions`, and a resource key once.
 *
 * @throws Invalid when one does not read.
 */
export function readStyles(value: unknown, suggestions: number): CheckpointStyle[] {
  return readNamedOnce(
    value,
    'styles',
    (entry, path) => {
      const f = plainFields(entry, path);
      return {
        ...readList(f, path, suggestions),
        styleId: manifestString(required(f, 'styleId'), at(path, 'styleId')),
      };
    },
    ({ key }) => key
  );
}

/**
 * Another resource a change added that the file holds under the names its key gives, by kind:
 * a list level (`w:lvl` of the list `numId`), a relationship of the part `owner` (its id from the
 * key), a chart's or a diagram's part of `type` (its name from the key and the type), or media
 * the manifest's `media` names (its name from the key and the content type). An open takes it out
 * of the file and gives it back to the suggestion, so a decision writes it or leaves it out as the
 * room does.
 */
export type CheckpointResource = CheckpointList & CheckpointResourceKind;

/** What a listed resource says past its suggestion and key, by its kind. */
export type CheckpointResourceKind =
  | {
      readonly kind: 'list-level';
      readonly numId: string;
      readonly level: number;
      readonly format: 'bullet' | 'ordered';
    }
  | { readonly kind: 'relationship'; readonly owner: string }
  | { readonly kind: 'part'; readonly type: OrderedFragmentPartType }
  | { readonly kind: 'media' };

const KINDS = ['list-level', 'relationship', 'part', 'media'] as const;
const PART_TYPES = Object.keys(FRAGMENT_PART_CONTENT_TYPES) as OrderedFragmentPartType[];

/**
 * The listed resources, each checked: every value comes from storage. Each names a suggestion of
 * the manifest's list of `suggestions`, and a resource key once.
 *
 * @throws Invalid when one does not read.
 */
export function readResources(value: unknown, suggestions: number): CheckpointResource[] {
  return readNamedOnce(
    value,
    'resources',
    (entry, path) => {
      const f = plainFields(entry, path);
      return { ...readList(f, path, suggestions), ...resourceFields(f, path) };
    },
    ({ key }) => key
  );
}

/** A listed resource's kind and what it says by its kind, checked. */
function resourceFields(f: Fields, path: string): CheckpointResourceKind {
  const kind = oneOf(required(f, 'kind'), at(path, 'kind'), KINDS);
  switch (kind) {
    case 'list-level': {
      const numId = manifestString(required(f, 'numId'), at(path, 'numId'));
      if (!NUM_ID.test(numId)) {
        throw new Invalid('invalid-value', at(path, 'numId'), 'not a list number');
      }

      return {
        kind,
        numId,
        level: int(required(f, 'level'), at(path, 'level'), 0, 8),
        format: oneOf(required(f, 'format'), at(path, 'format'), ['bullet', 'ordered'] as const),
      };
    }
    case 'relationship': {
      // A part name as OPC writes it, with no step out of the package
      const owner = manifestString(required(f, 'owner'), at(path, 'owner'));
      const name = normalizePartName(owner);
      if (!name.ok || name.partName !== owner) {
        throw new Invalid('invalid-value', at(path, 'owner'), 'not a part name');
      }

      return { kind, owner };
    }
    case 'part':
      return { kind, type: oneOf(required(f, 'type'), at(path, 'type'), PART_TYPES) };
    case 'media':
      return { kind };
  }
}
