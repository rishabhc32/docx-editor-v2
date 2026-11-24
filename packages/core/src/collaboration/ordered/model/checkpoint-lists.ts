// The list definitions and styles a checkpoint's manifest gives back to their suggestions (design
// D14).

import { at, int, Invalid, plainFields, required, type Fields } from '../change/read.ts';
import { manifestString, readNamedOnce } from './checkpoint-suggestions.ts';
import { orderedResourceToken } from './resources.ts';

/**
 * A list definition a change added, pending or accepted, which the file holds in its numbering
 * part as the room wrote it, under the numbers its key gives (`orderedListNumber`). An open takes
 * it out of the file and gives it back to the suggestion, so a decision writes it or leaves it out
 * as the room does, in the room's order. A list with picture bullets stays the file's: its images
 * are the file's media.
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
