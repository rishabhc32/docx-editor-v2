// Who typed the units of a pending insertion: the namespace of the room change that typed them,
// and its person.

import { mintedChange, parseOrderedChangeId } from '../change/ids.ts';
import { inserter } from './document.ts';
import type { OrderedPerson } from './person.ts';
import { isUnit, loadedSpace, type OrderedDocument, type OrderedItem } from './types.ts';

/** Units of one space that one room change typed: the offsets from `from` up to `to`. */
export interface TypedRun {
  readonly from: number;
  readonly to: number;
  readonly typer: string;
}

/** The typers a checkpoint kept for a document loaded from it: runs of typed units by space. */
const LOADED = new WeakMap<OrderedDocument, ReadonlyMap<string, readonly TypedRun[]>>();

/** Who typed a unit: the namespace of the room change that typed it, and its person. @public */
export interface OrderedTyper {
  /** The namespace of the room change that typed the unit. */
  readonly namespace: string;
  /** The person who owns the namespace. */
  readonly person: OrderedPerson;
}

/**
 * Who typed the unit at `offset` (in the item's space) of an item a pending suggestion inserted,
 * such as an organiser's words typed directly inside a party's insertion, new row, new note or
 * new text box. A copy of the room and a copy opened from its checkpoint give the same answer.
 * Null for a unit of any other item, and for one no room change typed (a file's own).
 *
 * @public
 */
export function orderedUnitTyper(
  doc: OrderedDocument,
  item: OrderedItem,
  offset: number
): OrderedTyper | null {
  const inside = offset >= item.offset && offset < item.offset + item.length;
  if (!isUnit(item) || !inside || inserter(doc, item) === null) {
    return null;
  }

  // A checkpoint's typer, else the change's: it typed units in the space `i:<change id>.<k>`,
  // and loaded those of the engine's template (a paste, a control's value) in the space of a
  // paragraph it made, `n:<change id>.<k>`; a prompt is typed by no one
  const runs = LOADED.get(doc)?.get(item.space);
  const kept = runs?.find((run) => offset >= run.from && offset < run.to);
  const node = loadedSpace(item.space)?.node ?? '';
  const change =
    item.prompt === true ? null : (mintedChange(item.space, 'i') ?? mintedChange(node, 'n'));
  const namespace = kept?.typer ?? parseOrderedChangeId(change ?? '')?.namespace;
  const person = namespace === undefined ? undefined : doc.namespaces.get(namespace);
  if (namespace === undefined || person === undefined) {
    return null;
  }

  return { namespace, person };
}

/** The typers a checkpoint named for the units of a document loaded from it, by space. */
export function loadedTypersOf(doc: OrderedDocument): ReadonlyMap<string, readonly TypedRun[]> {
  return LOADED.get(doc) ?? new Map();
}

/** Keep the typers a checkpoint names for the units of a document loaded from it, by space. */
export function keepLoadedTypers(
  doc: OrderedDocument,
  typers: ReadonlyMap<string, readonly TypedRun[]>
): void {
  LOADED.set(doc, typers);
}
