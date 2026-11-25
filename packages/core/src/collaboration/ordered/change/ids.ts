// Change ids, and the ids a change mints from its own id, so every copy makes the same ids.

import { fnv1a32 } from '../../../store/package/para-id.ts';
import { mintedNumber } from '../model/revision-ids.ts';

/** A namespace: letters, digits, `_` and `-`, 1 to 64 of them. @public */
export const ORDERED_NAMESPACE_PATTERN = /^[A-Za-z0-9_-]{1,64}$/;

/** `<namespace>:<counter>`: the sender's namespace and a counter the sender never reuses. */
const CHANGE_ID = /^([A-Za-z0-9_-]{1,64}):([1-9][0-9]{0,14})$/;

/** A suggestion id: the file's own (`S12`) or one a change minted (`S:ns:4.1`). */
const SUGGESTION_ID = /^[A-Za-z0-9:._-]{1,128}$/;

/** A change id's parts, or null when the id does not match the pattern. @public */
export function parseOrderedChangeId(
  id: string
): { readonly namespace: string; readonly counter: number } | null {
  const match = CHANGE_ID.exec(id);
  if (!match) return null;
  const namespace = match[1];
  const counter = Number(match[2]);
  if (namespace === undefined || !Number.isSafeInteger(counter)) return null;
  return { namespace, counter };
}

/** Whether a string is a well-formed change id. @public */
export function isOrderedChangeId(id: string): boolean {
  return parseOrderedChangeId(id) !== null;
}

/** Whether a string is a well-formed suggestion id. @public */
export function isOrderedSuggestionId(id: string): boolean {
  return SUGGESTION_ID.test(id);
}

/**
 * The change id for a namespace and a counter.
 *
 * @throws RangeError when the namespace or the counter is not valid.
 * @public
 */
export function orderedChangeId(namespace: string, counter: number): string {
  const id = `${namespace}:${counter}`;
  if (!ORDERED_NAMESPACE_PATTERN.test(namespace) || parseOrderedChangeId(id) === null)
    throw new RangeError('A change id needs a valid namespace and a counter from 1.');
  return id;
}

/** A new paragraph's identity: its mark's space and node id, and its Word ids. @public */
export interface OrderedMintedParagraph {
  /** The mark item's space (offset 0). It is also the paragraph's node id. */
  readonly space: string;
  /** `w14:paraId`: 8 hex digits, above 0 and below 0x80000000. */
  readonly paraId: string;
  /** `w14:textId`: 8 hex digits. */
  readonly textId: string;
}

/**
 * The ids one application of a change makes. Create a new minter for each application.
 *
 * @public
 */
export interface OrderedIdMinter {
  /** The change id the ids are made from. */
  readonly changeId: string;
  /** A new suggestion: `S:<change>.<k>`. */
  suggestion(): string;
  /** A space for units the change inserts (text, objects): `i:<change>.<k>`. */
  units(): string;
  /**
   * A new paragraph. Its `w14:paraId` avoids every id `taken` says is in use (upper-case hex);
   * copies whose `taken` answers the same make the same id.
   */
  paragraph(taken?: (paraId: string) => boolean): OrderedMintedParagraph;
  /** A space for a new container's open (offset 0) and close (offset 1) items: `c:<change>.<k>`. */
  container(): string;
  /** A space for a new marker with no width (a comment range): `x:<change>.<k>`. */
  marker(): string;
  /** A node id for another element the change creates: `n:<change>.<k>`. */
  node(): string;
  /** A format or property record: `f:<change>.<k>`. */
  record(): string;
  /** A move's name: `move:<change>.<k>`. */
  move(): string;
  /**
   * A resource outside the story text (media, a story part, a relationship, a list, a style, a
   * setting): `p:<change>.<k>`. Its Word names are functions of this key (design D14).
   */
  resource(): string;
  /**
   * A relationship id, `rId` and nine digits from 100000000, which no file's dense ids reach.
   * Two minted ids can meet with a chance of about one in 900 million.
   */
  relationship(): string;
  /**
   * A Word decimal id (a comment's, a note's), from 1000000000 to 2147483646, above any file's
   * dense ids. Two minted ids can meet with a chance of about one in 1.1 billion.
   */
  decimal(): number;
}

const hex8 = (value: number): string => value.toString(16).toUpperCase().padStart(8, '0');

/**
 * A `w14:paraId` from a seed: 8 hex digits above 0 and below 0x80000000, past every id `taken`
 * names. The probe is the seed's own sequence, so it ends the same on every copy.
 */
function paraIdFor(seed: string, taken: (paraId: string) => boolean): string {
  let candidate = 0;
  for (let attempt = 0; attempt < 64; attempt += 1) {
    candidate = fnv1a32(attempt === 0 ? seed : `${seed} ${attempt}`) & 0x7fffffff;
    if (candidate !== 0 && !taken(hex8(candidate))) return hex8(candidate);
  }

  // A part cannot hold 0x7fffffff paragraphs, so a free id follows
  for (;;) {
    candidate = (candidate % 0x7fffffff) + 1;
    if (!taken(hex8(candidate))) return hex8(candidate);
  }
}

/**
 * The id minter for one application of a change. Each kind of id has its own counter from 1, so
 * a paragraph keeps its id even when an earlier application (the client's own, before the room
 * ordered the change) minted another number of suggestions.
 *
 * @throws RangeError when `changeId` is not a valid change id.
 * @public
 */
export function orderedIdMinter(changeId: string): OrderedIdMinter {
  if (!isOrderedChangeId(changeId)) throw new RangeError('Not a valid change id.');
  const counters = new Map<string, number>();
  const next = (kind: string): string => {
    const k = (counters.get(kind) ?? 0) + 1;
    counters.set(kind, k);
    return `${changeId}.${k}`;
  };
  return {
    changeId,
    suggestion: () => `S:${next('S')}`,
    units: () => `i:${next('i')}`,
    paragraph: (taken = () => false) => {
      const space = `k:${next('k')}`;
      const paraId = paraIdFor(`ordered paragraph ${space}`, taken);
      const textId = hex8(fnv1a32(`ordered text ${space}`) & 0x7fffffff || 1);
      return { space, paraId, textId };
    },
    container: () => `c:${next('c')}`,
    marker: () => `x:${next('x')}`,
    node: () => `n:${next('n')}`,
    record: () => `f:${next('f')}`,
    move: () => `move:${next('move')}`,
    resource: () => `p:${next('p')}`,
    relationship: () => `rId${100_000_000 + (fnv1a32(`ordered rel ${next('r')}`) % 900_000_000)}`,
    decimal: () => mintedNumber(fnv1a32(`ordered decimal ${next('d')}`)),
  };
}

/**
 * The change a minted id of one prefix was made from, the inverse of the minter's formats
 * (`mintedChange('n:ana:4.2', 'n')` is `ana:4`), or null for any other id: a file's, or one of
 * another prefix.
 */
export function mintedChange(id: string, prefix: string): string | null {
  const dot = id.lastIndexOf('.');
  if (!id.startsWith(`${prefix}:`) || !/^[1-9][0-9]*$/.test(id.slice(dot + 1))) {
    return null;
  }

  const change = id.slice(prefix.length + 1, dot);
  return isOrderedChangeId(change) ? change : null;
}
