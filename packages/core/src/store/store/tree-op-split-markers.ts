// Which side of a paragraph split the zero-width range markers take. Split from
// `tree-op-apply.ts`, which owns the split itself.

import type { OoxmlNode } from '../package/ooxml-tree.ts';

/**
 * The range markers that CLOSE a span of content (17.13.6, 17.13.5.2, 17.13.4.x).
 *
 * A marker sitting exactly on the split has no character to sit before or after, so the
 * side it takes is a decision: an end marker stays with the head, leaving the range it
 * closes closed around the content it always covered, and everything else — a start
 * marker, a hyperlink, a picture's run — moves down with the caret. Bookmark markers are the
 * exception: every one of them stays with the head (see `zeroLengthGoesToHead`).
 */
const RANGE_END_MARKERS: ReadonlySet<string> = new Set([
  'bookmarkEnd',
  'commentRangeEnd',
  'moveFromRangeEnd',
  'moveToRangeEnd',
  'permEnd',
  'customXmlInsRangeEnd',
  'customXmlDelRangeEnd',
  'customXmlMoveFromRangeEnd',
  'customXmlMoveToRangeEnd',
]);

/** The range-START element each range-END element closes; `w:id` pairs the two. */
const RANGE_START_OF_END: ReadonlyMap<string, string> = new Map([
  ['bookmarkEnd', 'bookmarkStart'],
  ['commentRangeEnd', 'commentRangeStart'],
  ['moveFromRangeEnd', 'moveFromRangeStart'],
  ['moveToRangeEnd', 'moveToRangeStart'],
  ['permEnd', 'permStart'],
  ['customXmlInsRangeEnd', 'customXmlInsRangeStart'],
  ['customXmlDelRangeEnd', 'customXmlDelRangeStart'],
  ['customXmlMoveFromRangeEnd', 'customXmlMoveFromRangeStart'],
  ['customXmlMoveToRangeEnd', 'customXmlMoveToRangeStart'],
]);

const RANGE_START_MARKERS: ReadonlySet<string> = new Set(RANGE_START_OF_END.values());

function closesARange(node: OoxmlNode): boolean {
  return node.kind !== 'textValue' && RANGE_END_MARKERS.has(node.localName);
}

/** A marker's identity as `name\0id`, so a start and the end that closes it share a key. */
function rangeKey(localName: string, node: OoxmlNode): string | null {
  if (node.kind === 'textValue') return null;
  const id = node.attributes.find((attribute) => attribute.localName === 'id');
  return id ? `${localName}\0${id.value}` : null;
}

export function opensARange(node: OoxmlNode): string | null {
  if (node.kind === 'textValue' || !RANGE_START_MARKERS.has(node.localName)) return null;
  return rangeKey(node.localName, node);
}

/**
 * Whether a range-end marker closes a range that was already OPEN before this position.
 *
 * An end marker stays with the head so the range it closes stays closed around the content
 * it always covered — but only when its start is behind it. A range that opens AND closes at
 * the split (an empty bookmark, a comment anchored on the caret) has its start marker in the
 * tail: keeping the end behind emitted `<w:p>…<w:bookmarkEnd id="1"/></w:p><w:p><w:bookmark
 * Start id="1"/>…` — the pair inverted across two paragraphs, an end with no start before it.
 * Such an end follows its own start into the tail instead.
 */
export function closesAnOpenRange(child: OoxmlNode, openedHere: ReadonlySet<string>): boolean {
  if (!closesARange(child) || child.kind === 'textValue') return false;
  const opener = RANGE_START_OF_END.get(child.localName);
  const key = opener ? rangeKey(opener, child) : null;
  return key === null || !openedHere.has(key);
}

/**
 * Which half a paragraph child that measures ZERO characters belongs to.
 *
 * A hyperlink, a bookmark or comment marker, a run holding only a picture: none of them
 * contributes a text offset, but each sits at a definite POSITION between the runs that do.
 * Sending them all to the head — the rule this replaces — moved every hyperlink in a
 * sentence backwards past the caret, left comment ranges as empty marker pairs around the
 * wrong half, and carried an inline picture into the paragraph below when the user pressed
 * Enter at the end of the line.
 */
export function zeroLengthGoesToHead(
  child: OoxmlNode,
  position: number,
  offset: number,
  openedHere: ReadonlySet<string>
): boolean {
  if (position !== offset) return position < offset;
  // Bookmark markers at the split stay with the head, all of them: a bookmark's start keeps its
  // paragraph, and a point bookmark stays before the new paragraph mark.
  return (
    child.kind === 'bookmarkStart' ||
    child.kind === 'bookmarkEnd' ||
    closesAnOpenRange(child, openedHere)
  );
}

/** What a split reads of an atomic unit: where it ends, and the nodes that leave with it. */
interface AtomUnit {
  readonly end: number;
  readonly removeNodeIds?: readonly string[];
}

/**
 * For every node an atomic unit owns, where that unit ends. A complex field counts ONE unit, at
 * its begin character; its instruction, separator, cached result and end measure nothing.
 * Placed by position, a split right after the field sent those to the new paragraph and left
 * the begin character behind: one field broken across two paragraphs, its instruction read as
 * text. They go where their unit goes instead.
 */
export function atomUnitEnds(segments: readonly AtomUnit[]): ReadonlyMap<string, number> {
  const ends = new Map<string, number>();
  for (const segment of segments) {
    for (const id of segment.removeNodeIds ?? []) ends.set(id, segment.end);
  }
  return ends;
}

/** Where the atomic unit that owns a node of this zero-length child ends, or undefined. */
export function atomUnitEndIn(
  child: OoxmlNode,
  ends: ReadonlyMap<string, number>
): number | undefined {
  if (ends.size === 0) return undefined;
  const pending: OoxmlNode[] = [child];
  for (let node = pending.pop(); node !== undefined; node = pending.pop()) {
    const end = ends.get(node.id);
    if (end !== undefined) return end;
    if (node.kind !== 'textValue') pending.push(...node.children);
  }
  return undefined;
}
