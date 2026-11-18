// Word's Smart cut and paste ("Adjust sentence and word spacing automatically", on by default):
// a paste between two words gets a space on each side, so the pasted words do not run into the
// text around them (Word 16.114 for Mac, paste evidence 2, samples `~$paste-out-e2-*.docx`).

import type { OoxmlElement, OoxmlNode, OoxmlParagraphNode } from '../package/ooxml-tree.ts';
import { WML_NAMESPACE_URI } from '../package/ooxml-tree.ts';
import { spaceProperties } from './fragment-landing-inside.ts';
import { textRun } from './tree-op-content-controls.ts';
import { paragraphPropertiesNodeOf } from './tree-op-nodes.ts';
import { segmentsOf } from './tree-op-segments.ts';

const WORD_CHARACTER = /^[\p{L}\p{N}]$/u;
const UNSPACED_SCRIPT =
  /^[\p{Script=Han}\p{Script=Hiragana}\p{Script=Katakana}\p{Script=Thai}\p{Script=Lao}\p{Script=Khmer}\p{Script=Myanmar}]$/u;

const THAI = /^\p{Script=Thai}$/u;

/** A character Word's spacing puts a space beside: a letter or a digit of a spaced script. */
export function isSpacedWordCharacter(char: string | undefined): boolean {
  return char !== undefined && WORD_CHARACTER.test(char) && !UNSPACED_SCRIPT.test(char);
}

const isThai = (char: string | undefined): boolean => char !== undefined && THAI.test(char);

/**
 * Whether Word puts a space between the text before a paste and the paste's first character
 * (evidence W4, `~$paste-out-w4-*.docx`). A word of a script written with spaces (Latin and
 * others) gets one after a word, after a sentence's period ("Stop. fee Go") and after Thai
 * ("สวัสดี fee ครับ"). A Thai paste gets one after Thai ("สวัสดี ราคาครับ"). Chinese, Japanese, the other
 * scripts written without spaces (Lao, Khmer, Myanmar) and other punctuation get none.
 */
function spaceBefore(before: string | undefined, first: string | undefined): boolean {
  if (isSpacedWordCharacter(first)) {
    return isSpacedWordCharacter(before) || before === '.' || isThai(before);
  }

  return isThai(first) && isThai(before);
}

/**
 * Whether Word puts a space between the paste's last character and the text after it. A word of
 * a spaced script gets one before a word or Thai, but none before a comma ("Go fee, then"). A
 * paste that ends in Thai, or in any script written without spaces, gets none.
 */
function spaceAfter(last: string | undefined, after: string | undefined): boolean {
  return isSpacedWordCharacter(last) && (isSpacedWordCharacter(after) || isThai(after));
}

/** The first character (code point) of a text, or undefined. */
function firstCharacter(text: string): string | undefined {
  const point = text.codePointAt(0);
  return point === undefined ? undefined : String.fromCodePoint(point);
}

/** The last character (code point) of a text, or undefined. */
function lastCharacter(text: string): string | undefined {
  if (text.length === 0) return undefined;
  const low = text.charCodeAt(text.length - 1);
  const pair = low >= 0xdc00 && low <= 0xdfff && text.length > 1;
  return pair ? text.slice(-2) : text.slice(-1);
}

/** A visible character beside a place, with the run that holds it. */
export interface SpacingNeighbour {
  /** The character, or `''` for an item that is no text (a tab, a picture, a field). */
  readonly char: string;
  readonly run: OoxmlNode | null;
}

const isRemoved = (node: OoxmlNode): boolean =>
  node.kind !== 'textValue' &&
  node.namespaceUri === WML_NAMESPACE_URI &&
  (node.localName === 'del' || node.localName === 'moveFrom');

/** The runs of a paragraph by id, and the ids of those a tracked deletion or move holds. */
function runsOf(paragraph: OoxmlParagraphNode): {
  readonly runs: ReadonlyMap<string, OoxmlNode>;
  readonly removed: ReadonlySet<string>;
} {
  const runs = new Map<string, OoxmlNode>();
  const removed = new Set<string>();
  const walk = (node: OoxmlNode, inRemoved: boolean, depth: number): void => {
    if (node.kind === 'textValue' || depth > 64) return;
    if (node.kind === 'run') {
      runs.set(node.id, node);
      if (inRemoved) removed.add(node.id);
      return;
    }

    for (const child of node.children) {
      walk(child, inRemoved || isRemoved(node), depth + 1);
    }
  };
  for (const child of paragraph.children) {
    walk(child, false, 1);
  }

  return { runs, removed };
}

/**
 * The visible characters on each side of `offset` in a paragraph, as a reader sees them: text a
 * tracked deletion or move holds is passed over. Null where the paragraph has none.
 */
export function spacingNeighbours(
  paragraph: OoxmlParagraphNode,
  offset: number
): { readonly before: SpacingNeighbour | null; readonly after: SpacingNeighbour | null } {
  const segments = segmentsOf(paragraph);
  const { runs, removed } = runsOf(paragraph);
  const neighbour = (
    segment: (typeof segments)[number],
    text: (value: string) => string | undefined
  ): SpacingNeighbour => ({
    char: segment.node.kind === 'textValue' ? (text(segment.node.value) ?? '') : '',
    run: runs.get(segment.runId) ?? null,
  });
  let before: SpacingNeighbour | null = null;
  for (let k = segments.length - 1; k >= 0 && before === null; k -= 1) {
    const segment = segments[k]!;
    if (segment.start >= offset || segment.end <= segment.start || removed.has(segment.runId)) {
      continue;
    }

    const cut = Math.min(segment.end, offset) - segment.start;
    before = neighbour(segment, (value) => lastCharacter(value.slice(0, cut)));
  }

  let after: SpacingNeighbour | null = null;
  for (let k = 0; k < segments.length && after === null; k += 1) {
    const segment = segments[k]!;
    if (segment.end <= offset || segment.end <= segment.start || removed.has(segment.runId)) {
      continue;
    }

    const from = Math.max(segment.start, offset) - segment.start;
    after = neighbour(segment, (value) => firstCharacter(value.slice(from)));
  }

  return { before, after };
}

/**
 * Plain text a paste puts between `before` and `after`, with Word's spaces (`spaceBefore`,
 * `spaceAfter`): mostly one before it when it starts with a word and follows one, one after it
 * when it ends with a word and a word follows. The spaces take the pasted text's format.
 */
export function smartSpacedText(
  text: string,
  before: SpacingNeighbour | null,
  after: SpacingNeighbour | null
): string {
  if (text.length === 0) return text;
  const lead = spaceBefore(before?.char, firstCharacter(text));
  const trail = spaceAfter(lastCharacter(text), after?.char);
  return `${lead ? ' ' : ''}${text}${trail ? ' ' : ''}`;
}

/**
 * Pasted blocks with Word's spaces, for a paste between `before` and `after`: a run with a
 * space before the first paragraph's text in the format of its first run, and, when the paste
 * ends without its last paragraph mark, a run with a space after the last paragraph's text in
 * the format of the text after the place. Each space is part of the paste, so a tracked paste
 * holds both in its one insertion. Blocks that need no space come back as they are.
 */
export function smartSpacedBlocks(
  blocks: readonly OoxmlNode[],
  lastMarkCovered: boolean,
  before: SpacingNeighbour | null,
  after: SpacingNeighbour | null,
  nextId: () => string
): readonly OoxmlNode[] {
  const first = blocks[0];
  const last = blocks[blocks.length - 1];
  if (first === undefined || last === undefined) return blocks;
  const head = first.kind === 'paragraph' ? spacingNeighbours(first, 0).after : null;
  const tail =
    last.kind === 'paragraph' && !lastMarkCovered
      ? spacingNeighbours(last, Number.MAX_SAFE_INTEGER).before
      : null;
  const lead = head !== null && spaceBefore(before?.char, head.char);
  const trail = tail !== null && spaceAfter(tail.char, after?.char);
  if (!lead && !trail) return blocks;
  const out = [...blocks];
  if (lead) {
    const paragraph = out[0] as OoxmlElement;
    const pPr = paragraphPropertiesNodeOf(paragraph);
    const space = textRun(nextId, ' ', spaceProperties(head.run, nextId));
    const rest = paragraph.children.filter((child) => child !== pPr);
    out[0] = {
      ...paragraph,
      children: pPr ? [pPr, space, ...rest] : [space, ...rest],
    } as OoxmlNode;
  }

  if (trail) {
    const paragraph = out[out.length - 1] as OoxmlElement;
    const space = textRun(nextId, ' ', spaceProperties(after?.run ?? null, nextId));
    out[out.length - 1] = { ...paragraph, children: [...paragraph.children, space] } as OoxmlNode;
  }

  return out;
}
