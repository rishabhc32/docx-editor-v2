// A change's properties as the engine's own vocabulary. Run properties and most paragraph
// properties are flat (`OoxmlProperty`); the two paragraph properties that carry their meaning
// in children are the engine's own ops, as the editor sends them: `w:tabs` is
// `setParagraphTabStops` (a `clear` tab suppresses a stop that is in force), and `w:numPr` is
// `setListNumbering` or `setListLevel`.

import type { OoxmlProperty, TabStopWrite } from '../../../store/store/tree-op-types.ts';
import { OrderedRefusal } from '../model/refusal.ts';
import type { OrderedFormatWrite, OrderedParagraphOp } from '../model/types.ts';
import { ORDERED_CHANGE_LIMITS as LIMITS, type OrderedChangeProperty } from './types.ts';

const TWIPS = /^-?\d{1,6}$/;
const NUM_ID = /^\d{1,9}$/;
const LEVEL = /^[0-8]$/;
const ALIGNMENTS: Readonly<Record<string, TabStopWrite['alignment']>> = {
  left: 'left',
  start: 'left',
  center: 'center',
  right: 'right',
  end: 'right',
  decimal: 'decimal',
  bar: 'bar',
};
const LEADERS = new Set<string>(['none', 'dot', 'hyphen', 'underscore', 'heavy', 'middleDot']);

const refuse = (detail: string): never => {
  throw new OrderedRefusal('invalid-edit', detail);
};

/** Flat properties, as the engine's ops take them. */
export function flatProperties(properties: readonly OrderedChangeProperty[]): OoxmlProperty[] {
  return properties.map((property) => {
    if (property.remove) refuse(`${property.localName} is not taken off here`);
    if (property.children !== undefined) refuse(`${property.localName} takes no children here`);
    return property.attributes === undefined
      ? { localName: property.localName }
      : { localName: property.localName, attributes: property.attributes };
  });
}

/** A format's press: the properties it sets, and the names of those it takes off. */
export function formatWrite(properties: readonly OrderedChangeProperty[]): OrderedFormatWrite {
  const remove = properties.filter((p) => p.remove).map((p) => p.localName);
  const set = flatProperties(properties.filter((p) => !p.remove));
  return remove.length === 0 ? { properties: set } : { properties: set, remove };
}

function twips(value: string | undefined, what: string): number {
  if (value === undefined || !TWIPS.test(value)) return refuse(`${what} in twips`);
  const n = Number(value);
  if (Math.abs(n) > LIMITS.columnWidth) return refuse(`${what} past 22 inches`);
  return n;
}

function tabStops(property: OrderedChangeProperty): OrderedParagraphOp {
  if (property.attributes !== undefined) refuse('w:tabs takes no attributes');
  const stops: TabStopWrite[] = [];
  const inForce: number[] = [];
  for (const tab of property.children ?? []) {
    const val = tab.attributes?.val;
    const position = twips(tab.attributes?.pos, 'a tab position');
    if (val === 'clear') {
      inForce.push(position);
      continue;
    }

    const alignment = val === undefined ? undefined : ALIGNMENTS[val];
    if (alignment === undefined) return refuse('a tab of a kind the engine writes');
    const leader = tab.attributes?.leader;
    if (leader !== undefined && !LEADERS.has(leader)) return refuse('a tab leader');
    stops.push({
      positionTwips: position,
      alignment,
      ...(leader === undefined ? {} : { leader: leader as NonNullable<TabStopWrite['leader']> }),
    });
  }

  return {
    op: 'setParagraphTabStops',
    stops,
    ...(inForce.length > 0 ? { inForcePositionsTwips: inForce } : {}),
  };
}

function numbering(property: OrderedChangeProperty): OrderedParagraphOp {
  if (property.attributes !== undefined) refuse('w:numPr takes no attributes');
  let numId: string | undefined;
  let level: number | undefined;
  for (const child of property.children ?? []) {
    const val = child.attributes?.val;
    if (child.localName === 'numId') {
      if (val === undefined || !NUM_ID.test(val)) refuse('a list number');
      numId = val;
    } else {
      if (val === undefined || !LEVEL.test(val)) refuse('a list level from 0 to 8');
      level = Number(val);
    }
  }

  // Word's list number 0 is no list
  if (numId === undefined && level !== undefined) return { op: 'setListLevel', level };
  if (numId === undefined || numId === '0') return { op: 'setListNumbering', numId: null };
  return { op: 'setListNumbering', numId, level: level ?? 0 };
}

/**
 * Paragraph properties as the paragraph presses that write them: the flat ones and every one
 * taken off as one `setParagraphProperties`, then the tab stops and the list, each its own press.
 */
export function paragraphOps(properties: readonly OrderedChangeProperty[]): OrderedParagraphOp[] {
  const flat: OrderedChangeProperty[] = [];
  const ops: OrderedParagraphOp[] = [];
  for (const property of properties) {
    // A property taken off is the flat press's too: the engine's op leaves it out
    if (property.remove) flat.push(property);
    else if (property.localName === 'tabs') ops.push(tabStops(property));
    else if (property.localName === 'numPr') ops.push(numbering(property));
    else flat.push(property);
  }

  return flat.length > 0 ? [{ op: 'setParagraphProperties', ...formatWrite(flat) }, ...ops] : ops;
}
