// Random changes of group 9.10 for the multi-person sessions: a paste with formatting (one or
// more paragraphs, a style the paste brings, a list the paste brings), a new list on a
// paragraph, and for an organiser a paragraph removal and Protect Document. Places come from
// the copy's view now, so some are refused by name, as the other kinds are.

import type { OrderedFragmentElement } from '../../change/fragment.ts';
import type { OrderedChangeAddress, OrderedChangePlace } from '../../change/types.ts';
import type { ChangeBody, Who } from './random-change.ts';

const run = (text: string, bold: boolean): OrderedFragmentElement => ({
  e: 'w:r',
  c: [
    ...(bold ? [{ e: 'w:rPr', c: [{ e: 'w:b' }] }] : []),
    { e: 'w:t', a: { 'xml:space': 'preserve' }, c: [text] },
  ],
});

const LEVEL: OrderedFragmentElement = {
  e: 'w:lvl',
  a: { 'w:ilvl': '0' },
  c: [
    { e: 'w:start', a: { 'w:val': '1' } },
    { e: 'w:numFmt', a: { 'w:val': 'lowerLetter' } },
    { e: 'w:lvlText', a: { 'w:val': '(%1)' } },
    { e: 'w:lvlJc', a: { 'w:val': 'left' } },
  ],
};

const LIST = { abstract: { e: 'w:abstractNum', c: [LEVEL] } };

const STYLE: OrderedFragmentElement = {
  e: 'w:style',
  a: { 'w:type': 'paragraph', 'w:styleId': 'PastedClause' },
  c: [
    { e: 'w:name', a: { 'w:val': 'Pasted Clause' } },
    { e: 'w:rPr', c: [{ e: 'w:i' }] },
  ],
};

/** A random change of the group, or null when the copy offers no target. */
export function randomPasteChange(
  who: Who,
  rand: () => number,
  edit: { readonly part: string; readonly mode: 'edit' | 'suggest'; readonly into?: string },
  place: () => OrderedChangePlace | undefined,
  mark: () => OrderedChangeAddress | undefined
): ChangeBody | null {
  const roll = rand();
  if (roll < 0.6) {
    const at = place();
    if (!at) return null;
    const count = 1 + Math.floor(rand() * 3);
    const listed = rand() < 0.4;
    const styled = rand() < 0.3;
    const blocks: OrderedFragmentElement[] = [];
    for (let k = 0; k < count; k += 1) {
      const pPr: OrderedFragmentElement[] = [];
      if (styled) pPr.push({ e: 'w:pStyle', a: { 'w:val': 'PastedClause' } });
      if (listed)
        pPr.push({
          e: 'w:numPr',
          c: [
            { e: 'w:ilvl', a: { 'w:val': '0' } },
            { e: 'w:numId', a: { 'w:val': '#0' } },
          ],
        });
      blocks.push({
        e: 'w:p',
        c: [
          ...(pPr.length > 0 ? [{ e: 'w:pPr', c: pPr }] : []),
          run(`Pasted ${k} `, rand() < 0.5),
          run('term', false),
        ],
      });
    }
    return {
      kind: 'insert-fragment',
      ...edit,
      place: at,
      blocks,
      ...(rand() < 0.5 ? { whole: true } : {}),
      ...(listed ? { lists: [LIST] } : {}),
      ...(styled ? { styles: [STYLE] } : {}),
    };
  }
  if (roll < 0.8) {
    const at = mark();
    return at
      ? { kind: 'new-list', ...edit, list: LIST, paragraphs: [{ mark: at, level: 0 }] }
      : null;
  }
  if (!who.organiser) return null;
  const direct = { part: edit.part, mode: 'edit' } as const;
  if (roll < 0.93) {
    const at = mark();
    return at ? { kind: 'delete-paragraph', ...direct, mark: at } : null;
  }
  return {
    kind: 'setting',
    ...direct,
    setting: { op: 'setDocumentProtection', enforce: rand() < 0.5 },
  };
}
