// The fixtures of the lock tests (`room-locked-controls.test.ts`, `room-locked-places.test.ts`):
// a content control of each lock value inline, in a block, in a table and around a table, and a
// room on a file that answers each change with its verdict.

import { applyOrderedChange } from '../change/apply.ts';
import { loadOrderedPackage } from '../index.ts';
import type { OrderedDocument, OrderedItem } from '../model/types.ts';
import { orderedBlobHash } from '../protocol/blobs.ts';
import { OrderedRoom } from '../room/room.ts';
import { suggestionsOnlyPolicy } from '../room/suggestions-only.ts';
import { BODY, FOOTNOTES, opens } from './change-apply-fixture.ts';
import { actor } from './room-harness.ts';
import { docx, MATRIX_EXTRAS, p, packageOf, RED_PNG, t } from './word-fixtures.ts';

export const LOCKS = ['contentLocked', 'sdtContentLocked', 'sdtLocked', 'unlocked'] as const;
export type Lock = (typeof LOCKS)[number];
/** The lock forbids editing the content. */
export const content = (lock: Lock) => lock === 'contentLocked' || lock === 'sdtContentLocked';
/** The lock forbids removing the control. */
export const removal = (lock: Lock) => lock === 'sdtLocked' || lock === 'sdtContentLocked';

export const lockOf = (lock: Lock) => (lock === 'unlocked' ? '' : `<w:lock w:val="${lock}"/>`);
export const sdt = (props: string, inner: string) =>
  `<w:sdt><w:sdtPr>${props}</w:sdtPr><w:sdtContent>${inner}</w:sdtContent></w:sdt>`;
export const PROMPT_FACE = '<w:rStyle w:val="PlaceholderText"/>';
export const REFERENCE =
  '<w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteReference w:id="1"/></w:r>';
// The fixture package's own link relationship (`rIdLink`)
export const LINK = `<w:hyperlink r:id="rIdLink">${t('site')}</w:hyperlink>`;
export const EXTRAS = MATRIX_EXTRAS;

export type Layout = 'inline' | 'block' | 'table' | 'tableInBlock';

/**
 * The fixtures. Inline: "Pay in [USD] now." with a rich text control holding "USD" (view
 * offsets 7 to 10), then "Moved words here.". Block: "Before.", a block control holding "Locked
 * words." and "More locked.", then "Moved words here.". Table: a 2 by 2 table whose first cell
 * holds the inline control; then "Moved words here.". Table in block: the table inside a block
 * control of that lock, its cells plain.
 */
export function fileOf(layout: Layout, lock: Lock, inner = t('USD')): Uint8Array {
  const props = `<w:alias w:val="Currency"/>${lockOf(lock)}`;
  const moved = p(t('Moved words here.'));
  const cell = (xml: string) =>
    `<w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/></w:tcPr>${xml}</w:tc>`;
  const table = (first: string) =>
    '<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr>' +
    '<w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/></w:tblGrid>' +
    `<w:tr>${cell(first)}${cell(p(t('B1')))}</w:tr><w:tr>${cell(p(t('A2')))}${cell(p(t('B2')))}</w:tr>` +
    '</w:tbl>';
  switch (layout) {
    case 'inline':
      return docx(p(t('Pay in ') + sdt(props, inner) + t(' now.')) + moved, EXTRAS);
    case 'block':
      return docx(
        p(t('Before.')) + sdt(props, p(t('Locked words.')) + p(t('More locked.'))) + moved,
        EXTRAS
      );
    case 'table':
      return docx(table(p(t('Pay in ') + sdt(props, inner) + t(' now.'))) + moved, EXTRAS);
    case 'tableInBlock':
      return docx(p(t('Before.')) + sdt(props, table(p(t('A1')))) + moved, EXTRAS);
  }
}

export const OLGA = actor('Olga', 'olga-1', 'organiser');
export const BOB = actor('Bob', 'bob-1', 'party');
export type Who = 'organiser' | 'party';

export const PNG = RED_PNG;
export const MEDIA = {
  sha256: orderedBlobHash(PNG),
  contentType: 'image/png',
  size: PNG.byteLength,
};

/** A room on a file; `send` answers `accepted` or `<code>/<reason>`. */
export function roomOn(file: Uint8Array) {
  let tick = 0;
  const room = new OrderedRoom({
    document: loadOrderedPackage(packageOf(file)),
    applyChange: (doc, change, ctx) => void applyOrderedChange(doc, change, ctx),
    policy: suggestionsOnlyPolicy(),
    now: () => Date.UTC(2026, 9, 30) + 1000 * tick++,
    blobs: { has: (hash) => hash === MEDIA.sha256 },
  });
  let k = 0;
  const send = (who: Who, change: Record<string, unknown>) => {
    const sender = who === 'organiser' ? OLGA : BOB;
    k += 1;
    const id = `${sender.namespace}:${k}`;
    const out = room.submit(sender, [{ v: 1, id, ...change } as never]);
    const refused = out.refused[0];
    return refused ? `${refused.code}${refused.reason ? `/${refused.reason}` : ''}` : 'accepted';
  };
  return { room, doc: room.document, send };
}

export const edit = (who: Who) =>
  ({ part: BODY, mode: who === 'organiser' ? 'edit' : 'suggest' }) as const;
export const control = (doc: OrderedDocument): OrderedItem => {
  const found = opens(doc, 'sdt')[0] ?? opens(doc, 'sdtBlock')[0];
  if (!found) throw new Error('no control');
  return found;
};
export const firstOf = (doc: OrderedDocument, token: string): OrderedItem => {
  const found = opens(doc, token)[0];
  if (!found) throw new Error(`no ${token}`);
  return found;
};
/** The atom of the first loaded drawing. */
export function pictureOf(doc: OrderedDocument): OrderedItem {
  for (const item of doc.stories.get(BODY)?.tree.items() ?? [])
    if (item.kind === 'atom' && item.node?.kind === 'drawing') return item;
  throw new Error('no picture');
}
/** The open item of the footnote the fixture's reference names (`w:id="1"`). */
export function noteOf(doc: OrderedDocument): OrderedItem {
  for (const item of doc.stories.get(FOOTNOTES)?.tree.items() ?? []) {
    const node = item.node;
    if (item.kind !== 'open' || item.token !== 'note' || !node || node.kind === 'textValue')
      continue;
    if (node.attributes.some((a) => a.localName === 'id' && a.value === '1')) return item;
  }
  throw new Error('no footnote');
}
export const one = (text: string) => ({
  e: 'w:p',
  c: [{ e: 'w:r', c: [{ e: 'w:t', c: [text] }] }],
});

export const BOTH: readonly Who[] = ['organiser', 'party'];
