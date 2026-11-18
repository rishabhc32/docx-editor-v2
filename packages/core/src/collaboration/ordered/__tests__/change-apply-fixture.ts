// The document the change apply tests edit: text, a link, a control, a list item, a table, a
// note reference, and a header, with the matrix package's notes, comments and numbering.

import { serializeOoxmlPart } from '../../../store/package/ooxml-tree.ts';
import { orderedIdMinter } from '../change/ids.ts';
import { applyOrderedChange, type OrderedChangeOutcome } from '../change/apply.ts';
import type { OrderedChange } from '../change/types.ts';
import { composeOrderedPart, orderedViewParagraph } from '../compose/compose.ts';
import { orderedPlaceAt, type OrderedPlace } from '../edits/places.ts';
import { loadOrderedPackage } from '../load/load.ts';
import { isHidden } from '../model/document.ts';
import { OrderedRefusal } from '../model/refusal.ts';
import type { OrderedDocument, OrderedItem } from '../model/types.ts';
import { withUndoRecord } from '../model/undo-record.ts';
import { snapshot } from './snapshot.ts';
import { docx, MATRIX_EXTRAS, p, packageOf, R, t, W } from './word-fixtures.ts';
import { testPerson } from './test-persons.ts';

export const BODY = '/word/document.xml';
export const HEADER = '/word/header1.xml';
export const FOOTNOTES = '/word/footnotes.xml';
export const DATE = '2026-10-28T09:30:00Z';

const cell = (text: string) =>
  `<w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/></w:tcPr>${p(t(text))}</w:tc>`;
const table = (rows: readonly (readonly string[])[]) =>
  '<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr>' +
  '<w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/></w:tblGrid>' +
  rows.map((row) => `<w:tr>${row.map(cell).join('')}</w:tr>`).join('') +
  '</w:tbl>';

const BODY_XML =
  p(t('The fee is 100 USD.')) +
  p(
    t('See ') +
      '<w:hyperlink r:id="rIdLink"><w:r><w:rPr><w:rStyle w:val="Hyperlink"/></w:rPr>' +
      '<w:t>the terms</w:t></w:r></w:hyperlink>' +
      t('.')
  ) +
  p(
    t('Party: ') +
      '<w:sdt><w:sdtPr><w:id w:val="5"/><w:text/></w:sdtPr><w:sdtContent>' +
      `${t('Acme')}</w:sdtContent></w:sdt>`
  ) +
  p(t('First item'), '<w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>') +
  table([
    ['A1', 'B1'],
    ['A2', 'B2'],
  ]) +
  p(t('Due in 30 days.')) +
  p(t('Signed here.'));

/** The fixture's bytes. */
export const APPLY_FILE = docx(BODY_XML, {
  ...MATRIX_EXTRAS,
  parts: {
    ...MATRIX_EXTRAS.parts,
    'word/header1.xml': `<w:hdr xmlns:w="${W}"><w:p><w:r><w:t>Header text</w:t></w:r></w:p></w:hdr>`,
  },
  relationships: `${MATRIX_EXTRAS.relationships ?? ''}<Relationship Id="rIdHeader1" Type="${R}/header" Target="header1.xml"/>`,
  contentTypes: `${MATRIX_EXTRAS.contentTypes ?? ''}<Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>`,
});

/** A fresh copy of the fixture. */
export const load = (): OrderedDocument => loadOrderedPackage(packageOf(APPLY_FILE));

/** The text a shown paragraph of a story reads now. */
export function textOf(doc: OrderedDocument, index: number, story = BODY): string {
  return (
    orderedViewParagraph(doc, story, index)
      ?.chars.map((c) => c.ch)
      .join('') ?? '<none>'
  );
}

/** The place before the n-th character of a shown paragraph (its end at its length). */
export function at(doc: OrderedDocument, index: number, n: number, story = BODY): OrderedPlace {
  const view = orderedViewParagraph(doc, story, index);
  if (!view) throw new Error(`no paragraph ${index}`);
  return orderedPlaceAt(view, n);
}

/** The address of a shown paragraph's mark. */
export function markOf(doc: OrderedDocument, index: number, story = BODY): OrderedPlace {
  const view = orderedViewParagraph(doc, story, index);
  if (!view) throw new Error(`no paragraph ${index}`);
  return { space: view.mark.space, offset: view.mark.offset };
}

/** The open items of a container token that show, in story order. */
export function opens(doc: OrderedDocument, token: string, story = BODY): OrderedItem[] {
  const out: OrderedItem[] = [];
  for (const item of doc.stories.get(story)?.tree.items() ?? [])
    if (item.kind === 'open' && item.token === token && !isHidden(doc, item)) out.push(item);
  return out;
}

export const addressOf = (item: OrderedItem): OrderedPlace => ({
  space: item.space,
  offset: item.offset,
});

type Body = OrderedChange extends infer C
  ? C extends OrderedChange
    ? Omit<C, 'v' | 'id'>
    : never
  : never;

/** Apply one change as the room does: in an undo record, with a minter for its id. */
export function apply(
  doc: OrderedDocument,
  id: string,
  body: Body,
  author = 'Ana',
  date = DATE
): OrderedChangeOutcome {
  const change = { v: 1, id, ...body } as OrderedChange;
  return withUndoRecord(doc, () =>
    applyOrderedChange(doc, change, { person: testPerson(author), date, ids: orderedIdMinter(id) })
  ).value;
}

/** The reason a change is refused, after checking it changed nothing. */
export function refusal(doc: OrderedDocument, id: string, body: Body, author = 'Ana'): string {
  const before = JSON.stringify(snapshot(doc, false));
  const comments = [...doc.comments.keys()].join();
  try {
    apply(doc, id, body, author);
  } catch (error) {
    if (!(error instanceof OrderedRefusal)) throw error;
    if (JSON.stringify(snapshot(doc, false)) !== before) return `${error.reason} (changed)`;
    if ([...doc.comments.keys()].join() !== comments) return `${error.reason} (comments)`;
    return error.reason;
  }
  return 'applied';
}

/** Every story part as XML. */
export const partsOf = (doc: OrderedDocument): string[] =>
  [...doc.stories.keys()].map((name) => serializeOoxmlPart(composeOrderedPart(doc, name).part));

export const EDIT = { part: BODY, mode: 'suggest' } as const;
export const DIRECT = { part: BODY, mode: 'edit' } as const;
