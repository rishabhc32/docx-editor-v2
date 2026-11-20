// A complex field whose runs are not all in one container: a table of contents field over one
// paragraph, its result a link, as Word writes a one-row table of contents. A save and reopen, and
// a fold, keep the link and each suggestion inside the field (acceptance room s2-short: a party's
// column break in the row's link, then another party's page field after it). And a field written
// in one run, as the engine's own `insertPageField` writes it, loads and saves back wherever it
// sits.

import { describe, expect, test } from 'bun:test';
import { strFromU8, unzipSync } from 'fflate';
import { applyOrderedChange } from '../change/apply.ts';
import { orderedChangeId, orderedIdMinter } from '../change/ids.ts';
import { validateOrderedChange } from '../change/validate.ts';
import { openOrderedWordFile } from '../client/open.ts';
import { orderedViewParagraph } from '../compose/compose.ts';
import { orderedPackageBytes } from '../compose/package.ts';
import { orderedCards } from '../decide.ts';
import type { OoxmlPackage } from '../../../store/package/ooxml-package.ts';
import { isHidden } from '../model/document.ts';
import { OrderedRefusal } from '../model/refusal.ts';
import type { OrderedDocument, OrderedItem } from '../model/types.ts';
import { withUndoRecord } from '../model/undo-record.ts';
import { orderedStampDate } from '../room/room.ts';
import {
  addressOf,
  apply,
  at,
  BODY,
  DIRECT,
  EDIT,
  markOf,
  textOf,
} from './change-apply-fixture.ts';
import { checkFold, folded, mediaStore } from './fold-check.ts';
import { seededRandom } from './random.ts';
import { randomChange, type ChangeBody } from './sessions/random-change.ts';
import { checkSaveAndLoad } from './sessions/session.ts';
import { testPerson } from './test-persons.ts';
import { docx, p, t } from './word-fixtures.ts';

const fldChar = (type: string) => `<w:r><w:fldChar w:fldCharType="${type}"/></w:r>`;
const instr = (code: string) =>
  `<w:r><w:instrText xml:space="preserve"> ${code} </w:instrText></w:r>`;
const TOC = 'TOC \\o "1-3" \\h';
const tocBegin = fldChar('begin') + instr(TOC) + fldChar('separate');
const link = (runs: string) => `<w:hyperlink w:anchor="_Toc1">${runs}</w:hyperlink>`;
const pageTab = '<w:r><w:ptab w:relativeTo="margin" w:alignment="right" w:leader="dot"/></w:r>';
const REV = 'w:author="Ana" w:date="2026-10-28T09:00:00Z"';
const inserted = (runs: string) => `<w:ins w:id="7" ${REV}>${runs}</w:ins>`;
const columnBreak = '<w:r><w:br w:type="column"/></w:r>';
const heading = p(
  '<w:bookmarkStart w:id="0" w:name="_Toc1"/>' + t('Clause one') + '<w:bookmarkEnd w:id="0"/>',
  '<w:pStyle w:val="Heading1"/>'
);

/** A field in one run: its begin, instruction, separator and end, as the engine writes it. */
const oneRun = (code: string) =>
  '<w:r><w:fldChar w:fldCharType="begin"/>' +
  `<w:instrText xml:space="preserve"> ${code} </w:instrText>` +
  '<w:fldChar w:fldCharType="separate"/><w:fldChar w:fldCharType="end"/></w:r>';

/** The saved body, its ids and the paragraph ids Word adds left out. */
function savedBody(doc: OrderedDocument, base: OoxmlPackage): string {
  const xml = strFromU8(unzipSync(orderedPackageBytes(doc, base))['word/document.xml']!);
  return xml
    .slice(xml.indexOf('<w:body>'), xml.indexOf('</w:body>'))
    .replace(/ w:id="[^"]*"/g, '')
    .replace(/ w14:\w+="[^"]*"/g, '');
}

/** Who made each card, in story order. */
const cardMakers = (doc: OrderedDocument): string[] =>
  orderedCards(doc).map(
    (card) => doc.suggestions.get(card.suggestions[0] ?? '')?.attribution.author ?? ''
  );

/** The body's items. */
const bodyItems = (doc: OrderedDocument): Iterable<OrderedItem> =>
  doc.stories.get(BODY)?.tree.items() ?? [];

/** The field character items among `items`, by type. */
function fieldChars(items: Iterable<OrderedItem>, type: string): OrderedItem[] {
  return [...items].filter((item) => {
    const node = item.node;
    return (
      item.kind === 'zero' &&
      node !== null &&
      node.kind !== 'textValue' &&
      node.localName === 'fldChar' &&
      node.attributes.some((a) => a.localName === 'fldCharType' && a.value === type)
    );
  });
}

/** Save, reopen and fold once and twice: each gives back the room's file and cards. */
async function checkReopen(doc: OrderedDocument, base: OoxmlPackage, label: string) {
  const body = savedBody(doc, base);
  const makers = cardMakers(doc);
  checkSaveAndLoad(doc, label, base);
  await checkFold(doc, base, mediaStore(base), label);
  const once = await folded(doc, base, doc.file + 1, mediaStore(base));
  const twice = await folded(once.document, once.base, doc.file + 2, mediaStore(base));
  for (const copy of [once, twice]) {
    expect({ label, body: savedBody(copy.document, copy.base) }).toEqual({ label, body });
    expect({ label, makers: cardMakers(copy.document) }).toEqual({ label, makers });
  }
}

describe('a table of contents field over one paragraph, its result a link', () => {
  test("the organiser's one-row result, a party's break in its link, a page field after it", async () => {
    const file = docx(p(tocBegin) + p(link(t('Old row'))) + p(fldChar('end')) + heading);
    const { document: doc, base } = openOrderedWordFile(file, { file: 1 });
    const [separate] = fieldChars(bodyItems(doc), 'separate');
    const [end] = fieldChars(bodyItems(doc), 'end');
    if (!separate || !end) {
      throw new Error('fixture');
    }

    // The result is one row; the join of its paragraph and the end's ends the field there
    apply(
      doc,
      'olga:1',
      {
        kind: 'toc-result',
        ...DIRECT,
        separate: addressOf(separate),
        end: addressOf(end),
        instruction: TOC,
        rows: [
          { level: 0, text: 'Clause one', heading: markOf(doc, 3), bookmark: '_Toc1', page: '2' },
        ],
        bookmarks: [{ heading: markOf(doc, 3), name: '_Toc1' }],
      },
      'Olga'
    );
    apply(doc, 'olga:2', { kind: 'join', ...DIRECT, mark: markOf(doc, 0) }, 'Olga');
    expect(textOf(doc, 0)).toBe('Clause one2');
    expect(fieldChars(bodyItems(doc), 'end').length).toBe(1);
    apply(doc, 'ana:1', { kind: 'insert-break', ...EDIT, place: at(doc, 0, 5), break: 'column' });
    apply(
      doc,
      'cy:1',
      { kind: 'insert-field', ...EDIT, place: { mark: markOf(doc, 0) }, field: 'PAGE' },
      'Cy'
    );
    const body = savedBody(doc, base);
    expect(body).toMatch(
      /<w:p>(?:(?!<\/w:p>).)*TOC (?:(?!<\/w:p>).)*<w:hyperlink [^>]*><w:r>(?:(?!<\/w:hyperlink>).)*<w:ins w:author="Ana"(?:(?!<\/w:p>).)*w:fldCharType="end"(?:(?!<\/w:p>).)*<w:ins w:author="Cy"/
    );
    expect(cardMakers(doc).sort()).toEqual(['Ana', 'Cy']);
    await checkReopen(doc, base, 'the room-made field');
  });

  test("a Word file's field with a suggestion in its link keeps both after a page field", async () => {
    const row = t('Clau') + inserted(columnBreak) + t('se one') + pageTab + t('2');
    const file = docx(p(tocBegin + link(row) + fldChar('end')) + heading);
    const { document: doc, base } = openOrderedWordFile(file, { file: 1 });
    expect(cardMakers(doc)).toEqual(['Ana']);
    apply(
      doc,
      'cy:1',
      { kind: 'insert-field', ...EDIT, place: { mark: markOf(doc, 0) }, field: 'PAGE' },
      'Cy'
    );
    const body = savedBody(doc, base);
    expect(body).toContain(
      `<w:hyperlink w:anchor="_Toc1"><w:r><w:t>Clau</w:t></w:r><w:ins ${REV}>${columnBreak}</w:ins>`
    );
    expect(cardMakers(doc).sort()).toEqual(['Ana', 'Cy']);
    await checkReopen(doc, base, 'the Word file');
  });
});

describe('a field in one run', () => {
  const page = oneRun('PAGE');
  const pageRef = oneRun('PAGEREF _Toc1 \\h');
  const deletedPage = page.replaceAll('instrText', 'delInstrText');
  const numPages = fldChar('begin') + instr('NUMPAGES') + fldChar('separate') + t('3');
  const quote = fldChar('begin') + instr('QUOTE "x"') + fldChar('separate');
  const tocRow = tocBegin + link(t('Row') + pageTab + t('2')) + fldChar('end');
  for (const [label, runs, field] of [
    ['between words', t('a') + page + t('b'), page],
    ['after another field', numPages + fldChar('end') + page, page],
    ['after another field in one run', oneRun('NUMPAGES') + page, page],
    ['inside another field', quote + page + fldChar('end'), page],
    ['inside a link', link(t('a') + page + t('b')), page],
    ['in a table of contents link', tocBegin + link(t('Row') + pageRef) + fldChar('end'), pageRef],
    ['after a table of contents field', tocRow + page, page],
    ['inside an insertion', inserted(page), page],
    ['inside a deletion', `<w:del w:id="8" ${REV}>${deletedPage}</w:del>`, deletedPage],
  ] as const) {
    test(`${label}: loads, saves back and reopens with a suggestion beside it`, async () => {
      const { document: doc, base } = openOrderedWordFile(docx(p(runs) + heading), { file: 1 });
      const before = savedBody(doc, base);
      expect(before).toContain(field);

      // A suggestion in the paragraph: it is drawn again from its items, with the same content
      apply(
        doc,
        'cy:1',
        { kind: 'insert', ...EDIT, place: { mark: markOf(doc, 0) }, text: 'z' },
        'Cy'
      );
      const content = before.slice(before.indexOf('<w:p>') + 5, before.indexOf('</w:p>'));
      expect(savedBody(doc, base)).toContain(content);
      await checkReopen(doc, base, label);
    });
  }
});

describe('random edits around fields after fields', () => {
  // The lasting fold test's random sessions type no field and hold no table of contents: here
  // three people edit at random a file with a one-paragraph table of contents whose link holds
  // a suggestion, and suggest page fields after fields, links and each other's fields
  const PEOPLE = ['Olga', 'Ana', 'Cy'];
  const row = t('Clau') + inserted(columnBreak) + t('se one') + pageTab + t('2');
  const numPages = fldChar('begin') + instr('NUMPAGES') + fldChar('separate') + t('3');
  const FILE = docx(
    p(tocBegin + link(row) + fldChar('end')) +
      p(t('Fee ') + numPages + fldChar('end') + t(' due.')) +
      p(t('Page ') + oneRun('PAGE') + t(' of the terms.')) +
      heading +
      p(t('Signed here.'))
  );
  for (const seed of [7, 23]) {
    test(`seed ${seed}: each fold gives back the room's file and cards`, async () => {
      const { document: doc, base } = openOrderedWordFile(FILE, { file: 1 });
      const rand = seededRandom(seed);
      const send = (who: number, n: number, body: ChangeBody) => {
        const author = PEOPLE[who] ?? 'Olga';
        const namespace = author.toLowerCase();
        const checked = validateOrderedChange(
          { v: 1, id: orderedChangeId(namespace, n), ...body },
          { namespace, parts: doc.stories }
        );
        if (!checked.ok) {
          return;
        }

        try {
          withUndoRecord(doc, () =>
            applyOrderedChange(doc, checked.value, {
              person: testPerson(author),
              date: orderedStampDate(Date.UTC(2026, 9, 29) + 1000 * who),
              ids: orderedIdMinter(checked.value.id),
            })
          );
        } catch (error) {
          if (!(error instanceof OrderedRefusal)) {
            throw error;
          }
        }
      };

      // A page field after the table of contents' field first, then at random
      for (let n = 1; n <= 120; n += 1) {
        const who = n % PEOPLE.length;
        const mode = who === 0 && rand() < 0.3 ? 'edit' : 'suggest';
        let body: ChangeBody | null;
        if (n === 1 || rand() < 0.25) {
          // At the end of a paragraph where no field is open: a field inside a field that ends
          // in a later paragraph is no unit to the engine (design D19, open)
          const marks: OrderedItem[] = [];
          let open = 0;
          for (const item of bodyItems(doc)) {
            if (isHidden(doc, item)) {
              continue;
            }

            open += fieldChars([item], 'begin').length - fieldChars([item], 'end').length;
            if (item.kind === 'mark' && open === 0) {
              marks.push(item);
            }
          }

          const mark = n === 1 ? marks[0] : marks[Math.floor(rand() * marks.length)];
          body = mark
            ? {
                kind: 'insert-field',
                part: BODY,
                mode,
                place: { mark: addressOf(mark) },
                field: 'PAGE',
              }
            : null;
        } else {
          body = randomChange(doc, { author: PEOPLE[who] ?? '', organiser: who === 0 }, rand);
        }

        if (body !== null) {
          send(who, n, body);
        }
      }

      await checkReopen(doc, base, `seed ${seed}`);
    }, 60_000);
  }
});
