// A decision costs what its cards hold, not what the document holds: the engine check of a card
// draws only the blocks its plan reads (`compose/region.ts`). A room of 400 paragraphs with
// tables, a move and many cards: one decision, and the decided cards report, each cost much less
// than one compose of the whole body. A check that composed the whole part for every decision,
// as the room once did (70 ms a decision in a large saved room), costs more than one
// compose and fails this. Both times are taken in the same run, so a busy machine slows both.
// Measured: a decision about a fifth of a compose; with a whole-part check about twice one.

import { describe, expect, test } from 'bun:test';
import { orderedCardSeen, orderedDecidedCards } from '../change/apply-decide.ts';
import { applyOrderedChange } from '../change/apply.ts';
import { orderedIdMinter } from '../change/ids.ts';
import type { OrderedDecideChange } from '../change/types.ts';
import { openOrderedWordFile } from '../client/open.ts';
import { composeOrderedPart } from '../compose/compose.ts';
import { orderedCards } from '../decide.ts';
import { takeBack, withUndoRecord } from '../model/undo-record.ts';
import type { OrderedDocument } from '../model/types.ts';
import { addressOf, apply, at, BODY, DATE, EDIT, opens } from './change-apply-fixture.ts';
import { testPerson } from './test-persons.ts';
import { docx, p, t } from './word-fixtures.ts';

const PARAGRAPHS = 400;

const cell = (text: string) => `<w:tc><w:p>${t(text)}</w:p></w:tc>`;
const table = (k: number) =>
  '<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/></w:tblGrid>' +
  `<w:tr>${cell(`Item ${k}`)}${cell('Fee')}</w:tr><w:tr>${cell('Due')}${cell('30 days')}</w:tr>` +
  '</w:tbl>';

/** The shown paragraph of clause `k`: each table before it shows four cell paragraphs. */
const clause = (k: number) => k + 4 * Math.floor((k + 19) / 40);

/** The room: Ana's word in every fourth paragraph, Ben's bold, Ana's move and a new row. */
function room(): OrderedDocument {
  const body = Array.from(
    { length: PARAGRAPHS },
    (_, k) => p(t(`Clause ${k} of the agreement.`)) + (k % 40 === 20 ? table(k) : '')
  );
  const { document: doc } = openOrderedWordFile(docx(body.join('')), { file: 1 });
  for (let k = 0; k < PARAGRAPHS; k += 4) {
    apply(doc, `ana:${k + 1}`, {
      kind: 'insert',
      ...EDIT,
      place: at(doc, clause(k), 7),
      text: 'new ',
    });
  }

  for (let k = 2; k < PARAGRAPHS; k += 8) {
    const bold = [{ localName: 'b' }];
    apply(
      doc,
      `ben:${k}`,
      {
        kind: 'format',
        ...EDIT,
        from: at(doc, clause(k), 0),
        to: at(doc, clause(k), 6),
        properties: bold,
      },
      'Ben'
    );
  }

  const move = { from: at(doc, 1, 0), to: at(doc, 1, 6), place: at(doc, clause(300), 0) };
  apply(doc, 'ana:500', { kind: 'move', ...EDIT, ...move });
  const row = addressOf(opens(doc, 'tr')[0]!);
  apply(doc, 'ana:501', { kind: 'insert-row', ...EDIT, row, where: 'below' });
  return doc;
}

/** The median time of `runs` calls of `step`, in ms. */
function median(runs: number, step: (k: number) => void): number {
  const times: number[] = [];
  for (let k = 0; k < runs; k += 1) {
    const start = performance.now();
    step(k);
    times.push(performance.now() - start);
  }

  times.sort((a, b) => a - b);
  return times[Math.floor(runs / 2)]!;
}

describe('the cost of a decision', () => {
  test('one decision and its report cost much less than a compose of the whole body', () => {
    const doc = room();
    const cards = orderedCards(doc);
    expect(cards.length).toBeGreaterThan(150);
    const change = (k: number): OrderedDecideChange => {
      const card = cards[(k * 7) % cards.length]!;
      return {
        v: 1,
        id: `olive:${k + 1}`,
        kind: 'decide',
        cards: [{ suggestions: orderedCardSeen(doc, card.suggestions[0]!)! }],
        status: k % 2 === 0 ? 'accepted' : 'rejected',
      };
    };

    // Warm both paths once
    composeOrderedPart(doc, BODY);
    orderedDecidedCards(doc, change(0));
    const compose = median(9, () => composeOrderedPart(doc, BODY));
    const report = median(21, (k) => orderedDecidedCards(doc, change(k)));
    const decision = median(21, (k) => {
      const id = `olive:${k + 1}`;
      const { record } = withUndoRecord(doc, () =>
        applyOrderedChange(doc, change(k), {
          person: testPerson('Olive'),
          date: DATE,
          ids: orderedIdMinter(id),
        })
      );
      takeBack(doc, record);
    });
    expect({ report: report < compose / 2, decision: decision < compose / 2 }).toEqual({
      report: true,
      decision: true,
    });
  });
});
