// The grid lines a table's cells no longer use, with real editors and one ordered room, against
// the editor alone. A decision that takes away the only row over a line removes the line, live,
// after a save and reopen, and after one and two folds. A direct row deletion removes none, also
// in a table an earlier decision reached. The room once joined the grid at every drawing while
// the table held a decided label, which a save does not keep (random session `sessions-two` seed
// 5), so the live room joined after a direct edit where a fold and the editor alone do not.

import { afterEach, describe, expect, test } from 'bun:test';
import { docx, p, t } from '../word-fixtures.ts';
import { openOrderedWordFile } from '../../client/open.ts';
import { suggestionsOnlyPolicy } from '../../room/suggestions-only.ts';
import { checkpointBytes, mediaStore } from '../fold-check.ts';
import { decidedCheck, type Act } from './decided-harness.ts';
import { closeAll, join, localRoom, ORGANISER, settle, surface } from './editor-harness.ts';
import { reviewTestModule } from './review-module.ts';
import {
  alone,
  caretIn,
  caretInSurface,
  commandOn,
  comparedBody,
  converged,
  tableCommand,
  type Surface,
} from './tables-harness.ts';

afterEach(closeAll);

const tc = (text: string, props: string) => `<w:tc><w:tcPr>${props}</w:tcPr>${p(t(text))}</w:tc>`;
const w = (width: number) => `<w:tcW w:w="${width}" w:type="dxa"/>`;

/** A table of two columns: a first row over both, and a second row of two cells. */
const SPANNED = docx(
  p(t('The fees are below.')) +
    '<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr>' +
    '<w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/></w:tblGrid>' +
    `<w:tr>${tc('Fee and rate', w(6000) + '<w:gridSpan w:val="2"/>')}</w:tr>` +
    `<w:tr>${tc('100', w(3000))}${tc('5', w(3000))}</w:tr>` +
    '</w:tbl>' +
    p(t('Signed here.'))
);

/** The table as a reader of the file sees it: no ids, dates or run properties. */
const tableOf = (xml: string): string =>
  (/<w:tbl>.*<\/w:tbl>/s.exec(xml)?.[0] ?? 'no table')
    .replace(/ w:(id|date|rsidR|rsidRPr|rsidDel|rsidTr)="[^"]*"| w14:(paraId|textId)="[^"]*"/g, '')
    .replace(/ xml:space="preserve"|<w:rPr>.*?<\/w:rPr>/g, '')
    .replace(/<\/w:r><w:r>/g, '')
    .replace(/<\/w:t><w:t>/g, '');

const check = decidedCheck(tableOf);

/** Delete the row of "100", as a suggestion. */
const deleteRow: Act = (person) => {
  person.editor.setMode('suggesting');
  caretIn(person, '100');
  expect(tableCommand(person, { type: 'deleteRow' }).ok).toBe(true);
};

describe('a decision that takes away the last row over a grid line', () => {
  test("a party's row deletion, accepted", async () => {
    await check('a row deletion', SPANNED, [deleteRow], 'accept', 'ben');
  });
});

const review = reviewTestModule();

/** Ana, a party, and Olive, the organiser editing directly, in a room on `file`. */
async function seats(file: Uint8Array, number = 1) {
  const room = localRoom(file, { file: number, policy: suggestionsOnlyPolicy() });
  const ana = await join(room, 'Ana', { wordFile: file, review });
  const olive = await join(room, 'Olive', {
    wordFile: file,
    mode: 'edit',
    login: ORGANISER,
    review,
  });
  return { room, ana, olive };
}

/** Olive's direct deletion of the row of "100". */
const directDeletion = (s: Surface) => {
  caretInSurface(s, '100');
  expect(commandOn(s, { type: 'deleteRow' }).ok).toBe(true);
};

describe('a direct row deletion in a table a decision reached', () => {
  test('removes no grid line: live, after a fold, as the editor alone', async () => {
    // Ana's new row under the first, accepted by Olive: a decision reaches the table
    const live = await seats(SPANNED);
    live.ana.editor.setMode('suggesting');
    caretIn(live.ana, 'Fee and rate');
    expect(tableCommand(live.ana, { type: 'insertRow', where: 'below' }).ok).toBe(true);
    await settle();
    const command = { type: 'resolveAllReviewChanges', action: 'accept', scope: 'document' };
    expect(live.olive.editor.exec(command as never).ok).toBe(true);
    await settle();
    const decided = new Uint8Array(await live.olive.editor.save());
    const { base } = openOrderedWordFile(SPANNED, { file: 1 });
    const folded = await checkpointBytes(live.room.room.document, base, mediaStore(base));
    const expected = alone(
      { author: 'Olive', mode: 'edit', review, file: decided },
      directDeletion
    );

    directDeletion(surface(live.olive));
    await settle();
    converged(live.room, live.ana, live.olive);
    expect(comparedBody(surface(live.olive).session.part())).toBe(expected);
    closeAll();

    // The room folded after the decision, opened again, the same deletion
    const again = await seats(folded, 2);
    directDeletion(surface(again.olive));
    await settle();
    converged(again.room, again.ana, again.olive);
    expect(comparedBody(surface(again.olive).session.part())).toBe(expected);
  });
});
