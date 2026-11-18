// A drag that changes no cell's width, with real editors and one room (task 9.7; audit R1 row 29).
//
// The cells already have the widths the drag gives, so the engine records the grid alone: a lone
// `w:tblGridChange`. Word keeps the new grid when it rejects one (Word for Mac, 2026-10-29:
// `~$word-lone-grid-reject.docx`), and so does the editor alone (`unboundTableHistories`). The
// room gave the recorded grid back. Each decision is compared with the editor alone, live and
// on the saved file reopened.

import { afterEach, describe, expect, test } from 'bun:test';
import type { OrderedLocalRoom } from '../../client/local-room.ts';
import { orderedCards } from '../../decide.ts';
import { suggestionsOnlyPolicy } from '../../room/suggestions-only.ts';
import { docx, p, t } from '../word-fixtures.ts';
import {
  closeAll,
  editorXml,
  join,
  localRoom,
  ORGANISER,
  settle,
  surface,
  type Person,
} from './editor-harness.ts';
import { reviewTestModule } from './review-module.ts';
import {
  alone,
  commandOn,
  comparedBody,
  converged,
  gridOf,
  tableIdsIn,
  type Surface,
} from './tables-harness.ts';

afterEach(closeAll);

const review = reviewTestModule();

const cell = (text: string, width: number) =>
  `<w:tc><w:tcPr><w:tcW w:w="${width}" w:type="dxa"/></w:tcPr>${p(t(text))}</w:tc>`;

/** A table whose grid (3000, 3000) differs from its cells' widths (4000, 2000). */
const FILE = docx(
  p(t('The fees are below.')) +
    '<w:tbl><w:tblPr><w:tblW w:w="6000" w:type="dxa"/><w:tblLayout w:type="fixed"/></w:tblPr>' +
    '<w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/></w:tblGrid>' +
    `<w:tr>${cell('Fee', 4000)}${cell('Rate', 2000)}</w:tr>` +
    '</w:tbl>' +
    p(t('Signed here.'))
);

async function seats(file = FILE) {
  const room = localRoom(file, { policy: suggestionsOnlyPolicy() });
  const ana = await join(room, 'Ana', { wordFile: file, review });
  const olive = await join(room, 'Olive', {
    wordFile: file,
    mode: 'edit',
    login: ORGANISER,
    review,
  });
  return { room, ana, olive };
}

type Via = 'the editor' | 'the host';

/**
 * Decide every pending suggestion: by the editor's own Reject All (or Accept All), or by the
 * host's decision of each card of the room (`requestDecision`, as a host decides).
 */
async function decideAll(
  room: OrderedLocalRoom,
  person: Person,
  action: 'accept' | 'reject',
  via: Via
): Promise<void> {
  if (via === 'the editor') {
    const command = { type: 'resolveAllReviewChanges', action, scope: 'document' };
    expect(person.editor.exec(command as never).ok).toBe(true);
  } else {
    const status = action === 'accept' ? 'accepted' : 'rejected';
    for (const card of orderedCards(room.room.document))
      expect(person.handle.session.requestDecision(card.suggestions[0] ?? '', status)).not.toBe(
        null
      );
  }
  await settle();
}

/** The editor's drag of the first column edge onto the cells' own widths. */
function dragOntoCells(s: Surface): void {
  const ids = tableIdsIn(s.session.part().root);
  const command = {
    type: 'commitTableColumnDividerResize',
    target: {
      sourceRevision: s.session.packageRevision(),
      tableId: ids.table,
      leftGridColumnId: ids.grid[0],
      rightGridColumnId: ids.grid[1],
      isHeaderRepeat: false,
    },
    leftWidthTwips: 4000,
    rightWidthTwips: 2000,
  };
  expect(commandOn(s, command).ok).toBe(true);
}

describe('a drag that changes no cell width (a lone grid record), as the editor alone', () => {
  for (const via of ['the editor', 'the host'] as const)
    for (const action of ['accept', 'reject'] as const)
      test(`a party's drag, ${action}ed by ${via}, live and on the saved file`, async () => {
        const expected = alone(
          { author: 'Ana', mode: 'suggesting', review, file: FILE, decide: action },
          dragOntoCells
        );
        const live = await seats();
        dragOntoCells(surface(live.ana));
        await settle();
        converged(live.room, live.ana, live.olive);
        expect(gridOf(live.olive)).toEqual([4000, 2000]);
        expect(editorXml(live.olive)).toContain('<w:tblGridChange ');
        expect(editorXml(live.olive)).not.toContain('<w:tcPrChange ');
        const saved = new Uint8Array(await live.ana.editor.save());
        await decideAll(live.room, live.olive, action, via);
        converged(live.room, live.ana, live.olive);
        expect(comparedBody(surface(live.olive).session.part())).toBe(expected);
        expect(gridOf(live.olive)).toEqual([4000, 2000]);
        closeAll();
        // The saved file, reopened in a room, decided the same way
        const reopened = await seats(saved);
        await decideAll(reopened.room, reopened.olive, action, via);
        converged(reopened.room, reopened.ana, reopened.olive);
        expect(comparedBody(surface(reopened.olive).session.part())).toBe(expected);
      });
});
