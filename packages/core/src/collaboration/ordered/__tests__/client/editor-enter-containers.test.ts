// Enter inside a link and inside a control that holds several lines (task 3.5), with real
// editors: the container splits in two, as the engine splits it. A decision keeps the two
// containers, as the engine and a reopened Word file do; taking the Enter back leaves one. A
// control that holds one line refuses Enter by name, before the editor commits, and so does an
// inline rich text control: Word ignores Enter inside it (Word 16.114, 2026-10-30, samples
// `~$i2-split.docx` and `~$i2n-split.docx`: no new paragraph, no revision, and typing at the same
// caret works).

import { afterEach, describe, expect, test } from 'bun:test';
import type { OrderedLocalRoom } from '../../client/local-room.ts';
import { docx, p, t } from '../word-fixtures.ts';
import {
  caretAt,
  closeAll,
  editorTexts,
  editorXml,
  join,
  localRoom,
  ORGANISER,
  sameAsRoom,
  screenFollowsList,
  settle,
  surface,
  type Person,
} from './editor-harness.ts';

afterEach(closeAll);

const sdt = (properties: string, text: string): string =>
  `<w:sdt><w:sdtPr>${properties}</w:sdtPr><w:sdtContent>${t(text)}</w:sdtContent></w:sdt>`;

/** A link, a rich text control, a plain text control and a multi-line one, one in each paragraph. */
const FILE = docx(
  p(t('See ') + `<w:hyperlink r:id="rIdLink">${t('the terms')}</w:hyperlink>` + t(' now.')) +
    p(t('Party: ') + sdt('<w:alias w:val="Rich"/><w:id w:val="7"/>', 'Acme Ltd')) +
    p(t('Name: ') + sdt('<w:alias w:val="Plain"/><w:id w:val="8"/><w:text/>', 'Ana Smith')) +
    p(
      t('Notes: ') +
        sdt('<w:alias w:val="Lines"/><w:id w:val="9"/><w:text w:multiLine="1"/>', 'Acme Ltd')
    )
);

const count = (xml: string, what: string): number => xml.split(what).length - 1;

function converged(room: OrderedLocalRoom, ...people: Person[]): void {
  for (const person of people) {
    expect({ person: person.name, refusals: person.refusals }).toEqual({
      person: person.name,
      refusals: [],
    });
    expect(sameAsRoom(person, room)).toBe(true);
    expect(screenFollowsList(person)).toBe(true);
  }
  const [first, ...rest] = people;
  if (first) for (const other of rest) expect(editorTexts(other)).toEqual(editorTexts(first));
}

async function seats() {
  const room = localRoom(FILE);
  const ana = await join(room, 'Ana', { wordFile: FILE });
  const org = await join(room, 'Org', { mode: 'edit', login: ORGANISER, wordFile: FILE });
  return { room, ana, org };
}

describe('Enter inside a link or a control', () => {
  test('splits a link in two, and a decision keeps the two links', async () => {
    const { room, ana, org } = await seats();
    caretAt(ana, 0, 7);
    surface(ana).splitParagraph();
    await settle();
    expect(editorTexts(org).slice(0, 2)).toEqual(['See the', ' terms now.']);
    expect(count(editorXml(org), '<w:hyperlink ')).toBe(2);
    converged(room, ana, org);
    // The organiser rejects the Enter: one paragraph, still two links (as the engine keeps them)
    const id = [...room.room.document.suggestions.keys()][0];
    if (!id) throw new Error('no suggestion');
    expect(org.handle.session.requestDecision(id, 'rejected')).not.toBeNull();
    await settle();
    expect(editorTexts(ana)[0]).toBe('See the terms now.');
    expect(count(editorXml(ana), '<w:hyperlink ')).toBe(2);
    converged(room, ana, org);
  });

  test("taking one's own Enter back leaves one link again", async () => {
    const { room, ana, org } = await seats();
    caretAt(ana, 0, 7);
    surface(ana).splitParagraph();
    await settle();
    surface(ana).undo();
    await settle();
    expect(editorTexts(org)[0]).toBe('See the terms now.');
    expect(count(editorXml(org), '<w:hyperlink ')).toBe(1);
    converged(room, ana, org);
  });

  test('splits a multi-line plain text control; the organiser can type in both halves', async () => {
    const { room, ana, org } = await seats();
    caretAt(org, 3, 11);
    surface(org).splitParagraph();
    await settle();
    expect(editorTexts(ana).slice(3, 5)).toEqual(['Notes: Acme', ' Ltd']);
    expect(count(editorXml(ana), '<w:sdt>')).toBe(4);
    caretAt(org, 4, 1);
    surface(org).insertPlainText('X');
    await settle();
    expect(editorTexts(ana)[4]).toBe(' XLtd');
    converged(room, ana, org);
    // The organiser's undo takes the typing and then the Enter back: one control again
    surface(org).undo();
    await settle();
    surface(org).undo();
    await settle();
    expect(editorTexts(ana)[3]).toBe('Notes: Acme Ltd');
    expect(count(editorXml(ana), '<w:sdt>')).toBe(3);
    converged(room, ana, org);
  });

  for (const who of ['party', 'organiser'] as const)
    test(`an inline rich text control refuses Enter by name, as Word ignores it (${who})`, async () => {
      const { room, ana, org } = await seats();
      const person = who === 'party' ? ana : org;
      const before = editorXml(person);
      caretAt(person, 1, 11);
      surface(person).splitParagraph();
      await settle();
      expect(person.refusals.map((refusal) => refusal.reason)).toEqual(['single-line-control']);
      expect(editorXml(person)).toBe(before);
      expect(room.room.order).toBe(0);
      person.refusals.length = 0;
      // Typing at the same caret works, as in Word
      caretAt(person, 1, 11);
      surface(person).insertPlainText('X');
      await settle();
      expect(editorTexts(org)[1]).toBe('Party: AcmeX Ltd');
      expect(count(editorXml(org), '<w:sdt>')).toBe(3);
      converged(room, ana, org);
    });

  test('a control that holds one line refuses Enter by name', async () => {
    const { room, ana, org } = await seats();
    const before = editorXml(ana);
    caretAt(ana, 2, 9);
    surface(ana).splitParagraph();
    await settle();
    expect(ana.refusals).toEqual([
      {
        code: 'ordered-refused',
        reason: 'single-line-control',
        detail: 'Enter in a control that holds one line',
      },
    ]);
    expect(editorXml(ana)).toBe(before);
    expect(room.room.order).toBe(0);
    ana.refusals.length = 0;
    converged(room, ana, org);
  });
});
