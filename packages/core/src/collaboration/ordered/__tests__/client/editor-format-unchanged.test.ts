// A format that the words already have, through real editors and one room: the
// organiser's direct format and a party's suggested one each leave every copy, and the saved
// file, exactly as they were. Word makes no change for such a press. The engine used to split
// the run and re-emit its `w:rPr` in schema order, so the room wrote no record while the saved
// package changed.

import { afterEach, describe, expect, test } from 'bun:test';
import { readOoxmlPackage } from '../../../../store/package/ooxml-package.ts';
import { serializeOoxmlPart } from '../../../../store/package/ooxml-tree.ts';
import { docx, p, t } from '../word-fixtures.ts';
import {
  BODY,
  caretAt,
  closeAll,
  editorXml,
  join,
  listXml,
  localRoom,
  ORGANISER,
  sameAsRoom,
  settle,
  surface,
  type Person,
} from './editor-harness.ts';
import type { OrderedLocalRoom } from '../../client/local-room.ts';

afterEach(closeAll);

// `w:b` precedes `w:sz` in `CT_RPr`, and `w:pStyle` precedes `w:jc` in `CT_PPr`: both are
// authored out of order, so a rewrite of either container shows in the XML.
const FILE = docx(
  p(
    t('The fee is 100 USD.', '<w:sz w:val="28"/><w:b/>'),
    '<w:jc w:val="center"/><w:pStyle w:val="Quote"/><w:rPr><w:sz w:val="28"/><w:b/></w:rPr>'
  ) + p(t('Due in 30 days.'))
);

async function people() {
  const room = localRoom(FILE);
  const ana = await join(room, 'Ana', { wordFile: FILE });
  const olive = await join(room, 'Olive', { wordFile: FILE, mode: 'edit', login: ORGANISER });
  return { room, ana, olive };
}

async function savedBody(person: Person): Promise<string> {
  const read = readOoxmlPackage(new Uint8Array(await person.editor.save()), {});
  if (!read.ok) throw new Error(read.reason);
  return serializeOoxmlPart(read.package.parts.get(BODY)!);
}

/** Each copy is the room's, nothing is pending or refused, and nothing changed. */
async function unchanged(
  room: OrderedLocalRoom,
  everyone: readonly Person[],
  before: { order: number; xml: string; saved: string }
): Promise<void> {
  for (const person of everyone) {
    expect(person.refusals).toEqual([]);
    expect(sameAsRoom(person, room)).toBe(true);
    expect(editorXml(person)).toBe(listXml(person));
    expect(editorXml(person)).toBe(before.xml);
    expect(await savedBody(person)).toBe(before.saved);
  }
  expect(room.room.order).toBe(before.order);
  expect(
    [...room.room.document.suggestions.values()].filter((s) => s.status === 'pending')
  ).toEqual([]);
}

async function stateOf(room: OrderedLocalRoom, person: Person) {
  return { order: room.room.order, xml: editorXml(person), saved: await savedBody(person) };
}

describe('a format the words already have changes nothing in a room', () => {
  for (const who of ['Olive', 'Ana'] as const) {
    test(`${who}: the same size over part of a run`, async () => {
      const { room, ana, olive } = await people();
      const person = who === 'Olive' ? olive : ana;
      const before = await stateOf(room, person);
      caretAt(person, 0, 4, 7);
      surface(person).setRunProperty('sz', { val: '28' });
      await settle();
      await unchanged(room, [ana, olive], before);
    });

    test(`${who}: bold and size over the whole paragraph, mark included`, async () => {
      const { room, ana, olive } = await people();
      const person = who === 'Olive' ? olive : ana;
      const before = await stateOf(room, person);
      caretAt(person, 0, 0, 19);
      surface(person).setRunProperty('sz', { val: '28' });
      surface(person).setRunProperty('b');
      await settle();
      await unchanged(room, [ana, olive], before);
    });

    test(`${who}: the alignment the paragraph already has`, async () => {
      const { room, ana, olive } = await people();
      const person = who === 'Olive' ? olive : ana;
      const before = await stateOf(room, person);
      caretAt(person, 0, 3);
      surface(person).setParagraphProperty('jc', { val: 'center' });
      await settle();
      await unchanged(room, [ana, olive], before);
    });
  }

  test('a new size still reaches every copy', async () => {
    const { room, ana, olive } = await people();
    caretAt(olive, 0, 4, 7);
    surface(olive).setRunProperty('sz', { val: '32' });
    await settle();
    for (const person of [ana, olive]) {
      expect(person.refusals).toEqual([]);
      expect(sameAsRoom(person, room)).toBe(true);
      expect(editorXml(person)).toContain('<w:sz w:val="32"/>');
    }
  });
});
