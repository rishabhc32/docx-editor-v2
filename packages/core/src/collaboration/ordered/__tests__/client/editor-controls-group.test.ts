// A group control (`w:group`) with real editors and one ordered room. Word refuses every edit of
// what a group holds outside its nested controls (desktop Word for Mac, 2026-11-26): a deletion
// of pending words there keeps the words, the revision and the group. The organiser and a party
// meet the same refusal by name (`locked`), and every copy stays the room's. The control nested
// in the group stays editable. The editor once took the organiser's deletion of Ana's words and
// the emptied group was gone.

import { afterEach, describe, expect, test } from 'bun:test';
import { serializeOoxmlPart } from '../../../../store/package/ooxml-tree.ts';
import { openOrderedWordFile } from '../../client/open.ts';
import { composeOrderedPart } from '../../compose/compose.ts';
import type { OrderedDocument } from '../../model/types.ts';
import { docx, p, t } from '../word-fixtures.ts';
import {
  BODY,
  caretAt,
  closeAll,
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

const sdt = (props: string, inner: string) =>
  `<w:sdt><w:sdtPr>${props}</w:sdtPr><w:sdtContent>${inner}</w:sdtContent></w:sdt>`;

/**
 * "A: " + an outer control > a group > Ana's pending "xy", " ", a nested control "ab", then
 * " end.": the group's own content is 3 to 6, the nested control 6 to 8.
 */
const FILE = docx(
  p(
    t('A: ') +
      sdt(
        '<w:tag w:val="outer"/>',
        sdt(
          '<w:tag w:val="group"/><w:group/>',
          '<w:ins w:id="5" w:author="Ana" w:date="2026-11-26T00:00:00Z">' +
            `${t('xy')}</w:ins>${t(' ')}${sdt('<w:tag w:val="n1"/>', t('ab'))}`
        )
      ) +
      t(' end.')
  ) + p(t('Signed.'))
);

/** The group as the room composes it, without ids and dates. */
const groupOf = (doc: OrderedDocument): string =>
  (
    /<w:sdt><w:sdtPr><w:tag w:val="group"\/>.*?<w:tag w:val="n1"\/>.*?<\/w:sdt><\/w:sdtContent><\/w:sdt>/s.exec(
      serializeOoxmlPart(composeOrderedPart(doc, BODY).part)
    )?.[0] ?? '<none>'
  ).replace(/ w:(date|id)="[^"]*"/g, '');

async function seats() {
  const room = localRoom(FILE);
  const olive = await join(room, 'Olive', { mode: 'edit', login: ORGANISER, wordFile: FILE });
  const bob = await join(room, 'Bob', { wordFile: FILE });
  return { room, olive, bob };
}

function converged(room: ReturnType<typeof localRoom>, ...people: Person[]): void {
  for (const person of people) {
    expect({ person: person.name, room: sameAsRoom(person, room) }).toEqual({
      person: person.name,
      room: true,
    });
    expect({ person: person.name, screen: screenFollowsList(person) }).toEqual({
      person: person.name,
      screen: true,
    });
  }
}

describe("a group control's own content in an ordered room", () => {
  for (const who of ['Olive', 'Bob'] as const) {
    test(`${who}'s deletion of Ana's words in the group is refused by name`, async () => {
      const { room, olive, bob } = await seats();
      const before = groupOf(room.room.document);
      expect(before).toContain('xy');
      const person = who === 'Olive' ? olive : bob;
      caretAt(person, 0, 3, 5);
      surface(person).deleteBackward();
      await settle();
      expect(person.refusals.map((refusal) => refusal.reason)).toEqual(['locked']);
      expect(room.room.order).toBe(0);
      converged(room, olive, bob);
      expect(groupOf(room.room.document)).toBe(before);
    });

    test(`${who}'s typing in the group's own content is refused by name`, async () => {
      const { room, olive, bob } = await seats();
      const person = who === 'Olive' ? olive : bob;
      caretAt(person, 0, 4);
      surface(person).insertPlainText('Q');
      await settle();
      expect(person.refusals.map((refusal) => refusal.reason)).toEqual(['locked']);
      expect(room.room.order).toBe(0);
      converged(room, olive, bob);
    });
  }

  test("a party's words in the nested control come and go; the group stays", async () => {
    const { room, olive, bob } = await seats();
    caretAt(bob, 0, 7);
    surface(bob).insertPlainText('Q');
    await settle();
    expect(bob.refusals).toEqual([]);
    converged(room, olive, bob);
    expect(groupOf(room.room.document)).toContain('Q');

    caretAt(bob, 0, 7, 8);
    surface(bob).deleteBackward();
    await settle();
    expect(bob.refusals).toEqual([]);
    converged(room, olive, bob);
    const kept = groupOf(room.room.document);
    expect(kept).toContain('<w:group/>');
    expect(kept).not.toContain('Q');

    // Saved and reopened, the group and Ana's words are there
    const saved = new Uint8Array(await bob.editor.save());
    expect(groupOf(openOrderedWordFile(saved, { file: 2 }).document)).toBe(kept);
  });
});
