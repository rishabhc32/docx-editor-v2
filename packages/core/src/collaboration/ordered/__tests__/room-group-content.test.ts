// A group control's own content in a room (security): Word refuses every edit of what a group
// holds outside its nested controls, and the controls nested in it stay editable. The room
// refuses the same edits by the engine's name (`locked`), for the organiser and a party alike,
// whatever a client sends. The engine's side is `store/__tests__/content-control-group-content`.

import { describe, expect, test } from 'bun:test';
import { at } from './change-apply-fixture.ts';
import { bodyFingerprint } from './room-harness.ts';
import { BOTH, edit, roomOn, sdt, type Who } from './room-locked-fixture.ts';
import { docx, p, t } from './word-fixtures.ts';

const GROUP = '<w:tag w:val="group"/><w:group/>';

/** "A: " + a group > "G " + a nested control "ab" + " end.": own 3 to 5, nested 5 to 7. */
const INLINE = docx(
  p(t('A: ') + sdt(GROUP, t('G ') + sdt('<w:tag w:val="n1"/>', t('ab'))) + t(' end.')) +
    p(t('After.'))
);

/** A block group: "Own words.", a nested block control "Field.", then "After.". */
const BLOCK = docx(
  sdt(GROUP, p(t('Own words.')) + sdt('<w:tag w:val="n2"/>', p(t('Field.')))) + p(t('After.'))
);

type Make = (doc: Parameters<typeof at>[0], who: Who) => Record<string, unknown>;

const insert =
  (line: number, offset: number): Make =>
  (doc, who) => ({ kind: 'insert', ...edit(who), place: at(doc, line, offset), text: 'Q' });
const remove =
  (line: number, from: number, to: number): Make =>
  (doc, who) => ({
    kind: 'delete',
    ...edit(who),
    from: at(doc, line, from),
    to: at(doc, line, to),
  });
const split =
  (line: number, offset: number): Make =>
  (doc, who) => ({ kind: 'split', ...edit(who), place: at(doc, line, offset) });
const bold =
  (line: number, from: number, to: number): Make =>
  (doc, who) => ({
    kind: 'format',
    ...edit(who),
    from: at(doc, line, from),
    to: at(doc, line, to),
    properties: [{ localName: 'b' }],
  });

/** Each case's verdict for the organiser and a party, and whether a refusal changed the body. */
function verdicts(file: Uint8Array, cases: Readonly<Record<string, Make>>): string[] {
  const out: string[] = [];
  for (const [name, make] of Object.entries(cases)) {
    for (const who of BOTH) {
      const { doc, send } = roomOn(file);
      const before = bodyFingerprint(doc);
      const verdict = send(who, make(doc, who));
      const changed = verdict !== 'accepted' && bodyFingerprint(doc) !== before;
      out.push(`${name} ${who}: ${verdict}${changed ? ' (changed)' : ''}`);
    }
  }
  return out;
}

const all = (cases: Readonly<Record<string, Make>>, verdict: string) =>
  Object.keys(cases).flatMap((name) => BOTH.map((who) => `${name} ${who}: ${verdict}`));

describe("a room refuses an edit of a group control's own content (Word)", () => {
  const inline = {
    typing: insert(0, 4),
    'a deletion': remove(0, 3, 5),
    'a deletion into the nested control': remove(0, 4, 6),
    'a deletion of the group and words around it': remove(0, 1, 9),
    Enter: split(0, 4),
    'a format': bold(0, 3, 5),
  };
  test('an inline group', () => {
    expect(verdicts(INLINE, inline)).toEqual(all(inline, 'apply-refused/locked'));
  });

  const block = {
    typing: insert(0, 3),
    'a deletion': remove(0, 0, 3),
    Enter: split(0, 3),
    'a format': bold(0, 0, 3),
  };
  test('a block group', () => {
    expect(verdicts(BLOCK, block)).toEqual(all(block, 'apply-refused/locked'));
  });
});

describe('the controls nested in a group, and the words around it, stay editable', () => {
  const inline = {
    typing: insert(0, 6),
    'a deletion': remove(0, 5, 7),
    'typing outside': insert(0, 1),
  };
  test('an inline group', () => {
    expect(verdicts(INLINE, inline)).toEqual(all(inline, 'accepted'));
  });

  const block = { typing: insert(1, 3), 'a deletion': remove(1, 0, 3) };
  test('a block group', () => {
    expect(verdicts(BLOCK, block)).toEqual(all(block, 'accepted'));
  });
});
